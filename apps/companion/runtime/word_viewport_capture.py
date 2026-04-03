import ctypes as C
import json
import re
import sys
import time
from ctypes import wintypes
from pathlib import Path

from PIL import Image
from zbl import Capture


user32 = C.windll.user32


class RECT(C.Structure):
    _fields_ = [
        ("left", wintypes.LONG),
        ("top", wintypes.LONG),
        ("right", wintypes.LONG),
        ("bottom", wintypes.LONG),
    ]


EnumWindowsProc = C.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


def normalize_text(value):
    return re.sub(r"\s+", " ", (value or "")).strip().lower()


def get_window_text(hwnd):
    length = user32.GetWindowTextLengthW(hwnd)
    buffer = C.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buffer, length + 1)
    return buffer.value


def get_class_name(hwnd):
    buffer = C.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buffer, 256)
    return buffer.value


def get_window_rect(hwnd):
    rect = RECT()
    if not user32.GetWindowRect(hwnd, C.byref(rect)):
        raise OSError(f"GetWindowRect failed for {hwnd}.")
    return {
        "left": int(rect.left),
        "top": int(rect.top),
        "right": int(rect.right),
        "bottom": int(rect.bottom),
        "width": int(rect.right - rect.left),
        "height": int(rect.bottom - rect.top),
    }


def enumerate_child_viewports(parent_hwnd):
    children = []

    @EnumWindowsProc
    def callback(hwnd, _lparam):
        if not user32.IsWindowVisible(hwnd):
            return True
        class_name = get_class_name(hwnd)
        if class_name != "_WwG":
            return True
        rect = get_window_rect(hwnd)
        if rect["width"] <= 0 or rect["height"] <= 0:
            return True
        children.append(
            {
                "handle": int(hwnd),
                "className": class_name,
                "rect": rect,
            }
        )
        return True

    user32.EnumChildWindows(parent_hwnd, callback, 0)
    children.sort(key=lambda child: child["rect"]["width"] * child["rect"]["height"], reverse=True)
    return children


def score_title_match(window_title, expected):
    if not expected:
        return 0

    title_norm = normalize_text(window_title)
    expected_norm = normalize_text(expected)
    if not title_norm or not expected_norm:
        return 0
    if title_norm == expected_norm:
        return 1000
    if expected_norm in title_norm or title_norm in expected_norm:
        return 800

    expected_stem = normalize_text(Path(expected).stem)
    if expected_stem and expected_stem in title_norm:
        return 600
    return 0


def enumerate_word_windows():
    windows = []
    foreground = int(user32.GetForegroundWindow() or 0)

    @EnumWindowsProc
    def callback(hwnd, _lparam):
        if not user32.IsWindowVisible(hwnd):
            return True
        class_name = get_class_name(hwnd)
        if class_name != "OpusApp":
            return True
        rect = get_window_rect(hwnd)
        if rect["width"] <= 0 or rect["height"] <= 0:
            return True
        children = enumerate_child_viewports(hwnd)
        if not children:
            return True
        windows.append(
            {
                "handle": int(hwnd),
                "className": class_name,
                "title": get_window_text(hwnd),
                "rect": rect,
                "viewport": children[0],
                "isForeground": int(hwnd) == foreground,
            }
        )
        return True

    user32.EnumWindows(callback, 0)
    return windows


def choose_window(expected_title, expected_caption):
    candidates = []
    for window in enumerate_word_windows():
        score = 0
        matched_by = None
        caption_score = score_title_match(window["title"], expected_caption)
        title_score = score_title_match(window["title"], expected_title)
        if caption_score:
            score += caption_score
            matched_by = "windowCaption"
        if title_score:
            score += title_score
            matched_by = matched_by or "documentTitle"
        if window["isForeground"]:
            score += 200
        score += min((window["viewport"]["rect"]["width"] * window["viewport"]["rect"]["height"]) // 20000, 120)
        candidates.append((score, matched_by, window))

    if not candidates:
        raise RuntimeError("No visible Word document window with a document viewport was found.")

    candidates.sort(key=lambda item: item[0], reverse=True)
    score, matched_by, window = candidates[0]
    if score <= 0:
        raise RuntimeError("Unable to match the active Word document window.")
    return matched_by or "bestVisibleViewport", window


def grab_window_frame(window_handle):
    last_error = None
    try:
        with Capture(window_handle=int(window_handle)) as capture:
            frame = None
            deadline = time.monotonic() + 5.0
            while frame is None and time.monotonic() < deadline:
                try:
                    frame = capture.try_grab()
                except StopIteration as exc:
                    raise RuntimeError("Viewport capture stopped before a frame became available.") from exc
                if frame is None:
                    time.sleep(0.03)
            if frame is None:
                frame = capture.grab()
            return hex(window_handle), frame
    except Exception as exc:
        last_error = exc

    raise RuntimeError(f"Windows.Graphics.Capture failed for Word window {window_handle}: {last_error}")


def crop_viewport(frame, window_rect, viewport_rect):
    frame_height = int(frame.shape[0])
    frame_width = int(frame.shape[1])
    window_width = max(1, int(window_rect["width"]))
    window_height = max(1, int(window_rect["height"]))
    scale_x = frame_width / window_width
    scale_y = frame_height / window_height

    left = int(round((viewport_rect["left"] - window_rect["left"]) * scale_x))
    top = int(round((viewport_rect["top"] - window_rect["top"]) * scale_y))
    right = int(round((viewport_rect["right"] - window_rect["left"]) * scale_x))
    bottom = int(round((viewport_rect["bottom"] - window_rect["top"]) * scale_y))

    left = max(0, min(frame_width - 1, left))
    top = max(0, min(frame_height - 1, top))
    right = max(left + 1, min(frame_width, right))
    bottom = max(top + 1, min(frame_height, bottom))
    cropped = frame[top:bottom, left:right, :]
    return cropped, scale_x, scale_y


def save_png(path, frame):
    rgba = frame[:, :, [2, 1, 0, 3]].copy()
    Image.fromarray(rgba, "RGBA").save(path)


def main():
    request = json.load(sys.stdin)
    output_dir = Path(request["outputDir"])
    output_dir.mkdir(parents=True, exist_ok=True)

    matched_by, window = choose_window(request.get("documentTitle"), request.get("windowCaption"))
    capture_handle, frame = grab_window_frame(window["handle"])
    viewport_frame, scale_x, scale_y = crop_viewport(frame, window["rect"], window["viewport"]["rect"])

    viewport_path = output_dir / "viewport.png"
    save_png(viewport_path, viewport_frame)

    window_path = None
    if request.get("includeWindowFrame"):
        window_path = output_dir / "window.png"
        save_png(window_path, frame)

    response = {
        "matchedBy": matched_by,
        "captureHandle": capture_handle,
        "window": {
            "handle": hex(window["handle"]),
            "className": window["className"],
            "title": window["title"],
            "rect": window["rect"],
        },
        "viewport": {
            "handle": hex(window["viewport"]["handle"]),
            "className": window["viewport"]["className"],
            "rect": window["viewport"]["rect"],
        },
        "frame": {
            "width": int(frame.shape[1]),
            "height": int(frame.shape[0]),
            "scaleX": round(scale_x, 4),
            "scaleY": round(scale_y, 4),
        },
        "viewportImagePath": str(viewport_path),
        "windowImagePath": str(window_path) if window_path else None,
    }
    print(json.dumps(response))


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(str(error), file=sys.stderr)
        sys.exit(1)
