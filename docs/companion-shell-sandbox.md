# Companion Shell Sandbox Policy

`SECURITY-006` is the gate for any future shell capability. Until this policy is implemented and the destructive probes pass, Pi-Office must not expose raw host bash, PowerShell, command prompt, `edit`, or `write` tools from the taskpane or companion.

## Default State

- Shell is unavailable in Basic taskpane-only mode.
- Shell is unavailable in the optional companion unless the sandbox provider reports a supported, tested isolation backend.
- The taskpane must never construct Pi's default local `BashOperations` against the user's real machine.
- Missing sandbox support must fail closed as `shell unavailable`; it must not fall back to raw host process execution.

## Capability Boundary

Readable roots:

- Only user-selected saved-document folders or explicit workspace roots.
- No implicit user profile, home directory, drive root, system folder, browser profile, Office cache, or repo root access.
- Symlinks, junctions, and shortcuts must resolve inside an approved readable root before access.

Writable roots:

- A companion-created scratch directory only.
- Scratch must be outside the user document/workspace tree unless the user explicitly exports a result.
- No in-place writes, deletes, renames, package installs, commits, pushes, or generated file drops in the readable root.

Denied secrets:

- `.env`, `.env.*`, credential stores, SSH keys, cloud config folders, Office tokens/cache, browser profiles, npm/pip auth files, and common key/cert extensions.
- Denied patterns apply before tool execution and again to any file path surfaced through command output or symlink traversal.

Network:

- Disabled by default.
- No package manager network access by default.
- Any future network mode requires a separate explicit policy, visible user approval, and probe coverage.

## Command Classes

Allowed candidates after sandbox implementation:

- Read-only inspection commands such as `pwd`, `ls`, `find`, `rg`, `grep`, `wc`, `head`, `tail`, `cat` for allowed files, and language/tool version checks.
- Commands must be run with timeouts, output caps, environment scrubbing, and command audit logs.

Denied candidates:

- Mutating filesystem commands outside scratch.
- `git commit`, `git push`, destructive VCS commands, package installs, service/process managers, shell profile edits, registry edits, credential reads, and network clients.
- Nested shells or interpreters that can bypass policy unless wrapped by the same sandbox controls.

## Platform Backends

Windows:

- Prefer WSL2, Docker, or Hyper-V-backed isolation.
- Native Windows process execution remains unavailable until path, ACL, symlink/junction, profile, network, and secret-denial behavior is proven.

Linux:

- Prefer a mount/user/network namespace sandbox such as bubblewrap or a container backend.
- Approved readable roots mount read-only; scratch mounts read-write; network namespace is off by default.

macOS:

- Prefer a containerized or equivalent sandbox backend.
- Native process execution remains unavailable until file, symlink, profile, network, and secret-denial behavior is proven.

## Taskpane/Companion Protocol

The companion-owned shell provider must expose:

- Capability discovery: `unavailable`, `available`, or `degraded`, with backend name and probe timestamp.
- A session-scoped shell policy snapshot: readable roots, scratch root, denied patterns, network policy, timeout, output cap.
- A command request contract with category, command, args, cwd, environment allowlist, and user approval reference.
- A redacted audit result with exit code, capped stdout/stderr, touched paths, denied paths, duration, and policy version.

Office tools still execute only in the taskpane. Shell tools never execute Office.js.

## Required Probes

Before shell can become available, automated probes must prove:

- Writing, deleting, renaming, or chmod/chown outside scratch fails.
- Reading `.env`, SSH keys, cloud credential folders, browser profiles, and Office cache/token paths fails.
- Symlink/junction escape attempts fail.
- Network calls fail by default.
- Package install commands fail by default.
- `git push`, `git commit`, and destructive VCS commands fail by default.
- Output caps and timeouts terminate noisy or long-running commands.
- Environment variables are scrubbed unless explicitly allowed.

## Current Implementation Status

- The active companion exposes read-only file tools and read-only MCP execution only.
- There is no companion shell route and no custom Pi `BashOperations` backend.
- `scripts/office-tests/src/external-context-gaps.test.ts` asserts that raw `bash`, `edit`, and `write` tools are not present in active session tools and that the companion server has no shell/bash/exec route.
