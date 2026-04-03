import { execFile } from "node:child_process";

export async function openExternal(url: string): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    const child = execFile("cmd", ["/c", "start", "", url], { windowsHide: true }, (error) => {
      if (error) {
        reject(error);
      } else {
        resolve();
      }
    });

    child.unref();
  });
}
