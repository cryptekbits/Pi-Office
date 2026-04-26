import https from "node:https";

const requiredUrls = [
  "https://localhost:3443/",
  "https://localhost:3443/shortcuts.json",
  "https://localhost:3443/brand/icon-16.png?v=20260421",
  "https://localhost:3443/brand/icon-32.png?v=20260421",
  "https://localhost:3443/brand/icon-80.png?v=20260421",
];

const timeoutMs = 3000;

function probe(url) {
  return new Promise((resolve, reject) => {
    const request = https.request(
      url,
      {
        method: "GET",
        rejectUnauthorized: false,
        timeout: timeoutMs,
      },
      (response) => {
        response.resume();
        response.on("end", () => {
          const statusCode = response.statusCode ?? 0;
          if (statusCode >= 200 && statusCode < 400) {
            resolve();
            return;
          }
          reject(new Error(`${url} returned HTTP ${statusCode}`));
        });
      },
    );

    request.on("timeout", () => {
      request.destroy(new Error(`${url} timed out after ${timeoutMs}ms`));
    });
    request.on("error", reject);
    request.end();
  });
}

async function main() {
  const failures = [];

  for (const url of requiredUrls) {
    try {
      await probe(url);
    } catch (error) {
      failures.push(error instanceof Error ? error.message : String(error));
    }
  }

  if (failures.length > 0) {
    console.error("[sideload preflight] The taskpane dev host is not ready.");
    console.error("[sideload preflight] Start it in another terminal with: npm run dev");
    console.error("[sideload preflight] The optional companion on https://localhost:3444 is not required for add-in load.");
    for (const failure of failures) {
      console.error(`[sideload preflight] ${failure}`);
    }
    process.exit(1);
  }

  console.log("[sideload preflight] OK (taskpane host resources are reachable on https://localhost:3443).");
}

await main();
