import https from "node:https";

export const HEALTH_TIMEOUT_MS = 5_000;

export async function requestBridgeHealth(
  config,
  certificate,
  requestFactory = https.get,
  timeoutMs = HEALTH_TIMEOUT_MS
) {
  return new Promise((resolve, reject) => {
    let deadline;
    const request = requestFactory(`${config.origin}/v1/health`, {
      ca: certificate,
      headers: { Authorization: `Bearer ${config.token}` }
    }, (response) => {
      const chunks = [];
      response.on("data", (chunk) => chunks.push(chunk));
      response.on("end", () => {
        clearTimeout(deadline);
        resolve({
          ok: response.statusCode >= 200 && response.statusCode < 300,
          body: Buffer.concat(chunks).toString("utf8")
        });
      });
    });
    deadline = setTimeout(() => {
      request.destroy(new Error(
        `Bridge health check did not respond within ${timeoutMs}ms. ` +
        "Confirm that the local bridge is running and the configured port is available."
      ));
    }, timeoutMs);
    request.on("error", (error) => {
      clearTimeout(deadline);
      reject(error);
    });
  });
}
