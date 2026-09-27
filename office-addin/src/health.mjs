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
    let settled = false;
    const settle = (callback, value) => {
      if (settled) {
        return false;
      }
      settled = true;
      clearTimeout(deadline);
      callback(value);
      return true;
    };
    const request = requestFactory(`${config.origin}/v1/health`, {
      ca: certificate,
      headers: { Authorization: `Bearer ${config.token}` }
    }, (response) => {
      const chunks = [];
      response.on("data", (chunk) => chunks.push(chunk));
      response.on("end", () => {
        settle(resolve, {
          ok: response.statusCode >= 200 && response.statusCode < 300,
          body: Buffer.concat(chunks).toString("utf8")
        });
      });
      response.on("aborted", () => {
        settle(reject, new Error("Bridge health response was aborted before completion."));
      });
      response.on("error", (error) => {
        settle(reject, new Error(`Bridge health response failed: ${error.message}`, {
          cause: error
        }));
      });
      response.on("close", () => {
        if (!response.complete) {
          settle(reject, new Error("Bridge health response closed before completion."));
        }
      });
    });
    if (!settled) {
      deadline = setTimeout(() => {
        const error = new Error(
          `Bridge health check did not respond within ${timeoutMs}ms. ` +
          "Confirm that the local bridge is running and the configured port is available."
        );
        settle(reject, error);
        request.destroy(error);
      }, timeoutMs);
    }
    request.on("error", (error) => {
      settle(reject, error);
    });
  });
}
