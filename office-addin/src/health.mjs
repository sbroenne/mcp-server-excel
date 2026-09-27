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
    let request;
    let terminalError;
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
    const destroyRequest = (error) => {
      if (request && !request.destroyed) {
        request.destroy(error);
      }
    };
    const fail = (error) => {
      if (settle(reject, error)) {
        terminalError = error;
        destroyRequest(error);
      }
    };
    request = requestFactory(`${config.origin}/v1/health`, {
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
        fail(new Error("Bridge health response was aborted before completion."));
      });
      response.on("error", (error) => {
        fail(new Error(`Bridge health response failed: ${error.message}`, {
          cause: error
        }));
      });
      response.on("close", () => {
        if (!response.complete) {
          fail(new Error("Bridge health response closed before completion."));
        }
      });
    });
    request.on("error", fail);
    if (!settled) {
      deadline = setTimeout(() => {
        fail(new Error(
          `Bridge health check did not respond within ${timeoutMs}ms. ` +
          "Confirm that the local bridge is running and the configured port is available."
        ));
      }, timeoutMs);
    } else {
      destroyRequest(terminalError);
    }
  });
}
