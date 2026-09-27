import { readFile } from "node:fs/promises";
import https from "node:https";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { MAX_BODY_BYTES, PROTOCOL_VERSION } from "./constants.mjs";
import { OfficeBridgeState } from "./bridge-state.mjs";
import { authorizeRequest, BridgeHttpError } from "./security.mjs";

const directory = path.dirname(fileURLToPath(import.meta.url));
const icon = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
  "base64"
);

export async function startServer(config) {
  const state = new OfficeBridgeState();
  const [key, cert] = await Promise.all([
    readFile(config.privateKeyPath),
    readFile(config.certificatePath)
  ]);
  const server = https.createServer({ key, cert }, createRequestHandler(config, state));
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(config.port, "127.0.0.1", resolve);
  });
  return server;
}

export function createRequestHandler(config, state = new OfficeBridgeState()) {
  return (request, response) => {
    void route(request, response, config, state);
  };
}

async function route(request, response, config, state) {
  try {
    const url = new URL(request.url, config.origin);
    if (request.method === "GET" && url.pathname === "/taskpane.html") {
      return send(response, 200, await readFile(path.join(directory, "taskpane.html"), "utf8"), "text/html");
    }
    if (request.method === "GET" && url.pathname === "/addin.mjs") {
      return send(response, 200, await readFile(path.join(directory, "addin.mjs"), "utf8"), "text/javascript");
    }
    if (request.method === "GET" && url.pathname === "/constants.mjs") {
      return send(response, 200, await readFile(path.join(directory, "constants.mjs"), "utf8"), "text/javascript");
    }
    if (request.method === "GET" && url.pathname.startsWith("/icon-")) {
      return send(response, 200, icon, "image/png");
    }

    authorizeRequest(request, config);
    if (request.method === "GET" && url.pathname === "/v1/health") {
      return json(response, 200, {
        status: "running",
        protocolVersion: PROTOCOL_VERSION,
        enabledActions: ["bridge.health"]
      });
    }

    const body = request.method === "POST" ? await readJson(request) : null;
    if (request.method === "POST" && url.pathname === "/v1/sessions") {
      return json(response, 201, state.registerSession(body));
    }
    if (request.method === "POST" && url.pathname === "/v1/office/activate") {
      return json(response, 200, state.activate(body));
    }
    if (request.method === "POST" && url.pathname === "/v1/requests") {
      return json(response, 202, state.createRequest(body));
    }
    if (request.method === "POST" && url.pathname === "/v1/office/next") {
      return json(response, 200, state.takeNext(body));
    }
    if (request.method === "POST" && url.pathname === "/v1/office/results") {
      return json(response, 200, state.complete(body));
    }
    if (request.method === "POST" && url.pathname === "/v1/requests/cancel") {
      return json(response, 200, { cancelled: state.cancel(body.requestId) });
    }
    throw new BridgeHttpError(404, "Route not found.");
  } catch (error) {
    const statusCode = error.statusCode ?? (error instanceof TypeError ? 400 : 409);
    json(response, statusCode, { success: false, errorMessage: error.message });
  }
}

async function readJson(request) {
  const chunks = [];
  let length = 0;
  for await (const chunk of request) {
    length += chunk.length;
    if (length > MAX_BODY_BYTES) {
      throw new BridgeHttpError(413, "Request body is too large.");
    }
    chunks.push(chunk);
  }
  try {
    const value = JSON.parse(Buffer.concat(chunks).toString("utf8"));
    if (!value || Array.isArray(value) || typeof value !== "object") {
      throw new Error();
    }
    return value;
  } catch {
    throw new BridgeHttpError(400, "Request body must be a JSON object.");
  }
}

function json(response, statusCode, value) {
  send(response, statusCode, JSON.stringify(value), "application/json");
}

function send(response, statusCode, body, contentType = "text/plain") {
  response.writeHead(statusCode, {
    "Content-Type": `${contentType}; charset=utf-8`,
    "Cache-Control": "no-store",
    "X-Content-Type-Options": "nosniff"
  });
  response.end(body);
}
