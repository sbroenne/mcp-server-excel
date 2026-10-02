import { timingSafeEqual } from "node:crypto";

function equalSecret(actual, expected) {
  const actualBuffer = Buffer.from(actual ?? "", "utf8");
  const expectedBuffer = Buffer.from(expected, "utf8");
  return actualBuffer.length === expectedBuffer.length
    && timingSafeEqual(actualBuffer, expectedBuffer);
}

export function authorizeRequest(request, config) {
  const host = request.headers.host;
  if (host !== `localhost:${config.port}` && host !== `127.0.0.1:${config.port}`) {
    throw new BridgeHttpError(400, "Invalid loopback Host header.");
  }

  const origin = request.headers.origin;
  if (origin && origin !== config.origin) {
    throw new BridgeHttpError(403, "Origin is not allowed.");
  }

  const authorization = request.headers.authorization;
  const prefix = "Bearer ";
  if (!authorization?.startsWith(prefix)
      || !equalSecret(authorization.slice(prefix.length), config.token)) {
    throw new BridgeHttpError(401, "Bridge authentication failed.");
  }
}

export class BridgeHttpError extends Error {
  constructor(statusCode, message) {
    super(message);
    this.statusCode = statusCode;
  }
}
