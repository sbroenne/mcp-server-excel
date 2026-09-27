import assert from "node:assert/strict";
import http from "node:http";
import test from "node:test";
import { OfficeBridgeState } from "../src/bridge-state.mjs";
import { createRequestHandler } from "../src/server.mjs";

const config = {
  port: 47132,
  origin: "https://localhost:47132",
  token: "t".repeat(43)
};

test("HTTP protocol rejects unauthenticated traffic and correlates a health request", async (context) => {
  let now = 100;
  const state = new OfficeBridgeState(() => now);
  const server = http.createServer(createRequestHandler(config, state));
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  context.after(() => new Promise((resolve) => server.close(resolve)));
  const port = server.address().port;

  const denied = await request(port, "GET", "/v1/health", {
    Host: `localhost:${config.port}`
  });
  assert.equal(denied.statusCode, 401);

  const headers = {
    Host: `localhost:${config.port}`,
    Authorization: `Bearer ${config.token}`,
    "Content-Type": "application/json"
  };
  const registered = await request(port, "POST", "/v1/sessions", headers, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx"
  });
  assert.equal(registered.statusCode, 201);

  const activated = await request(port, "POST", "/v1/office/activate", {
    ...headers,
    Origin: config.origin
  }, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    instanceId: "instance-1",
    protocolVersion: 1,
    host: "Excel",
    platform: "Mac",
    officeVersion: "16.112",
    requirementSets: ["1.1", "1.16"]
  });
  assert.equal(activated.statusCode, 200);

  const created = await request(port, "POST", "/v1/requests", headers, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    action: "bridge.health",
    timeoutMs: 1000
  });
  assert.equal(created.statusCode, 202);

  const next = await request(port, "POST", "/v1/office/next", {
    ...headers,
    Origin: config.origin
  }, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    instanceId: "instance-1"
  });
  assert.equal(next.statusCode, 200);
  assert.equal(next.body.requestId, created.body.requestId);

  const completed = await request(port, "POST", "/v1/office/results", {
    ...headers,
    Origin: config.origin
  }, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    instanceId: "instance-1",
    requestId: created.body.requestId,
    success: true,
    value: { requirementSets: ["1.1", "1.16"] }
  });
  assert.equal(completed.statusCode, 200);
  assert.equal(completed.body.result.errorMessage, null);

  const observed = await request(port, "POST", "/v1/requests/status", headers, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    requestId: created.body.requestId
  });
  assert.equal(observed.statusCode, 200);
  assert.equal(observed.body.status, "completed");
  assert.deepEqual(observed.body.result.value, { requirementSets: ["1.1", "1.16"] });

  const wrongWorkbook = await request(port, "POST", "/v1/requests/status", headers, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/other.xlsx",
    requestId: created.body.requestId
  });
  assert.equal(wrongWorkbook.statusCode, 409);

  const failed = await request(port, "POST", "/v1/requests", headers, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    action: "bridge.health",
    timeoutMs: 1000
  });
  await request(port, "POST", "/v1/office/next", {
    ...headers,
    Origin: config.origin
  }, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    instanceId: "instance-1"
  });
  await request(port, "POST", "/v1/office/results", {
    ...headers,
    Origin: config.origin
  }, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    instanceId: "instance-1",
    requestId: failed.body.requestId,
    success: false,
    errorMessage: "Office.js health probe failed."
  });
  const observedFailure = await request(port, "POST", "/v1/requests/status", headers, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    requestId: failed.body.requestId
  });
  assert.equal(observedFailure.body.status, "failed");
  assert.equal(observedFailure.body.result.errorMessage, "Office.js health probe failed.");

  const expiring = await request(port, "POST", "/v1/requests", headers, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    action: "bridge.health",
    timeoutMs: 10
  });
  now = 111;
  const observedExpiry = await request(port, "POST", "/v1/requests/status", headers, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    requestId: expiring.body.requestId
  });
  assert.equal(observedExpiry.body.status, "expired");

  const cancelling = await request(port, "POST", "/v1/requests", headers, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    action: "bridge.health",
    timeoutMs: 1000
  });
  await request(port, "POST", "/v1/office/next", {
    ...headers,
    Origin: config.origin
  }, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    instanceId: "instance-1"
  });
  const cancelled = await request(port, "POST", "/v1/requests/cancel", headers, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    requestId: cancelling.body.requestId
  });
  assert.equal(cancelled.body.cancelled, true);
  assert.equal(cancelled.body.dispatched, true);
  const observedCancellation = await request(port, "POST", "/v1/requests/status", headers, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    requestId: cancelling.body.requestId
  });
  assert.equal(observedCancellation.body.status, "cancelled");

  const misboundLateResult = await request(port, "POST", "/v1/office/results", {
    ...headers,
    Origin: config.origin
  }, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/other.xlsx",
    instanceId: "instance-1",
    requestId: cancelling.body.requestId,
    success: true
  });
  assert.equal(misboundLateResult.statusCode, 409);

  const lateResult = await request(port, "POST", "/v1/office/results", {
    ...headers,
    Origin: config.origin
  }, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    instanceId: "instance-1",
    requestId: cancelling.body.requestId,
    success: true,
    value: { late: true }
  });
  assert.equal(lateResult.statusCode, 200);
  assert.equal(lateResult.body.lateResultIgnored, true);

  const afterLate = await request(port, "POST", "/v1/requests", headers, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    action: "bridge.health",
    timeoutMs: 1000
  });
  const afterLateNext = await request(port, "POST", "/v1/office/next", {
    ...headers,
    Origin: config.origin
  }, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    instanceId: "instance-1"
  });
  assert.equal(afterLateNext.body.requestId, afterLate.body.requestId);
  const afterLateCompleted = await request(port, "POST", "/v1/office/results", {
    ...headers,
    Origin: config.origin
  }, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx",
    instanceId: "instance-1",
    requestId: afterLate.body.requestId,
    success: true
  });
  assert.equal(afterLateCompleted.body.status, "completed");

  const closed = await request(port, "POST", "/v1/sessions/close", headers, {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/book.xlsx"
  });
  assert.equal(closed.statusCode, 200);
  assert.equal(closed.body.removedRequests, 5);

  const reopened = await request(port, "POST", "/v1/sessions", headers, {
    sessionId: "session-2",
    workbookUrl: "file:///tmp/book.xlsx"
  });
  assert.equal(reopened.statusCode, 201);
});

function request(port, method, requestPath, headers = {}, body = null) {
  return new Promise((resolve, reject) => {
    const serialized = body === null ? null : JSON.stringify(body);
    const outgoing = http.request({
      host: "127.0.0.1",
      port,
      method,
      path: requestPath,
      headers: serialized === null
        ? headers
        : { ...headers, "Content-Length": Buffer.byteLength(serialized) }
    }, (response) => {
      const chunks = [];
      response.on("data", (chunk) => chunks.push(chunk));
      response.on("end", () => {
        const text = Buffer.concat(chunks).toString("utf8");
        resolve({
          statusCode: response.statusCode,
          body: text ? JSON.parse(text) : null
        });
      });
    });
    outgoing.on("error", reject);
    if (serialized !== null) {
      outgoing.write(serialized);
    }
    outgoing.end();
  });
}
