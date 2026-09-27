import assert from "node:assert/strict";
import test from "node:test";
import { OfficeBridgeState } from "../src/bridge-state.mjs";

const identity = {
  sessionId: "session-1",
  workbookUrl: "file:///tmp/book.xlsx",
  instanceId: "instance-1"
};

function activate(state, overrides = {}) {
  state.registerSession(identity);
  return state.activate({
    ...identity,
    protocolVersion: 1,
    host: "Excel",
    platform: "Mac",
    officeVersion: "16.112",
    requirementSets: ["1.1", "1.16"],
    ...overrides
  });
}

test("negotiates runtime capability metadata without enabling feature actions", () => {
  const state = new OfficeBridgeState();
  const session = activate(state);
  assert.equal(session.available, true);
  assert.deepEqual(session.runtime.requirementSets, ["1.1", "1.16"]);
  assert.deepEqual(session.enabledActions, ["bridge.health"]);
  assert.throws(() => state.createRequest({
    sessionId: identity.sessionId,
    action: "table.create",
    timeoutMs: 1000
  }), /not enabled/);
  assert.throws(() => state.createRequest({
    sessionId: identity.sessionId,
    action: "bridge.health",
    timeoutMs: 1000,
    payload: { unexpected: true }
  }), /does not accept/);
});

test("binds Office polling and results to exact workbook and instance", () => {
  const state = new OfficeBridgeState();
  activate(state);
  const request = state.createRequest({
    sessionId: identity.sessionId,
    action: "bridge.health",
    timeoutMs: 1000
  });
  assert.throws(() => state.takeNext({
    ...identity,
    workbookUrl: "file:///tmp/other.xlsx"
  }), /exact workbook/);
  assert.equal(state.takeNext(identity).requestId, request.requestId);
  assert.throws(() => state.complete({
    ...identity,
    instanceId: "other-instance",
    requestId: request.requestId,
    success: true
  }), /exact workbook/);
});

test("serializes requests per session and correlates results", () => {
  const state = new OfficeBridgeState();
  activate(state);
  const first = state.createRequest({
    sessionId: identity.sessionId,
    action: "bridge.health",
    timeoutMs: 1000
  });
  const second = state.createRequest({
    sessionId: identity.sessionId,
    action: "bridge.health",
    timeoutMs: 1000
  });
  assert.equal(state.takeNext(identity).requestId, first.requestId);
  assert.equal(state.takeNext(identity), null);
  state.complete({ ...identity, requestId: first.requestId, success: true });
  assert.equal(state.takeNext(identity).requestId, second.requestId);
});

test("expires deadlines and honors cancellation", () => {
  let now = 100;
  const state = new OfficeBridgeState(() => now);
  activate(state);
  const expired = state.createRequest({
    sessionId: identity.sessionId,
    action: "bridge.health",
    timeoutMs: 10
  });

  now = 111;
  assert.equal(state.getRequest({
    ...identity,
    requestId: expired.requestId
  }).status, "expired");
  assert.equal(state.takeNext(identity), null);

  const cancelled = state.createRequest({
    sessionId: identity.sessionId,
    action: "bridge.health",
    timeoutMs: 10
  });
  assert.equal(state.cancel({ ...identity, requestId: cancelled.requestId }), true);
  assert.equal(state.getRequest({
    ...identity,
    requestId: cancelled.requestId
  }).status, "cancelled");
});

test("teardown releases the workbook and removes all session requests", () => {
  const state = new OfficeBridgeState();
  activate(state);
  const completed = state.createRequest({
    sessionId: identity.sessionId,
    action: "bridge.health",
    timeoutMs: 1000
  });
  state.takeNext(identity);
  state.complete({ ...identity, requestId: completed.requestId, success: true });
  const outstanding = state.createRequest({
    sessionId: identity.sessionId,
    action: "bridge.health",
    timeoutMs: 1000
  });

  assert.throws(() => state.unregisterSession({
    sessionId: identity.sessionId,
    workbookUrl: "file:///tmp/other.xlsx"
  }), /exact workbook/);
  const teardown = state.unregisterSession(identity);

  assert.equal(teardown.terminalizedRequests, 1);
  assert.equal(teardown.removedRequests, 2);
  assert.throws(() => state.getRequest({
    ...identity,
    requestId: outstanding.requestId
  }), /not registered/);
  assert.doesNotThrow(() => state.registerSession({
    sessionId: "session-2",
    workbookUrl: identity.workbookUrl
  }));
});

test("prevents session identifiers from being rebound to another workbook", () => {
  const state = new OfficeBridgeState();
  state.registerSession(identity);
  assert.throws(() => state.registerSession({
    sessionId: identity.sessionId,
    workbookUrl: "file:///tmp/other.xlsx"
  }), /cannot be rebound/);
});
