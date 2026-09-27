import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import test from "node:test";
import { requestBridgeHealth } from "../src/health.mjs";

test("health check destroys a stalled request at its deadline", async () => {
  let destroyedError = null;
  const requestFactory = (url, options, callback) => {
    assert.equal(url, "https://localhost:47132/v1/health");
    assert.equal(options.headers.Authorization, "Bearer token");
    assert.equal(typeof callback, "function");
    const request = new EventEmitter();
    request.destroy = (error) => {
      destroyedError = error;
      request.emit("error", error);
    };
    return request;
  };

  await assert.rejects(
    requestBridgeHealth({
      origin: "https://localhost:47132",
      token: "token"
    }, Buffer.from("certificate"), requestFactory, 5),
    /did not respond within 5ms.*local bridge is running/s
  );
  assert.match(destroyedError.message, /configured port/);
});

test("health check rejects a partial response that aborts", async () => {
  const response = new EventEmitter();
  response.statusCode = 200;
  response.complete = false;
  const destroyErrors = [];
  const requestFactory = (url, options, callback) => {
    const request = new EventEmitter();
    request.destroy = (error) => {
      request.destroyed = true;
      destroyErrors.push(error);
      request.emit("error", error);
    };
    setImmediate(() => {
      callback(response);
      response.emit("data", Buffer.from('{"status":'));
      response.emit("aborted");
      response.emit("close");
    });
    return request;
  };

  const startedAt = Date.now();
  await assert.rejects(
    requestBridgeHealth({
      origin: "https://localhost:47132",
      token: "token"
    }, Buffer.from("certificate"), requestFactory, 1000),
    /response was aborted/
  );
  assert.ok(Date.now() - startedAt < 500, "Partial response abort should fail before the deadline.");
  assert.equal(destroyErrors.length, 1);
  assert.match(destroyErrors[0].message, /response was aborted/);
});

test("health check destroys the request when the response emits an error", async () => {
  const response = new EventEmitter();
  response.statusCode = 200;
  response.complete = false;
  const destroyErrors = [];
  const requestFactory = (url, options, callback) => {
    const request = new EventEmitter();
    request.destroy = (error) => {
      request.destroyed = true;
      destroyErrors.push(error);
      request.emit("error", error);
    };
    setImmediate(() => {
      callback(response);
      response.emit("error", new Error("socket read failed"));
      response.emit("close");
    });
    return request;
  };

  await assert.rejects(
    requestBridgeHealth({
      origin: "https://localhost:47132",
      token: "token"
    }, Buffer.from("certificate"), requestFactory, 1000),
    /response failed: socket read failed/
  );
  assert.equal(destroyErrors.length, 1);
  assert.match(destroyErrors[0].message, /response failed/);
});

test("health check destroys the request when the response closes incomplete", async () => {
  const response = new EventEmitter();
  response.statusCode = 200;
  response.complete = false;
  const destroyErrors = [];
  const requestFactory = (url, options, callback) => {
    const request = new EventEmitter();
    request.destroy = (error) => {
      request.destroyed = true;
      destroyErrors.push(error);
      request.emit("error", error);
    };
    setImmediate(() => {
      callback(response);
      response.emit("close");
    });
    return request;
  };

  await assert.rejects(
    requestBridgeHealth({
      origin: "https://localhost:47132",
      token: "token"
    }, Buffer.from("certificate"), requestFactory, 1000),
    /response closed before completion/
  );
  assert.equal(destroyErrors.length, 1);
  assert.match(destroyErrors[0].message, /response closed before completion/);
});
