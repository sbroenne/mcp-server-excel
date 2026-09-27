import assert from "node:assert/strict";
import test from "node:test";
import { authorizeRequest } from "../src/security.mjs";

const config = {
  port: 47132,
  origin: "https://localhost:47132",
  token: "a".repeat(43)
};

test("requires bearer authentication and exact loopback host", () => {
  assert.throws(() => authorizeRequest({ headers: {
    host: "localhost:47132",
    authorization: "Bearer wrong"
  } }, config), /authentication/);
  assert.throws(() => authorizeRequest({ headers: {
    host: "evil.example",
    authorization: `Bearer ${config.token}`
  } }, config), /Host/);
});

test("accepts native no-origin calls and rejects unexpected browser origins", () => {
  assert.doesNotThrow(() => authorizeRequest({ headers: {
    host: "127.0.0.1:47132",
    authorization: `Bearer ${config.token}`
  } }, config));
  assert.throws(() => authorizeRequest({ headers: {
    host: "localhost:47132",
    origin: "https://evil.example",
    authorization: `Bearer ${config.token}`
  } }, config), /Origin/);
});
