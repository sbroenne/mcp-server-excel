import assert from "node:assert/strict";
import { chmod, mkdtemp, readFile, stat, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import packageJson from "../package.json" with { type: "json" };
import { ADDIN_VERSION } from "../src/constants.mjs";
import { installBridge, removeBridge } from "../src/lifecycle.mjs";

test("package and Office manifest versions remain aligned", () => {
  assert.equal(`${packageJson.version}.0`, ADDIN_VERSION);
});

test("install, upgrade, and removal preserve token and private file permissions", async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), "excelmcp-officejs-"));
  const configDir = path.join(root, "config");
  const wefDir = path.join(root, "wef");
  const certificate = path.join(root, "certificate.pem");
  const privateKey = path.join(root, "private-key.pem");
  await writeFile(certificate, "certificate");
  await writeFile(privateKey, "private-key");
  await chmod(certificate, 0o600);
  await chmod(privateKey, 0o600);

  const first = await installBridge({ configDir, wefDir, certificate, privateKey });
  const firstConfig = JSON.parse(await readFile(first.configPath, "utf8"));
  const second = await installBridge({ configDir, wefDir, certificate, privateKey });
  const secondConfig = JSON.parse(await readFile(second.configPath, "utf8"));
  assert.equal(secondConfig.token, firstConfig.token);
  assert.equal((await stat(second.configPath)).mode & 0o777, 0o600);
  assert.match(await readFile(second.manifestPath, "utf8"), /ExcelApi" MinVersion="1.1"/);

  const unrelated = path.join(configDir, "keep.txt");
  await writeFile(unrelated, "not bridge-owned");
  await removeBridge({ configDir, wefDir });
  await assert.rejects(stat(second.configPath), { code: "ENOENT" });
  await assert.rejects(stat(second.manifestPath), { code: "ENOENT" });
  assert.equal(await readFile(unrelated, "utf8"), "not bridge-owned");
});
