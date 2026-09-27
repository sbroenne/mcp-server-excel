#!/usr/bin/env node
import { readFile } from "node:fs/promises";
import https from "node:https";
import process from "node:process";
import { defaultPaths, installBridge, removeBridge } from "./lifecycle.mjs";
import { startServer } from "./server.mjs";

const [command, ...args] = process.argv.slice(2);
const options = parseOptions(args);
const paths = defaultPaths();
const configDir = options["config-dir"] ?? paths.configDir;
const wefDir = options["wef-dir"] ?? paths.wefDir;

try {
  if (command === "install" || command === "upgrade") {
    const result = await installBridge({
      configDir,
      wefDir,
      certificate: options.certificate,
      privateKey: options["private-key"],
      port: options.port ? Number(options.port) : undefined
    });
    console.log(`Installed manifest: ${result.manifestPath}`);
    console.log(`Installed configuration: ${result.configPath}`);
  } else if (command === "remove") {
    await removeBridge({ configDir, wefDir });
    console.log("Removed the ExcelMcp Office.js bridge files. Restart Excel to unload the add-in.");
  } else if (command === "start") {
    const config = JSON.parse(await readFile(`${configDir}/bridge.json`, "utf8"));
    await startServer(config);
    console.log(`ExcelMcp Office.js bridge listening on ${config.origin}.`);
  } else if (command === "health") {
    const config = JSON.parse(await readFile(`${configDir}/bridge.json`, "utf8"));
    const certificate = await readFile(config.certificatePath);
    const result = await new Promise((resolve, reject) => {
      const request = https.get(`${config.origin}/v1/health`, {
        ca: certificate,
        headers: { Authorization: `Bearer ${config.token}` }
      }, (response) => {
        const chunks = [];
        response.on("data", (chunk) => chunks.push(chunk));
        response.on("end", () => resolve({
          ok: response.statusCode >= 200 && response.statusCode < 300,
          body: Buffer.concat(chunks).toString("utf8")
        }));
      });
      request.on("error", reject);
    });
    console.log(JSON.stringify(JSON.parse(result.body), null, 2));
    if (!result.ok) {
      process.exitCode = 1;
    }
  } else {
    throw new Error(
      "Usage: bridge <install|upgrade|start|health|remove> " +
      "[--certificate PATH --private-key PATH --port PORT --config-dir PATH --wef-dir PATH]"
    );
  }
} catch (error) {
  console.error(`Office.js bridge ${command ?? "command"} failed: ${error.message}`);
  process.exitCode = 1;
}

function parseOptions(values) {
  const parsed = {};
  for (let index = 0; index < values.length; index += 2) {
    const key = values[index];
    if (!key?.startsWith("--") || index + 1 >= values.length) {
      throw new Error(`Invalid option '${key ?? ""}'.`);
    }
    parsed[key.slice(2)] = values[index + 1];
  }
  return parsed;
}
