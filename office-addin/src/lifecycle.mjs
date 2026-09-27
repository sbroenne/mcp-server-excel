import { randomBytes } from "node:crypto";
import { chmod, copyFile, mkdir, readFile, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import { ADDIN_VERSION, DEFAULT_PORT, PROTOCOL_VERSION } from "./constants.mjs";
import { createManifest } from "./manifest.mjs";

export function defaultPaths(home = process.env.HOME) {
  if (!home) {
    throw new Error("HOME is required.");
  }
  return {
    configDir: path.join(home, "Library", "Application Support", "ExcelMcp", "officejs"),
    wefDir: path.join(
      home,
      "Library",
      "Containers",
      "com.microsoft.Excel",
      "Data",
      "Documents",
      "wef"
    )
  };
}

export async function installBridge(options) {
  const port = options.port ?? DEFAULT_PORT;
  if (!Number.isInteger(port) || port < 1024 || port > 65535) {
    throw new TypeError("port must be an integer between 1024 and 65535.");
  }
  await requireRegularFile(options.certificate, "certificate");
  await requireRegularFile(options.privateKey, "private key");
  await mkdir(options.configDir, { recursive: true, mode: 0o700 });
  await mkdir(options.wefDir, { recursive: true, mode: 0o700 });

  const configPath = path.join(options.configDir, "bridge.json");
  let token = randomBytes(32).toString("base64url");
  try {
    const previous = JSON.parse(await readFile(configPath, "utf8"));
    if (typeof previous.token === "string" && previous.token.length >= 32) {
      token = previous.token;
    }
  } catch (error) {
    if (error.code !== "ENOENT") {
      throw error;
    }
  }

  const certificatePath = path.join(options.configDir, "localhost.pem");
  const privateKeyPath = path.join(options.configDir, "localhost-key.pem");
  await copyFile(options.certificate, certificatePath);
  await copyFile(options.privateKey, privateKeyPath);
  await chmod(certificatePath, 0o600);
  await chmod(privateKeyPath, 0o600);

  const config = {
    addinVersion: ADDIN_VERSION,
    protocolVersion: PROTOCOL_VERSION,
    port,
    origin: `https://localhost:${port}`,
    token,
    certificatePath,
    privateKeyPath
  };
  await writeFile(configPath, `${JSON.stringify(config, null, 2)}\n`, { mode: 0o600 });
  await chmod(configPath, 0o600);
  const manifestPath = path.join(options.wefDir, "excelmcp-officejs.xml");
  await writeFile(manifestPath, createManifest(config), { mode: 0o600 });
  await chmod(manifestPath, 0o600);
  return { configPath, manifestPath, upgraded: true };
}

export async function removeBridge(options) {
  await rm(path.join(options.wefDir, "excelmcp-officejs.xml"), { force: true });
  await rm(options.configDir, { force: true, recursive: true });
}

async function requireRegularFile(filePath, name) {
  if (!filePath) {
    throw new TypeError(`${name} path is required.`);
  }
  const info = await stat(filePath);
  if (!info.isFile()) {
    throw new TypeError(`${name} must be a regular file.`);
  }
}
