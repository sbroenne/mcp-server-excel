import { spawnSync } from "node:child_process";
import { readdirSync } from "node:fs";

const files = [
  ...readdirSync(new URL("../src/", import.meta.url))
    .filter((file) => file.endsWith(".mjs"))
    .map((file) => new URL(`../src/${file}`, import.meta.url)),
  ...readdirSync(new URL("../test/", import.meta.url))
    .filter((file) => file.endsWith(".mjs"))
    .map((file) => new URL(`../test/${file}`, import.meta.url))
];
for (const file of files) {
  const result = spawnSync(process.execPath, ["--check", file.pathname], { stdio: "inherit" });
  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}
