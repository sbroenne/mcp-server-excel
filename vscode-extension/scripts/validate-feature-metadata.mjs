import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const extensionRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const manifest = JSON.parse(readFileSync(resolve(extensionRoot, 'package.json'), 'utf8'));
const extensionSource = readFileSync(resolve(extensionRoot, 'src', 'extension.ts'), 'utf8');

function assertText(value, message) {
  assert.ok(typeof value === 'string' && value.trim().length > 0, message);
}

function assertPackagedFile(path, label) {
  assertText(path, `${label} path must be declared.`);
  const absolutePath = resolve(extensionRoot, path);
  assert.ok(!relative(extensionRoot, absolutePath).startsWith('..'), `${label} path must stay within the extension: ${path}`);
  assert.ok(existsSync(absolutePath), `${label} file does not exist: ${path}`);
  return absolutePath;
}

assertText(manifest.name, 'Extension name must be declared.');
assertText(manifest.displayName, 'Extension display name must be declared.');
assertText(manifest.description, 'Extension description must be declared.');
assertText(manifest.publisher, 'Extension publisher must be declared.');
assertText(manifest.version, 'Extension version must be declared.');
assertText(manifest.engines?.vscode, 'Minimum VS Code engine version must be declared.');
assert.ok(manifest.categories?.length > 0, 'At least one Marketplace category must be declared.');
assert.ok(manifest.keywords?.length > 0, 'At least one Marketplace keyword must be declared.');

assertPackagedFile('README.md', 'Marketplace README');
assertPackagedFile('LICENSE', 'License');
const iconPath = assertPackagedFile(manifest.icon, 'Marketplace icon');
const icon = readFileSync(iconPath);
assert.equal(icon.subarray(1, 4).toString('ascii'), 'PNG', 'Marketplace icon must be a PNG file.');
assert.ok(icon.length >= 24, 'Marketplace icon is not a valid PNG file.');
const iconWidth = icon.readUInt32BE(16);
const iconHeight = icon.readUInt32BE(20);
assert.ok(iconWidth >= 128 && iconHeight >= 128, `Marketplace icon must be at least 128x128 pixels; found ${iconWidth}x${iconHeight}.`);

const providers = manifest.contributes?.mcpServerDefinitionProviders ?? [];
assert.ok(providers.length > 0, 'At least one MCP server definition provider must be contributed.');
const providerIds = new Set();
for (const provider of providers) {
  assertText(provider.id, 'Every MCP server definition provider must have an ID.');
  assertText(provider.label, `MCP server definition provider '${provider.id}' must have a label.`);
  assert.ok(!providerIds.has(provider.id), `MCP server definition provider ID is duplicated: ${provider.id}`);
  providerIds.add(provider.id);
  assert.ok(
    extensionSource.includes(`registerMcpServerDefinitionProvider('${provider.id}'`),
    `MCP server definition provider '${provider.id}' must be registered by src/extension.ts.`
  );
}

const skills = manifest.contributes?.chatSkills ?? [];
assert.ok(skills.length > 0, 'At least one chat skill must be contributed.');
const skillNames = new Set();
for (const skill of skills) {
  assertText(skill.path, 'Every chat skill must have a path.');
  assertText(skill.name, `Chat skill '${skill.path}' must declare a display name for the Features tab.`);
  assertText(skill.description, `Chat skill '${skill.path}' must declare a description for the Features tab.`);
  assert.ok(!skillNames.has(skill.name), `Chat skill name is duplicated: ${skill.name}`);
  skillNames.add(skill.name);

  const skillPath = assertPackagedFile(skill.path, 'Chat skill');
  const skillContents = readFileSync(skillPath, 'utf8');
  const frontmatterName = skillContents.match(/^---\s*\r?\n[\s\S]*?^name:\s*['"]?([^'"\r\n]+)['"]?\s*$[\s\S]*?^---\s*$/m)?.[1]?.trim();
  assert.equal(skill.name, frontmatterName, `Chat skill '${skill.path}' name must match its SKILL.md frontmatter.`);
}

console.log('VS Code extension Marketplace and feature metadata is complete.');
