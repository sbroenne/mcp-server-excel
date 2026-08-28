import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const extensionRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const manifest = JSON.parse(readFileSync(resolve(extensionRoot, 'package.json'), 'utf8'));

const providers = manifest.contributes?.mcpServerDefinitionProviders ?? [];
assert.ok(providers.length > 0, 'At least one MCP server definition provider must be contributed.');
for (const provider of providers) {
  assert.ok(provider.id?.trim(), 'Every MCP server definition provider must have an ID.');
  assert.ok(provider.label?.trim(), `MCP server definition provider '${provider.id}' must have a label.`);
}

const skills = manifest.contributes?.chatSkills ?? [];
assert.ok(skills.length > 0, 'At least one chat skill must be contributed.');
for (const skill of skills) {
  assert.ok(skill.path?.trim(), 'Every chat skill must have a path.');
  assert.ok(skill.name?.trim(), `Chat skill '${skill.path}' must declare a display name for the Features tab.`);
  assert.ok(skill.description?.trim(), `Chat skill '${skill.path}' must declare a description for the Features tab.`);

  const skillPath = resolve(extensionRoot, skill.path);
  assert.ok(existsSync(skillPath), `Chat skill file does not exist: ${skill.path}`);

  const skillContents = readFileSync(skillPath, 'utf8');
  const frontmatterName = skillContents.match(/^---\s*\r?\n[\s\S]*?^name:\s*['"]?([^'"\r\n]+)['"]?\s*$[\s\S]*?^---\s*$/m)?.[1]?.trim();
  assert.equal(skill.name, frontmatterName, `Chat skill '${skill.path}' name must match its SKILL.md frontmatter.`);
}

console.log('VS Code extension feature metadata is complete.');
