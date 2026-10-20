import assert from 'node:assert/strict';
import { test as nodeTest } from 'node:test';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import {
    compareTrees, publicationDecision, validatePublication, canonicalJson, git as readOnlyGit,
} from '../../scripts/PluginContent.mjs';

function test(name, body) {
    return nodeTest(name, async t => {
        const started = performance.now();
        // Synchronous fixtures delay TAP results; retain the active check on timeout.
        process.stderr.write(`[plugin-publication] START ${name}\n`);
        try {
            return await body(t);
        } finally {
            process.stderr.write(`[plugin-publication] END ${name} (${(performance.now() - started).toFixed(0)} ms)\n`);
        }
    });
}

function payload(version = '2.0.1') {
    const files = new Map();
    for (const name of ['excel-cli', 'excel-mcp']) {
        const prefix = `plugins/${name}/`;
        const manifest = {
            $schema: 'https://agent-plugins.org/schemas/1.0.0/plugin.schema.json',
            name, version, description: `${name} description`,
            author: { name: 'Author', url: 'https://github.com/sbroenne' },
            repository: 'https://github.com/sbroenne/mcp-server-excel-plugins',
            license: 'MIT', keywords: ['excel'],
        };
        for (const [name, text] of Object.entries({
            'plugin.json': JSON.stringify(manifest), 'version.txt': version,
            [`skills/${manifest.name}/VERSION`]: version,
            [`skills/${manifest.name}/SKILL.md`]: 'Use Excel 1.2.3',
            [`skills/${manifest.name}/references/range.md`]: 'range guide',
            'README.md': 'Install using npx @latest',
            [manifest.name === 'excel-cli' ? 'bin/start-cli.ps1' : 'mcp.json']:
                manifest.name === 'excel-cli' ? 'npx @latest' : '{"args":["npx","@latest"]}',
        })) files.set(prefix + name, { bytes: Buffer.from(text), mode: '100644' });
    }
    files.set('.github/plugin/marketplace.json', {
        bytes: Buffer.from(JSON.stringify({
            metadata: { version: '1.0.0' },
            plugins: ['excel-cli', 'excel-mcp'].map(name => ({ name, source: `./plugins/${name}`, version })),
        })), mode: '100644',
    });
    files.set('README.md', { bytes: Buffer.from('Marketplace'), mode: '100644' });
    return files;
}

function edit(files, name, text) { files.set(name, { bytes: Buffer.from(text), mode: '100644' }); }

function scopedPayload(version) {
    const files = payload(version);
    for (const name of ['excel-cli', 'excel-mcp']) {
        for (const file of ['SKILL.md', 'VERSION']) {
            const old = `plugins/${name}/skills/${name}/${file}`;
            files.set(`plugins/${name}/skills/${name}-report-formatting/${file}`, files.get(old));
            files.delete(old);
        }
        files.delete(`plugins/${name}/skills/${name}/references/range.md`);
        edit(files, `plugins/${name}/skills/${name}-report-formatting/references/report-formatting.md`, 'Formatting conventions');
    }
    return files;
}

function discoveryPayload(version) {
    const files = scopedPayload(version);
    edit(files, 'plugins/excel-cli/skills/excel-cli/SKILL.md', 'Discover the npx launcher');
    edit(files, 'plugins/excel-cli/skills/excel-cli/VERSION', version);
    return files;
}

export function registerPolicyTests() {
    test('CLI discovery plus formatting passes publication guards with both skill stamps validated', () => {
        const files = discoveryPayload('2.3.0');
        validatePublication(files);
        assert.deepEqual(compareTrees(payload(), files).changedPlugins, ['excel-cli', 'excel-mcp']);
        assert.deepEqual(compareTrees(files, discoveryPayload('2.4.0')).changedPaths, []);
        for (const skill of ['excel-cli', 'excel-cli-report-formatting']) {
            const stamp = `plugins/excel-cli/skills/${skill}/VERSION`;
            const invalid = discoveryPayload('2.3.0');
            edit(invalid, stamp, '2.0.1');
            assert.throws(() => validatePublication(invalid), /Missing or mismatched/);
            assert.deepEqual(validatePublication(invalid, { repairStamps: true }).repairs, [stamp]);
            invalid.delete(`plugins/excel-cli/skills/${skill}/SKILL.md`);
            assert.throws(() => validatePublication(invalid), /Missing|Mixed/);
        }
        for (const extra of [
            'plugins/excel-cli/skills/excel-cli/references/range.md',
            'plugins/excel-cli/skills/unexpected/SKILL.md',
            'plugins/excel-mcp/skills/excel-mcp/SKILL.md',
        ]) {
            const invalid = discoveryPayload('2.3.0');
            edit(invalid, extra, 'Unexpected old or unknown guidance');
            assert.throws(() => validatePublication(invalid), /Mixed/);
        }
        validatePublication(payload());
        validatePublication(scopedPayload('2.3.0'));
    });

    test('publication workflow validates every recognized skill and rejects stale discovery stamps', () => {
        const workflow = fs.readFileSync(path.join(repoRoot, '.github', 'workflows', 'publish-plugins.yml'), 'utf8');
        const section = workflow.split('      - name: Validate prepared skills')[1]
            .split('      - name: Checkout output-only published repository')[0];
        const body = section.split('        run: |')[1]
            .replace(/          python -m pip install[^\n]*\n[^\n]*\n/, '')
            .replace(/^          /gm, '');
        for (const files of [payload('2.3.0'), scopedPayload('2.3.0'), discoveryPayload('2.3.0')]) {
            const fixture = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-workflow-skills-'));
            try {
                writeTree(fixture, new Map([...files].filter(([name]) => name.startsWith('plugins/'))
                    .map(([name, file]) => [`built-plugins/${name.slice('plugins/'.length)}`, file])));
                const scripts = path.join(fixture, 'automation', 'scripts');
                fs.mkdirSync(scripts, { recursive: true });
                fs.copyFileSync(path.join(repoRoot, 'scripts', 'PluginContent.mjs'), path.join(scripts, 'PluginContent.mjs'));
                const script = `
                    $env:VERSION = '2.3.0'
                    $script:validated = @()
                    function skills-ref {
                        param($action, $skillPath)
                        if ($action -ne 'validate') { throw 'Unexpected skill command.' }
                        $script:validated += (Split-Path $skillPath -Leaf)
                    }
                    ${body}
                    ConvertTo-Json -InputObject $script:validated -Compress
                `;
                const run = () => JSON.parse(command('pwsh', ['-NoProfile', '-NonInteractive', '-Command', script], fixture));
                const expected = [...files.keys()].filter(name => /\/skills\/[^/]+\/SKILL\.md$/.test(name))
                    .map(name => name.split('/').at(-2)).sort();
                assert.deepEqual(run().sort(), expected);
                if (expected.includes('excel-cli-report-formatting') && expected.includes('excel-cli')) {
                    fs.writeFileSync(path.join(fixture, 'built-plugins', 'excel-cli', 'skills', 'excel-cli', 'VERSION'), '2.0.1');
                    assert.throws(run, /missing or mismatched skill version/);
                }
            } finally { fs.rmSync(fixture, { recursive: true }); }
        }
    });

    test('prepared publisher accepts CLI discovery plus formatting in a no-write preview', () => {
        const fixture = publicationFixture();
        try {
            fs.rmSync(fixture.built, { recursive: true });
            writeTree(fixture.built, new Map([...discoveryPayload('2.3.0')]
                .filter(([name]) => name.startsWith('plugins/'))
                .map(([name, file]) => [name.slice('plugins/'.length), file])));
            const text = publishFixture(fixture, ['-Preview']);
            const result = JSON.parse(text.slice(text.search(/^\{\r?$/m)));
            assert.equal(result.status, 'preview');
            assert.equal(result.decision, 'published');
            assert.deepEqual(result.changed_plugins, ['excel-cli', 'excel-mcp']);
            assert.equal(result.handoff, false);
            assert.equal(command('git', ['rev-parse', 'HEAD'], fixture.output).trim(), fixture.original);
            assert.equal(command('git', ['tag', '--list', 'v2.3.0'], fixture.output).trim(), '');
            assert.equal(command('git', ['status', '--porcelain'], fixture.output).trim(), '');
        } finally { fs.rmSync(fixture.root, { recursive: true }); }
    });

    test('formatting migration changes content and preserves version-only skip behavior', () => {
        const candidate = scopedPayload('2.3.0');
        validatePublication(candidate);
        assert.deepEqual(compareTrees(payload(), candidate).changedPlugins, ['excel-cli', 'excel-mcp']);
        assert.deepEqual(compareTrees(candidate, scopedPayload('2.4.0')).changedPaths, []);
        edit(candidate, 'plugins/excel-mcp/skills/excel-mcp/SKILL.md', 'Old broad skill');
        assert.throws(() => validatePublication(candidate), /Mixed/);
    });

    test('Claude Code marketplace must match the canonical catalog before versions are ignored', () => {
        function withClaude(version, claudeVersion = version) {
            const files = payload(version);
            edit(files, '.claude-plugin/marketplace.json', JSON.stringify({
                plugins: ['excel-cli', 'excel-mcp'].map(name => ({ name, source: `./plugins/${name}`, version: claudeVersion })),
            }));
            return files;
        }
        assert.deepEqual(compareTrees(withClaude('2.0.1'), withClaude('2.3.0')).changedPaths, []);
        assert.throws(() => compareTrees(payload(), withClaude('2.3.0', '2.0.1')), /Claude Code marketplace entry/);
        const missingEntry = withClaude('2.0.1');
        edit(missingEntry, '.claude-plugin/marketplace.json', JSON.stringify({
            plugins: [{ name: 'excel-cli', source: './plugins/excel-cli', version: '2.0.1' }],
        }));
        assert.throws(() => validatePublication(missingEntry), /exactly two Claude Code/);
        const wrongSource = withClaude('2.0.1');
        edit(wrongSource, '.claude-plugin/marketplace.json', JSON.stringify({
            plugins: ['excel-cli', 'excel-mcp'].map(name => ({ name, source: './plugins/other', version: '2.0.1' })),
        }));
        assert.throws(() => validatePublication(wrongSource), /Claude Code marketplace entry/);
    });

    test('identical and release-bookkeeping-only payloads skip every publication write and handoff', () => {
        const comparison = compareTrees(payload(), payload('2.3.0'));
        assert.deepEqual(comparison.changedPlugins, []);
        assert.deepEqual(comparison.changedPaths, []);
        assert.equal(comparison.baselineFingerprint, comparison.candidateFingerprint);
        assert.deepEqual(publicationDecision(comparison, {
            currentVersion: '2.0.1', version: '2.3.0', tagExists: false, manualRepair: false,
        }), { status: 'skipped', write: false, handoff: false });
    });

    test('root-only changes publish without a plugin listing handoff', () => {
        const candidate = payload('2.3.0');
        edit(candidate, 'README.md', 'New marketplace instructions');
        const comparison = compareTrees(payload(), candidate);
        assert.deepEqual(comparison.changedPlugins, []);
        assert.deepEqual(comparison.changedPaths, ['README.md']);
        assert.deepEqual(publicationDecision(comparison, {
            currentVersion: '2.0.1', version: '2.3.0', tagExists: false,
        }), { status: 'published', write: true, handoff: false });
    });

    test('each real distributed surface, added/removed files and launch versions are meaningful', () => {
        for (const name of [
            'README.md', 'plugin.json', 'mcp.json', 'bin/helper.ps1',
            'skills/excel-mcp/SKILL.md', 'skills/excel-mcp/references/range.md', 'assets/icon.png',
        ]) {
            const baseline = payload(), candidate = payload('2.3.0');
            const fullPath = `plugins/excel-mcp/${name}`;
            if (name === 'plugin.json') {
                const manifest = JSON.parse(candidate.get(fullPath).bytes);
                manifest.description = 'New description 9.9.9';
                edit(candidate, fullPath, JSON.stringify(manifest));
            } else if (name === 'mcp.json') edit(candidate, fullPath, '{"args":["npx","@2.3.0"]}');
            else edit(candidate, fullPath, 'Changed functional text 9.9.9');
            const comparison = compareTrees(baseline, candidate);
            assert.deepEqual(comparison.changedPlugins, ['excel-mcp'], name);
            assert.ok(comparison.changedPaths.includes(fullPath), name);
            assert.equal(publicationDecision(comparison, {
                currentVersion: '2.0.1', version: '2.3.0', tagExists: false,
            }).handoff, true);
            candidate.delete(fullPath);
            if (name.includes('helper') || name.includes('assets')) {
                edit(baseline, fullPath, 'removed');
                assert.ok(compareTrees(baseline, candidate).changedPaths.includes(fullPath));
            }
        }
    });

    test('canonical JSON preserves values and array order but not object order or formatting', () => {
        assert.equal(canonicalJson('{"b":2,"a":{"z":1,"y":true}}'), canonicalJson(' { "a": { "y":true, "z":1 }, "b":2 } '));
        assert.notEqual(canonicalJson('{"args":["a","b"]}'), canonicalJson('{"args":["b","a"]}'));
        const baseline = payload(), candidate = payload('2.3.0');
        const name = 'plugins/excel-mcp/plugin.json';
        edit(candidate, name, JSON.stringify(Object.fromEntries(Object.entries(JSON.parse(candidate.get(name).bytes)).reverse()), null, 4));
        assert.deepEqual(compareTrees(baseline, candidate).changedPaths, []);
        const market = JSON.parse(candidate.get('.github/plugin/marketplace.json').bytes);
        market.metadata.version = '2.0.0';
        edit(candidate, '.github/plugin/marketplace.json', JSON.stringify(market));
        assert.deepEqual(compareTrees(baseline, candidate).changedPaths, ['.github/plugin/marketplace.json']);
    });

    test('ambiguous duplicate JSON keys and numeric values that would lose precision fail visibly', () => {
        assert.throws(() => canonicalJson('{"version":"2.0.1","version":"2.3.0"}'), /Duplicate/);
        assert.throws(() => canonicalJson('{"value":9007199254740993}'), /precision/);
        assert.throws(() => canonicalJson('{"value":0.10000000000000001}'), /precision/);
        assert.equal(canonicalJson('{"value":1.00e0}'), canonicalJson('{"value":1}'));
    });

    test('invalid identity, JSON, stamps, missing required files and unsupported plugin layouts fail visibly', () => {
        for (const [name, value] of [
            ['plugins/excel-mcp/plugin.json', '{}'], ['plugins/excel-mcp/plugin.json', '{invalid'],
            ['plugins/excel-mcp/version.txt', '2.0.2'], ['plugins/excel-mcp/skills/excel-mcp/VERSION', null],
            ['plugins/excel-mcp/mcp.json', null], ['plugins/unknown/plugin.json', '{}'],
            ['plugins/excel-mcp/binary.dll', 'runtime'],
        ]) {
            const candidate = payload();
            if (value === null) candidate.delete(name); else edit(candidate, name, value);
            assert.throws(() => validatePublication(candidate));
        }
    });

    test('manual repair does not accept a malformed manifest version', () => {
        const baseline = payload();
        const name = 'plugins/excel-cli/plugin.json';
        const manifest = JSON.parse(baseline.get(name).bytes);
        manifest.version = 'invalid';
        edit(baseline, name, JSON.stringify(manifest));
        assert.throws(() => validatePublication(baseline, { repairStamps: true }), /Invalid release version/);
    });

    test('skipped releases retain prior version; next real release publishes; downgrades and conflicting tags fail', () => {
        const baseline = payload();
        for (const version of ['2.0.2', '2.1.0', '2.2.0']) {
            assert.equal(publicationDecision(compareTrees(baseline, payload(version)), {
                currentVersion: '2.0.1', version, tagExists: false,
            }).status, 'skipped');
        }
        const candidate = payload('2.3.0');
        edit(candidate, 'plugins/excel-cli/README.md', 'New CLI guidance');
        const comparison = compareTrees(baseline, candidate);
        assert.deepEqual(comparison.changedPlugins, ['excel-cli']);
        assert.equal(publicationDecision(comparison, { currentVersion: '2.0.1', version: '2.3.0' }).write, true);
        assert.throws(() => publicationDecision(comparison, { currentVersion: '2.4.0', version: '2.3.0' }), /Downgrade/);
        assert.throws(() => publicationDecision(comparison, { currentVersion: '2.0.1', version: '2.3.0', tagExists: true }), /tag/);
    });

    test('authorized exact repair can restore missing tag/stamps without rewriting an immutable existing tag', () => {
        const comparison = compareTrees(payload(), payload());
        assert.equal(publicationDecision(comparison, {
            currentVersion: '2.0.1', version: '2.0.1', manualRepair: true, tagExists: false,
        }).write, true);
        assert.equal(publicationDecision(comparison, {
            currentVersion: '2.0.1', version: '2.0.1', manualRepair: true, tagExists: true, rawChanged: false,
        }).write, false);
    });
}

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const gitLocalVariables = execFileSync('git', ['rev-parse', '--local-env-vars'], {
    cwd: repoRoot, encoding: 'utf8', windowsHide: true,
}).trim().split(/\r?\n/);

function command(executable, args, cwd, inherited = process.env) {
    const env = { ...inherited, GITHUB_OUTPUT: '', GITHUB_STEP_SUMMARY: '' };
    // Git hooks export source-repository context; fixtures must use their own Git state.
    for (const name of gitLocalVariables) delete env[name];
    for (const name of Object.keys(env)) {
        if (/^GIT_CONFIG_(KEY|VALUE)_\d+$/.test(name)) delete env[name];
    }
    return execFileSync(executable, args, { cwd, encoding: 'utf8', windowsHide: true, env });
}

export function registerGitIsolationTests() {
    test('disposable Git commands cannot inherit a hook repository, worktree or index', () => {
        const root = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-git-isolation-'));
        const foreignIndex = path.join(root, 'foreign-index');
        fs.writeFileSync(foreignIndex, 'Do not touch');
        try {
            const inherited = { ...process.env, GIT_DIR: path.join(root, 'foreign.git'),
                GIT_WORK_TREE: path.join(root, 'foreign-worktree'), GIT_INDEX_FILE: foreignIndex,
                GIT_CONFIG_COUNT: '1', GIT_CONFIG_KEY_0: 'core.bare', GIT_CONFIG_VALUE_0: 'true' };
            command('git', ['init', '--quiet'], root, inherited);
            assert.equal(command(process.execPath, ['-p', "Object.hasOwn(process.env, 'GIT_CONFIG_KEY_0')"], root, inherited).trim(), 'false');
            assert.ok(fs.existsSync(path.join(root, '.git')));
            assert.equal(fs.readFileSync(foreignIndex, 'utf8'), 'Do not touch');
            assert.equal(fs.existsSync(inherited.GIT_DIR), false);
        } finally { fs.rmSync(root, { recursive: true }); }
    });

    test('shared Git preparation commands cannot inherit write credentials or injected authentication settings', () => {
        const root = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-git-auth-isolation-'));
        try {
            init(root);
            fs.writeFileSync(path.join(root, 'probe.mjs'),
                'console.log(JSON.stringify(Object.entries(process.env).filter(([name,value]) => /(?:TOKEN|_PAT)$|^GIT_CONFIG(?:$|_)|^(?:GIT|SSH)_ASKPASS$|^GIT_SSH(?:$|_)/i.test(name) && value.includes("dummy")).map(([name]) => name)));');
            const env = { PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
                PLUGINS_REPO_TOKEN: 'dummy-write-token', GH_TOKEN: 'dummy-read-token',
                GIT_CONFIG_COUNT: '1', GIT_CONFIG_KEY_0: 'http.https://github.com/.extraheader',
                GIT_CONFIG_VALUE_0: 'AUTHORIZATION: dummy-header',
                GIT_CONFIG_PARAMETERS: "'http.extraheader=dummy-parameter'" };
            const probe = readOnlyGit(root, ['-c', `alias.probe=!"${process.execPath}" probe.mjs`, 'probe'], { env });
            assert.deepEqual(JSON.parse(probe), []);
            assert.equal(env.PLUGINS_REPO_TOKEN, 'dummy-write-token');
        } finally { fs.rmSync(root, { recursive: true }); }
    });
}

function commit(directory, message) {
    command('git', ['add', '-A'], directory);
    command('git', ['commit', '--quiet', '-m', message], directory);
    return command('git', ['rev-parse', 'HEAD'], directory).trim();
}
function init(directory) {
    fs.mkdirSync(directory, { recursive: true });
    command('git', ['init', '--quiet', '-b', 'main'], directory);
    command('git', ['config', 'user.name', 'Fixture'], directory);
    command('git', ['config', 'user.email', 'fixture@example.invalid'], directory);
    command('git', ['config', 'core.autocrlf', 'false'], directory);
}
function writeTree(directory, files) {
    for (const [name, file] of files) {
        fs.mkdirSync(path.dirname(path.join(directory, name)), { recursive: true });
        fs.writeFileSync(path.join(directory, name), file.bytes);
    }
}

function publicationFixture() {
    const currentVersion = '2.0.1', candidateVersion = '2.3.0';
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-publication-'));
    const source = path.join(root, 'source'), output = path.join(root, 'output'), built = path.join(root, 'built');
    init(source);
    init(output);
    fs.mkdirSync(path.join(source, 'scripts'));
    fs.mkdirSync(path.join(source, '.github', 'plugins', 'marketplace-repo'), { recursive: true });
    fs.writeFileSync(path.join(source, '.github', 'plugins', 'marketplace-repo', 'README.md'), 'Marketplace');
    fs.writeFileSync(path.join(source, 'scripts', 'Sync-PublishedPluginRepo.ps1'), `
param($PublishedRepoDir,$BuiltPluginsDir,$Version)
$ErrorActionPreference='Stop'
foreach ($name in @('excel-cli','excel-mcp')) {
    $destination=Join-Path $PublishedRepoDir "plugins/$name"
    Remove-Item -LiteralPath $destination -Recurse -Force
    Copy-Item -LiteralPath (Join-Path $BuiltPluginsDir $name) -Destination $destination -Recurse
}
$market=Get-Content (Join-Path $PublishedRepoDir '.github/plugin/marketplace.json') -Raw | ConvertFrom-Json
foreach ($plugin in $market.plugins) { $plugin.version=$Version }
$market | ConvertTo-Json -Depth 20 | Set-Content (Join-Path $PublishedRepoDir '.github/plugin/marketplace.json')
Copy-Item -LiteralPath (Join-Path $PSScriptRoot '../.github/plugins/marketplace-repo/README.md') -Destination (Join-Path $PublishedRepoDir 'README.md')
`);
    commit(source, 'Original source');
    command('git', ['tag', `v${currentVersion}`], source);
    writeTree(output, payload(currentVersion));
    const original = commit(output, 'Original publication');
    command('git', ['tag', `v${currentVersion}`], output);
    fs.writeFileSync(path.join(source, 'release.txt'), candidateVersion);
    const sourceCommit = commit(source, 'Candidate source');
    command('git', ['tag', `v${candidateVersion}`], source);
    command('git', ['checkout', '--quiet', '--detach', sourceCommit], source);
    for (const [name, file] of payload(candidateVersion)) if (name.startsWith('plugins/')) {
        writeTree(built, new Map([[name.slice('plugins/'.length), file]]));
    }
    return { root, source, output, built, original, sourceCommit };
}
function publishFixture(fixture, flags = []) {
    return command('pwsh', ['-NoProfile', '-File', path.join(repoRoot, 'scripts', 'Publish-PreparedPlugins.ps1'),
        '-SourceDirectory', fixture.source, '-PublishedRepoDirectory', fixture.output,
        '-BuiltPluginsDirectory', fixture.built, '-Version', '2.3.0', '-SourceCommit', fixture.sourceCommit, ...flags], repoRoot);
}

export function registerDeliveryTests() {
    test('source guard accepts clean exact source and retains the existing no-change publication', () => {
        const fixture = publicationFixture();
        try {
            const result = JSON.parse(publishFixture(fixture, ['-Preview']));
            assert.equal(result.status, 'preview');
            assert.equal(result.decision, 'skipped');
            assert.equal(result.published_tag, 'v2.0.1');
            assert.equal(result.published_commit, fixture.original);
            assert.equal(result.handoff, false);
            assert.deepEqual(result.changed_paths, []);
            assert.equal(command('git', ['rev-parse', 'HEAD'], fixture.source).trim(), fixture.sourceCommit);
            assert.equal(command('git', ['status', '--porcelain'], fixture.source).trim(), '');
            assert.equal(command('git', ['rev-parse', 'HEAD'], fixture.output).trim(), fixture.original);
            assert.equal(command('git', ['status', '--porcelain'], fixture.output).trim(), '');
            assert.equal(command('git', ['tag', '--list', 'v2.3.0'], fixture.output).trim(), '');
        } finally { fs.rmSync(fixture.root, { recursive: true }); }
    });

    test('source guard identifies a wrong checkout without changing either repository', () => {
        const fixture = publicationFixture();
        try {
            command('git', ['checkout', '--quiet', '--detach', 'v2.0.1'], fixture.source);
            const actualCommit = command('git', ['rev-parse', 'HEAD'], fixture.source).trim();
            assert.throws(() => publishFixture(fixture, ['-Preview']), error => {
                const diagnostic = error.stderr.toString();
                assert.match(diagnostic, /Source checkout HEAD/);
                assert.ok(diagnostic.includes(actualCommit));
                assert.ok(diagnostic.includes(fixture.sourceCommit));
                return true;
            });
            assert.equal(command('git', ['rev-parse', 'HEAD'], fixture.source).trim(), actualCommit);
            assert.equal(command('git', ['status', '--porcelain'], fixture.source).trim(), '');
            assert.equal(command('git', ['rev-parse', 'HEAD'], fixture.output).trim(), fixture.original);
            assert.equal(command('git', ['status', '--porcelain'], fixture.output).trim(), '');
            assert.equal(command('git', ['tag', '--list', 'v2.3.0'], fixture.output).trim(), '');
        } finally { fs.rmSync(fixture.root, { recursive: true }); }
    });

    for (const state of ['unstaged', 'staged', 'untracked']) {
        test(`source guard identifies ${state} changes without cleaning or publishing`, () => {
            const fixture = publicationFixture();
            try {
                const file = state === 'untracked' ? 'unexpected.txt' : 'release.txt';
                fs.writeFileSync(path.join(fixture.source, file), 'Keep this evidence');
                if (state === 'staged') command('git', ['add', file], fixture.source);
                const before = command('git', ['status', '--porcelain'], fixture.source);
                assert.ok(before.includes(file));
                assert.throws(() => publishFixture(fixture, ['-Preview']), error => {
                    const diagnostic = error.stderr.toString();
                    assert.match(diagnostic, /Source checkout must be clean/);
                    assert.ok(diagnostic.includes(fixture.sourceCommit));
                    assert.match(diagnostic, /Git status:/);
                    assert.ok(diagnostic.includes(file));
                    return true;
                });
                assert.equal(command('git', ['rev-parse', 'HEAD'], fixture.source).trim(), fixture.sourceCommit);
                assert.equal(command('git', ['status', '--porcelain'], fixture.source), before);
                assert.equal(fs.readFileSync(path.join(fixture.source, file), 'utf8'), 'Keep this evidence');
                assert.equal(command('git', ['rev-parse', 'HEAD'], fixture.output).trim(), fixture.original);
                assert.equal(command('git', ['status', '--porcelain'], fixture.output).trim(), '');
                assert.equal(command('git', ['tag', '--list', 'v2.3.0'], fixture.output).trim(), '');
            } finally { fs.rmSync(fixture.root, { recursive: true }); }
        });
    }

    test('real publication pushes a committed repair and file/folder swaps to a disposable local remote', () => {
        const fixture = publicationFixture();
        try {
            const old = { bytes: Buffer.from('Old asset'), mode: '100644' };
            const repaired = { bytes: Buffer.from('Repaired guidance'), mode: '100644' };
            writeTree(fixture.output, new Map([
                ['plugins/excel-cli/assets/icons/icon.png', old], ['plugins/excel-mcp/assets', old],
                ['plugins/excel-cli/README.md', repaired],
            ]));
            const repairedHead = commit(fixture.output, 'Simulate committed repair after the release tag');
            const remote = path.join(fixture.root, 'remote.git');
            command('git', ['clone', '--quiet', '--bare', fixture.output, remote], fixture.root);
            command('git', ['remote', 'add', 'origin', remote], fixture.output);
            fs.writeFileSync(path.join(fixture.built, 'excel-cli', 'assets'), 'Replacement asset');
            writeTree(fixture.built, new Map([
                ['excel-mcp/assets/icon.png', { bytes: Buffer.from('New asset'), mode: '100644' }],
                ['excel-cli/README.md', repaired],
            ]));
            const text = publishFixture(fixture);
            const result = JSON.parse(text.slice(text.search(/^\{\r?$/m)));
            assert.equal(result.status, 'published');
            assert.equal(result.handoff, true);
            assert.equal(result.published_tag, 'v2.3.0');
            assert.ok(!result.destination_changed_paths.includes('plugins/excel-cli/README.md'));
            assert.ok(result.distributed_changed_paths.includes('plugins/excel-cli/README.md'));
            assert.ok(result.changed_plugins.includes('excel-cli'));
            const head = command('git', ['rev-parse', 'HEAD'], fixture.output).trim();
            assert.notEqual(head, repairedHead);
            assert.equal(command('git', ['rev-parse', 'v2.0.1'], fixture.output).trim(), fixture.original);
            assert.equal(command('git', [`--git-dir=${remote}`, 'rev-parse', 'v2.3.0^{commit}'], fixture.root).trim(), head);
            const pushed = name => command('git', [`--git-dir=${remote}`, 'show', `v2.3.0:${name}`], fixture.root);
            assert.equal(pushed('plugins/excel-cli/assets'), 'Replacement asset');
            assert.equal(pushed('plugins/excel-mcp/assets/icon.png'), 'New asset');
            assert.equal(pushed('plugins/excel-cli/README.md'), 'Repaired guidance');
            assert.equal(command('git', ['status', '--porcelain'], fixture.output).trim(), '');
        } finally { fs.rmSync(fixture.root, { recursive: true }); }
    });
}
