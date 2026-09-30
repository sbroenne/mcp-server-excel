import assert from 'node:assert/strict';
import { test } from 'node:test';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import {
    compareTrees, publicationDecision, validatePublication, canonicalJson, hash,
} from '../../scripts/PluginContent.mjs';
import {
    planListings, parseState, stateMarker, assertListingPatch, assertAllowedPaths,
} from '../../scripts/AwesomeCopilotPolicy.mjs';
import { validateRequest, writesEnabled, submit, assertTemplate } from '../../scripts/Update-AwesomeCopilot.mjs';
import * as updater from '../../scripts/Update-AwesomeCopilot.mjs';

const sha1 = '1'.repeat(40), sha2 = '2'.repeat(40), sha3 = '3'.repeat(40), sha4 = '4'.repeat(40);
function listings() {
    return ['excel-cli', 'excel-mcp'].map(name => ({
        name, description: 'Curated description', version: '2.0.1',
        author: { name: 'Author' }, repository: 'https://github.com/sbroenne/mcp-server-excel-plugins',
        license: 'MIT', keywords: ['excel'],
        source: { source: 'github', repo: 'sbroenne/mcp-server-excel-plugins', path: `plugins/${name}`, ref: 'v2.0.1', sha: sha1 },
    }));
}

function plan(candidate = payload('2.3.0'), extras = {}) {
    const snapshots = new Map([[sha1, payload()], [sha2, candidate]]);
    return planListings({
        listings: listings(), tag: 'v2.3.0', commit: sha2,
        getTree: sha => { assert.ok(snapshots.has(sha), 'Missing published commit'); return snapshots.get(sha); },
        resolveTag: tag => ({ 'v2.0.1': sha1, 'v2.3.0': sha2 })[tag],
        pulls: [], ...extras,
    });
}

function pending(result, extras = {}) {
    const state = { ...result.state, head: sha3, base: sha1, bodyFingerprint: hash('') };
    return {
        number: 42, state: 'open', merged_at: null, body: stateMarker(state),
        user: { login: 'sbroenne' },
        head: { ref: 'excel-plugin-updates-123456abcdef', sha: sha3, repo: { full_name: 'sbroenne/awesome-copilot' } },
        base: { ref: 'main', repo: { full_name: 'github/awesome-copilot' } }, ...extras,
    };
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
        for (const [path, text] of Object.entries({
            'plugin.json': JSON.stringify(manifest), 'version.txt': version,
            [`skills/${name}/VERSION`]: version,
            [`skills/${name}/SKILL.md`]: 'Use Excel 1.2.3',
            [`skills/${name}/references/range.md`]: 'range guide',
            'README.md': 'Install using npx @latest',
            [name === 'excel-cli' ? 'bin/start-cli.ps1' : 'mcp.json']:
                name === 'excel-cli' ? 'npx @latest' : '{"args":["npx","@latest"]}',
        })) files.set(prefix + path, { bytes: Buffer.from(text), mode: '100644' });
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

function edit(files, path, text) { files.set(path, { bytes: Buffer.from(text), mode: '100644' }); }

test('identical and release-bookkeeping-only payloads skip every publication write and handoff', () => {
    const baseline = payload(), candidate = payload('2.3.0');
    const comparison = compareTrees(baseline, candidate);
    assert.deepEqual(comparison.changedPlugins, []);
    assert.deepEqual(comparison.changedPaths, []);
    assert.equal(comparison.baselineFingerprint, comparison.candidateFingerprint);
    const decision = publicationDecision(comparison, {
        currentVersion: '2.0.1', version: '2.3.0', tagExists: false, manualRepair: false,
    });
    assert.equal(decision.status, 'skipped');
    assert.equal(decision.write, false);
    assert.equal(decision.handoff, false);
});

test('root-only changes publish but do not invoke an agent or listing PR', () => {
    const baseline = payload(), candidate = payload('2.3.0');
    edit(candidate, 'README.md', 'New marketplace instructions');
    const comparison = compareTrees(baseline, candidate);
    assert.deepEqual(comparison.changedPlugins, []);
    assert.deepEqual(comparison.changedPaths, ['README.md']);
    assert.deepEqual(publicationDecision(comparison, {
        currentVersion: '2.0.1', version: '2.3.0', tagExists: false,
    }), { status: 'published', write: true, handoff: false });
});

test('each real distributed surface, added/removed files and launch versions are meaningful', () => {
    for (const path of [
        'README.md', 'plugin.json', 'mcp.json', 'bin/helper.ps1',
        'skills/excel-mcp/SKILL.md', 'skills/excel-mcp/references/range.md', 'assets/icon.png',
    ]) {
        const baseline = payload(), candidate = payload('2.3.0');
        const fullPath = `plugins/excel-mcp/${path}`;
        if (path === 'plugin.json') {
            const manifest = JSON.parse(candidate.get(fullPath).bytes);
            manifest.description = 'New description 9.9.9';
            edit(candidate, fullPath, JSON.stringify(manifest));
        } else if (path === 'mcp.json') edit(candidate, fullPath, '{"args":["npx","@2.3.0"]}');
        else edit(candidate, fullPath, 'Changed functional text 9.9.9');
        const comparison = compareTrees(baseline, candidate);
        assert.deepEqual(comparison.changedPlugins, ['excel-mcp'], path);
        assert.ok(comparison.changedPaths.includes(fullPath), path);
        assert.equal(publicationDecision(comparison, {
            currentVersion: '2.0.1', version: '2.3.0', tagExists: false,
        }).handoff, true);
        candidate.delete(fullPath);
        if (path.includes('helper') || path.includes('assets')) {
            edit(baseline, fullPath, 'removed');
            assert.ok(compareTrees(baseline, candidate).changedPaths.includes(fullPath));
        }
    }
});

test('canonical JSON preserves values and array order but not object order or formatting', () => {
    assert.equal(canonicalJson('{"b":2,"a":{"z":1,"y":true}}'), canonicalJson(' { "a": { "y":true, "z":1 }, "b":2 } '));
    assert.notEqual(canonicalJson('{"args":["a","b"]}'), canonicalJson('{"args":["b","a"]}'));
    const baseline = payload(), candidate = payload('2.3.0');
    const path = 'plugins/excel-mcp/plugin.json';
    edit(candidate, path, JSON.stringify(Object.fromEntries(Object.entries(JSON.parse(candidate.get(path).bytes)).reverse()), null, 4));
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
    for (const [path, value] of [
        ['plugins/excel-mcp/plugin.json', '{}'],
        ['plugins/excel-mcp/plugin.json', '{invalid'],
        ['plugins/excel-mcp/version.txt', '2.0.2'],
        ['plugins/excel-mcp/skills/excel-mcp/VERSION', null],
        ['plugins/excel-mcp/mcp.json', null],
        ['plugins/unknown/plugin.json', '{}'],
        ['plugins/excel-mcp/binary.dll', 'runtime'],
    ]) {
        const candidate = payload();
        if (value === null) candidate.delete(path); else edit(candidate, path, value);
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

test('manual catch-up compares actually listed content, not the previous release or newest product tag', () => {
    assert.equal(plan().action, 'noop');
    const candidate = payload('2.3.0');
    edit(candidate, 'plugins/excel-cli/README.md', 'Accumulated old change');
    const result = plan(candidate);
    assert.equal(result.action, 'create');
    assert.deepEqual(result.changedPlugins, ['excel-cli']);
    assert.equal(result.state.entries['excel-cli'].entry.source.sha, sha2);
    assert.equal(result.state.entries['excel-cli'].entry.source.ref, 'v2.3.0');
    assert.equal(result.state.entries['excel-cli'].entry.description, 'Curated description');
});

test('both changed plugins produce one combined proposal; relevant manifest metadata is reflected', () => {
    const candidate = payload('2.3.0');
    for (const name of ['excel-cli', 'excel-mcp']) {
        edit(candidate, `plugins/${name}/README.md`, 'Changed');
        const manifest = JSON.parse(candidate.get(`plugins/${name}/plugin.json`).bytes);
        manifest.description = `${name} new metadata`;
        edit(candidate, `plugins/${name}/plugin.json`, JSON.stringify(manifest));
    }
    const result = plan(candidate);
    assert.equal(result.action, 'create');
    assert.deepEqual(result.changedPlugins, ['excel-cli', 'excel-mcp']);
    assert.equal(result.state.entries['excel-mcp'].entry.description, 'excel-mcp new metadata');
    assert.equal(parseState(pending(result).body).proposalFingerprint, result.state.proposalFingerprint);
});

test('pending equivalent proposals do not write, even when candidate stamps differ', () => {
    const candidate = payload('2.3.0');
    edit(candidate, 'plugins/excel-cli/README.md', 'Changed');
    const prior = pending(plan(candidate));
    const next = payload('2.4.0');
    edit(next, 'plugins/excel-cli/README.md', 'Changed');
    const result = plan(next, { pulls: [prior], tag: 'v2.4.0', commit: sha4,
        getTree: sha => ({ [sha1]: payload(), [sha2]: candidate, [sha4]: next })[sha],
        resolveTag: tag => ({ 'v2.0.1': sha1, 'v2.3.0': sha2, 'v2.4.0': sha4 })[tag] });
    assert.equal(result.action, 'noop');
    assert.equal(result.reason, 'Pending proposal already contains equivalent content.');
});

test('new content refreshes the same PR and preserves an existing proposal for the other plugin', () => {
    const first = payload('2.3.0');
    edit(first, 'plugins/excel-cli/README.md', 'CLI proposal');
    const prior = pending(plan(first));
    const second = payload('2.4.0');
    edit(second, 'plugins/excel-mcp/README.md', 'MCP proposal');
    const result = plan(second, { pulls: [prior], tag: 'v2.4.0', commit: sha4,
        getTree: sha => ({ [sha1]: payload(), [sha2]: first, [sha4]: second })[sha],
        resolveTag: tag => ({ 'v2.0.1': sha1, 'v2.3.0': sha2, 'v2.4.0': sha4 })[tag] });
    assert.equal(result.action, 'update');
    assert.equal(result.pullNumber, 42);
    assert.equal(result.expectedHead, sha3);
    assert.equal(result.state.entries['excel-cli'].entry.source.sha, sha2);
    assert.deepEqual(Object.keys(result.state.entries).sort(), ['excel-cli', 'excel-mcp']);
});

test('human changes, duplicate owned PRs, wrong ownership and declined identical proposals block', () => {
    const candidate = payload('2.3.0');
    edit(candidate, 'plugins/excel-cli/README.md', 'Changed');
    const prior = pending(plan(candidate));
    assert.throws(() => plan(candidate, { pulls: [{ ...prior, head: { ...prior.head, sha: sha2 } }] }), /head/);
    assert.throws(() => plan(candidate, { pulls: [prior, { ...prior, number: 43 }] }), /one open/);
    assert.throws(() => plan(candidate, { pulls: [{ ...prior, user: { login: 'human' } }] }), /ownership/);
    assert.throws(() => plan(candidate, { pulls: [{ ...prior, head: { ...prior.head, repo: null } }] }), /ownership/);
    assert.throws(() => plan(candidate, { pulls: [{ ...prior, state: 'closed' }] }), /declined/);
    const bodyState = parseState(prior.body);
    bodyState.bodyFingerprint = '0'.repeat(64);
    assert.throws(() => parseState('Human changes\n' + stateMarker(bodyState)), /body changed/);
});

test('owned bodies cannot bypass protection by removing or corrupting their fingerprint', () => {
    const candidate = payload('2.3.0');
    edit(candidate, 'plugins/excel-cli/README.md', 'Changed');
    const prior = pending(plan(candidate));
    for (const value of [undefined, null, 'invalid']) {
        const state = parseState(prior.body);
        if (value === undefined) delete state.bodyFingerprint;
        else state.bodyFingerprint = value;
        assert.throws(() => plan(candidate, { pulls: [{ ...prior, body: 'Human edits\n' + stateMarker(state) }] }), /fingerprint/);
    }
});

test('an orphan from an older proposal blocks a different create branch but associated closed heads do not', () => {
    const candidate = payload('2.3.0');
    edit(candidate, 'plugins/excel-cli/README.md', 'Changed');
    const proposed = plan(candidate);
    const oldBranch = 'excel-plugin-updates-123456abcdef';
    const heads = `${sha3}\trefs/heads/${oldBranch}\n`;
    assert.throws(() => updater.assertForkHeads(proposed, heads, []), /orphan/);
    assert.doesNotThrow(() => updater.assertForkHeads(proposed, heads, [pending(proposed, { state: 'closed' })]));
    assert.throws(() => updater.assertForkHeads(proposed, 'invalid API data', []), /Invalid fork/);
});

test('orphan association includes paginated open and historical closed PR data', () => {
    const candidate = payload('2.3.0');
    edit(candidate, 'plugins/excel-cli/README.md', 'Changed');
    const proposed = plan(candidate);
    const visited = [];
    const pulls = updater.allPulls(endpoint => {
        visited.push(endpoint);
        if (endpoint.includes('state=open')) {
            if (endpoint.endsWith('page=2')) return [];
            return Array.from({ length: 100 }, (_, i) => ({
                number: i + 1, state: 'open', body: null, user: { login: 'another-user' },
            }));
        }
        if (endpoint.includes('state=closed')) return [];
        if (endpoint.startsWith('search/issues?')) {
            const last = endpoint.endsWith('page=2');
            return { incomplete_results: false, total_count: 101,
                items: Array.from({ length: last ? 1 : 100 }, (_, i) => ({ number: 1001 + i + (last ? 100 : 0) })) };
        }
        const number = Number(endpoint.split('/').at(-1));
        return pending(proposed, { number, state: 'closed', merged_at: '2026-01-01T00:00:00Z' });
    });
    assert.equal(pulls.length, 201);
    assert.ok(visited.some(endpoint => endpoint.includes('state=open') && endpoint.endsWith('page=2')));
    assert.ok(visited.some(endpoint => endpoint.startsWith('search/issues?') && endpoint.endsWith('page=2')));
    assert.ok(pulls.some(pr => pr.number === 1101));
    assert.doesNotThrow(() => updater.assertForkHeads(proposed, `${sha3}\trefs/heads/excel-plugin-updates-123456abcdef\n`, pulls));
});

test('missing search pages or invalid counts cannot silently omit historical owned PRs', () => {
    for (const response of [
        { incomplete_results: false, total_count: -1, items: [] },
        { incomplete_results: false, total_count: 1, items: [] },
        { incomplete_results: false, total_count: 101, items: [] },
    ]) {
        assert.throws(() => updater.allPulls(endpoint =>
            endpoint.startsWith('search/issues?') ? response : []), /API data|Truncated/);
    }
});

test('same-PR refresh tolerates head propagation and reconciles a body write whose response was lost', () => {
    const before = 'Original protected body';
    const after = 'New protected body';
    const proposed = { expectedHead: sha1, expectedBody: hash(before), pullNumber: 42 };
    const observations = [
        { state: 'open', head: { sha: sha1 }, body: before },
        { state: 'open', head: { sha: sha2 }, body: before },
        { state: 'open', head: { sha: sha2 }, body: after, html_url: 'https://github.com/github/awesome-copilot/pull/42' },
    ];
    let writes = 0, waits = 0;
    const result = updater.refreshPullAfterPush(proposed, sha2, after, {
        readPull: () => observations.shift(), readHead: () => sha2,
        writeBody: () => { writes++; throw new Error('Lost response'); }, wait: () => { waits++; },
    });
    assert.equal(result.body, after);
    assert.equal(writes, 1);
    assert.equal(waits, 2);
});

test('post-push refresh rejects unrelated edits and stops after bounded propagation retries', () => {
    const before = 'Protected body';
    const proposed = { expectedHead: sha1, expectedBody: hash(before), pullNumber: 42 };
    for (const live of [
        { state: 'closed', head: { sha: sha2 }, body: before },
        { state: 'open', head: { sha: sha4 }, body: before },
        { state: 'open', head: { sha: sha2 }, body: 'Human edit' },
    ]) {
        assert.throws(() => updater.refreshPullAfterPush(proposed, sha2, 'New body', {
            readPull: () => live, readHead: () => sha2, writeBody: () => assert.fail('No write allowed'), wait: () => {},
        }), /changed after push/);
    }
    let reads = 0;
    assert.throws(() => updater.refreshPullAfterPush(proposed, sha2, 'New body', {
        readPull: () => { reads++; return { state: 'open', head: { sha: sha1 }, body: before }; },
        readHead: () => sha2, writeBody: () => assert.fail('No premature write'), wait: () => {},
    }), /reconciliation/);
    assert.equal(reads, 8);
    assert.throws(() => updater.refreshPullAfterPush(proposed, sha2, 'New body', {
        readPull: () => assert.fail('No API write/read after foreign push'), readHead: () => sha4,
        writeBody: () => assert.fail('No overwrite'), wait: () => {},
    }), /Fork head changed/);
});

test('a partial body-write failure retries only the same protected transition', () => {
    const before = 'Original protected body', after = 'Verified new body';
    const proposed = { expectedHead: sha1, expectedBody: hash(before), pullNumber: 42 };
    let writes = 0;
    const result = updater.refreshPullAfterPush(proposed, sha2, after, {
        readPull: () => ({ state: 'open', head: { sha: sha2 }, body: before }),
        readHead: () => sha2, wait: () => {},
        writeBody: () => {
            if (++writes === 1) throw new Error('Transient API failure');
            return { state: 'open', head: { sha: sha2 }, body: after };
        },
    });
    assert.equal(result.body, after);
    assert.equal(writes, 2);
});

test('refresh cannot report success after the fork changes during a body write', () => {
    const before = 'Protected body', after = 'New body';
    const proposed = { expectedHead: sha1, expectedBody: hash(before), pullNumber: 42 };
    let head = sha2;
    assert.throws(() => updater.refreshPullAfterPush(proposed, sha2, after, {
        readPull: () => ({ state: 'open', head: { sha: sha2 }, body: before }),
        readHead: () => head, wait: () => {},
        writeBody: () => {
            head = sha4;
            return { state: 'open', head: { sha: sha2 }, body: after };
        },
    }), /Fork head changed/);
});

test('submission receipt records only the exact transition and public proposal, never credentials', () => {
    const files = { 'plugins/external.json': '[]\n' };
    const proposed = { action: 'update', state: { base: sha1 }, branch: 'excel-plugin-updates-123456abcdef',
        pullNumber: 42, expectedHead: sha2, expectedBody: hash('Original body'), files,
        token: 'must-not-be-copied', env: { AWESOME_COPILOT_PR_TOKEN: 'must-not-be-copied' } };
    const receipt = updater.createSubmissionReceipt(proposed, sha3, 'Exact new body');
    assert.deepEqual(receipt, { schema: 1, status: 'prepared', action: 'update',
        upstream: 'github/awesome-copilot', fork: 'sbroenne/awesome-copilot', base: sha1,
        branch: proposed.branch, pullNumber: 42, previousHead: sha2,
        expectedBody: proposed.expectedBody, head: sha3, body: 'Exact new body', files });
    assert.ok(!JSON.stringify(receipt).includes('must-not-be-copied'));
});

test('invalid API/listing/tag data and listing/pending downgrades fail instead of no-change', () => {
    assert.throws(() => plan(undefined, { pulls: null }), /API/);
    assert.throws(() => plan(undefined, { tag: 'main' }), /tag/);
    assert.throws(() => plan(undefined, { resolveTag: () => null }), /tag/);
    const bad = listings();
    bad[0].source.sha = sha3;
    assert.throws(() => plan(undefined, { listings: bad }), /ref\/SHA/);
    const newer = listings();
    newer[0].version = '2.4.0';
    assert.throws(() => plan(undefined, { listings: newer }), /Downgrade/);
    assert.throws(() => plan(undefined, { listings: [] }), /exactly/);
    const candidate = payload('2.3.0');
    edit(candidate, 'plugins/excel-cli/README.md', 'Changed');
    const prior = pending(plan(candidate));
    const state = parseState(prior.body);
    const badLocator = structuredClone(state);
    badLocator.entries['excel-cli'].entry.source.sha = sha3;
    assert.throws(() => plan(candidate, { pulls: [{ ...prior, body: stateMarker(badLocator) }] }), /locator/);
    state.entries['excel-cli'].entry.version = '2.4.0';
    state.entries['excel-cli'].entry.source.ref = 'v2.4.0';
    prior.body = stateMarker(state);
    assert.throws(() => plan(candidate, { pulls: [prior] }), /Downgrade/);
});

test('patch is limited to exactly affected entries and the two approved files', () => {
    const before = [...listings(), { name: 'unrelated', source: 'keep' }];
    const after = structuredClone(before);
    after[0].version = '2.3.0';
    assert.doesNotThrow(() => assertListingPatch(before, after, ['excel-cli']));
    after[1].description = 'not allowed';
    assert.throws(() => assertListingPatch(before, after, ['excel-cli']), /unaffected/);
    after.pop();
    assert.throws(() => assertListingPatch(before, after, ['excel-cli']), /entries/);
    assert.doesNotThrow(() => assertAllowedPaths(['plugins/external.json', '.github/plugin/marketplace.json']));
    assert.throws(() => assertAllowedPaths(['.github/workflows/evil.yml']), /allowed/);
    assert.throws(() => assertAllowedPaths([]), /empty/);
});

test('privileged writer rejects multiple/extra/schema-invalid outputs and honors trusted opt-in/preview/staged gates', () => {
    const trusted = { guardFingerprint: 'verified' };
    const item = { type: 'submit_marketplace_update', proposal_fingerprint: 'verified', body: 'A reviewed proposal body with at least forty characters.' };
    assert.doesNotThrow(() => validateRequest({ items: [item] }, trusted));
    for (const items of [[], [item, item], [{ ...item, repo: 'evil/repo' }], [{ ...item, proposal_fingerprint: 'wrong' }]]) {
        assert.throws(() => validateRequest({ items }, trusted));
    }
    for (const env of [
        {}, { PREVIEW: 'true', AWESOME_COPILOT_UPDATES_ENABLED: 'true' },
        { PREVIEW: 'false', AWESOME_COPILOT_UPDATES_ENABLED: 'false' },
        { PREVIEW: 'false', AWESOME_COPILOT_UPDATES_ENABLED: 'true', GH_AW_SAFE_OUTPUTS_STAGED: 'true' },
    ]) {
        assert.equal(writesEnabled(env), false);
        assert.equal(submit({ trustedPlan: trusted, output: { items: [item] }, env }).status, 'preview');
    }
    assert.equal(writesEnabled({ PREVIEW: 'false', AWESOME_COPILOT_UPDATES_ENABLED: 'true' }), true);
});

test('PR bodies preserve upstream headings, comments, checklist items and ordering', () => {
    const template = '## Checklist\n- [ ] Follow guidelines.\n<!-- Keep context -->\n## Description\n';
    assert.doesNotThrow(() => assertTemplate(template, template.replace('[ ]', '[x]') + 'Actual changes.'));
    for (const body of [
        template.replace('<!-- Keep context -->', ''),
        template.replace('- [ ] Follow guidelines.\n', ''),
        template.replace('## Checklist', '## Changed heading'),
        '## Description\n## Checklist\n- [ ] Follow guidelines.\n<!-- Keep context -->',
    ]) assert.throws(() => assertTemplate(template, body), /preserve upstream template/);
});

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
test('compiled updater isolates writes, blocks failed detection and never creates fallback issues', () => {
    const workflow = fs.readFileSync(path.join(repoRoot, '.github', 'workflows', 'update-awesome-copilot.lock.yml'), 'utf8');
    const agent = workflow.split('\n  agent:')[1].split('\n  conclusion:')[0];
    const writer = workflow.split('\n  submit_marketplace_update:')[1];
    assert.equal(agent.includes('secrets.AWESOME_COPILOT_PR_TOKEN'), false);
    assert.equal((workflow.match(/secrets\.AWESOME_COPILOT_PR_TOKEN/g) ?? []).length, 1);
    assert.match(writer, /needs\.agent\.result == 'success' && needs\.detection\.result == 'success'/);
    assert.match(writer, /needs\.detection\.outputs\.detection_success == 'true'/);
    assert.match(workflow, /GH_AW_DETECTION_CONTINUE_ON_ERROR: "false"/);
    assert.equal(workflow.includes('issues: write'), false);
    assert.equal(workflow.includes('GH_AW_MISSING_TOOL_CREATE_ISSUE: "true"'), false);
    assert.equal(workflow.includes('GH_AW_MISSING_DATA_CREATE_ISSUE: "true"'), false);
    assert.match(workflow, /github\/gh-aw\/actions\/setup@c35393777e5604a63721d09512263b1383301d4f/);
});

test('automatic updater requires a meaningful publication handoff and product release does not depend on it', () => {
    const publisher = fs.readFileSync(path.join(repoRoot, '.github', 'workflows', 'publish-plugins.yml'), 'utf8');
    const release = fs.readFileSync(path.join(repoRoot, '.github', 'workflows', 'release.yml'), 'utf8');
    assert.match(publisher, /if: needs\.publish\.outputs\.handoff == 'true' && vars\.AWESOME_COPILOT_UPDATES_ENABLED == 'true'/);
    assert.match(publisher, /published_tag: \$\{\{ needs\.publish\.outputs\.published_tag \}\}/);
    assert.match(release, /publish-plugins:\r?\n\s+needs: \[version, create-tag, create-release, publish\]/);
    assert.equal(release.slice(0, release.indexOf('\n  publish-plugins:')).includes('needs: publish-plugins'), false);
});

test('plugin-only hook plans build and select publication regressions without requiring Excel', () => {
    const script = ". .\\scripts\\Get-ValidationPlan.ps1; @('scripts/Publish-PreparedPlugins.ps1','scripts/PluginContent.mjs','.github/workflows/update-awesome-copilot.md','.github/workflows/publish-plugins.yml') | ForEach-Object { Get-ValidationPlan -Paths @($_) } | ConvertTo-Json -Compress";
    const plans = JSON.parse(command('pwsh', ['-NoProfile', '-Command', script], repoRoot));
    for (const result of plans) {
        assert.equal(result.Build, true);
        assert.equal(result.Plugins, true);
        assert.equal(result.Excel, false);
    }
    const runner = fs.readFileSync(path.join(repoRoot, 'scripts', 'Invoke-ExcelFreeTests.ps1'), 'utf8');
    assert.match(runner, /publish-plugins/);
});

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
    return execFileSync(executable, args, { cwd, encoding: 'utf8', windowsHide: true,
        env });
}

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

function publicationFixture({
    staleOverlay = false, currentVersion = '2.0.1', candidateVersion = '2.3.0',
    payloadVersion = candidateVersion, syncFails = false, candidateTagExists = false, ignoredNames = false,
} = {}) {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-publication-'));
    const source = path.join(root, 'source'), output = path.join(root, 'output'), built = path.join(root, 'built');
    init(source);
    init(output);
    fs.mkdirSync(path.join(source, 'scripts'));
    fs.mkdirSync(path.join(source, '.github', 'plugins', 'marketplace-repo'), { recursive: true });
    fs.writeFileSync(path.join(source, '.github', 'plugins', 'marketplace-repo', 'README.md'), 'Marketplace');
    if (staleOverlay) fs.writeFileSync(path.join(source, '.github', 'plugins', 'marketplace-repo', 'stale.txt'), 'Old owned file');
    fs.writeFileSync(path.join(source, 'scripts', 'Sync-PublishedPluginRepo.ps1'), syncFails ? "throw 'sync-root-cause'" : `
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
    const originalSource = commit(source, 'Original source');
    command('git', ['tag', `v${currentVersion}`], source);
    writeTree(output, payload(currentVersion));
    if (ignoredNames) fs.writeFileSync(path.join(output, '.gitignore'), '*.test.md\n*.log\ntemp/\n');
    if (staleOverlay) fs.writeFileSync(path.join(output, 'stale.txt'), 'Old owned file');
    const original = commit(output, 'Original publication');
    command('git', ['tag', `v${currentVersion}`], output);
    if (candidateTagExists && candidateVersion !== currentVersion) command('git', ['tag', `v${candidateVersion}`], output);
    fs.writeFileSync(path.join(source, 'release.txt'), candidateVersion);
    if (staleOverlay) fs.unlinkSync(path.join(source, '.github', 'plugins', 'marketplace-repo', 'stale.txt'));
    const preparedSource = commit(source, 'Candidate source');
    const sourceCommit = candidateVersion === currentVersion ? originalSource : preparedSource;
    if (candidateVersion !== currentVersion) command('git', ['tag', `v${candidateVersion}`], source);
    command('git', ['checkout', '--quiet', '--detach', sourceCommit], source);
    const candidate = payload(payloadVersion);
    for (const [name, file] of candidate) if (name.startsWith('plugins/')) {
        writeTree(built, new Map([[name.slice('plugins/'.length), file]]));
    }
    return { root, source, output, built, original, sourceCommit };
}
function publishFixture(fixture, flags = [], version = '2.3.0', sourceCommit = fixture.sourceCommit) {
    return command('pwsh', ['-NoProfile', '-File', path.join(repoRoot, 'scripts', 'Publish-PreparedPlugins.ps1'),
        '-SourceDirectory', fixture.source, '-PublishedRepoDirectory', fixture.output,
        '-BuiltPluginsDirectory', fixture.built, '-Version', version, '-SourceCommit', sourceCommit, ...flags], repoRoot);
}

test('publisher actually creates no destination commit/push/tag for version-only releases', () => {
    const fixture = publicationFixture();
    try {
        const text = publishFixture(fixture);
        assert.match(text, /"status": "skipped"/);
        assert.match(text, /"published_tag": "v2.0.1"/);
        assert.equal(command('git', ['rev-parse', 'HEAD'], fixture.output).trim(), fixture.original);
        assert.equal(command('git', ['tag', '--list', 'v2.3.0'], fixture.output).trim(), '');
        assert.equal(command('git', ['status', '--porcelain'], fixture.output).trim(), '');
    } finally { fs.rmSync(fixture.root, { recursive: true }); }
});

test('publisher rejects a dirty or wrong source checkout before touching the destination', () => {
    const fixture = publicationFixture();
    try {
        fs.writeFileSync(path.join(fixture.source, 'release.txt'), 'Unreleased change');
        assert.throws(() => publishFixture(fixture), /Source checkout must be clean/);
        commit(fixture.source, 'Unreleased source');
        assert.throws(() => publishFixture(fixture), /exact release commit/);
        assert.equal(command('git', ['rev-parse', 'HEAD'], fixture.output).trim(), fixture.original);
        assert.equal(command('git', ['status', '--porcelain'], fixture.output).trim(), '');
    } finally { fs.rmSync(fixture.root, { recursive: true }); }
});

if (process.env.PLUGIN_PUBLICATION_SCENARIO) {
    test('legacy publication guard scenario', () => {
        const scenario = JSON.parse(process.env.PLUGIN_PUBLICATION_SCENARIO);
        const fixture = publicationFixture({
            currentVersion: scenario.publishedVersion, candidateVersion: '1.2.3',
            payloadVersion: scenario.payloadVersion, syncFails: scenario.syncFails,
            candidateTagExists: scenario.tagExists,
        });
        try {
            const remote = path.join(fixture.root, 'remote.git');
            command('git', ['clone', '--quiet', '--bare', fixture.output, remote], fixture.root);
            command('git', ['remote', 'add', 'origin', remote], fixture.output);
            if (!scenario.tagExists && !scenario.manualRepair) {
                fs.writeFileSync(path.join(fixture.built, 'excel-cli', 'README.md'), 'Meaningful candidate guidance');
            }
            const action = () => publishFixture(fixture, scenario.manualRepair ? ['-ManualRepair'] : [], '1.2.3');
            if (scenario.succeeds) {
                const result = action();
                assert.match(result, /"status": "(skipped|published)"/);
                if (scenario.tagExists && !scenario.manualRepair) {
                    assert.match(result, /"status": "skipped"/);
                    assert.equal(command('git', ['rev-parse', 'HEAD'], fixture.output).trim(), fixture.original);
                }
                assert.equal(command('git', ['rev-parse', `v${scenario.publishedVersion}^{commit}`], fixture.output).trim(), fixture.original);
            } else {
                const expected = scenario.publishedVersion === '1.2.4' ? /Downgrade/
                    : scenario.tagExists ? /Existing tag conflicts/
                    : scenario.payloadVersion !== '1.2.3' ? /manifest version/
                    : /sync-root-cause/;
                assert.throws(action, expected);
                assert.equal(command('git', ['rev-parse', 'HEAD'], fixture.output).trim(), fixture.original);
                assert.equal(command('git', ['status', '--porcelain'], fixture.output).trim(), '');
            }
        } finally { fs.rmSync(fixture.root, { recursive: true }); }
    });
}

test('real publication after skipped versions pushes and tags only a disposable local remote', () => {
    const fixture = publicationFixture();
    try {
        publishFixture(fixture);
        const remote = path.join(fixture.root, 'remote.git');
        command('git', ['clone', '--quiet', '--bare', fixture.output, remote], fixture.root);
        command('git', ['remote', 'add', 'origin', remote], fixture.output);
        fs.writeFileSync(path.join(fixture.built, 'excel-cli', 'README.md'), 'New CLI guidance');
        const result = publishFixture(fixture);
        assert.match(result, /"status": "published"/);
        assert.match(result, /"handoff": true/);
        assert.match(result, /"published_tag": "v2.3.0"/);
        assert.notEqual(command('git', ['rev-parse', 'HEAD'], fixture.output).trim(), fixture.original);
        assert.equal(command('git', ['rev-parse', 'v2.0.1'], fixture.output).trim(), fixture.original);
        assert.equal(command('git', [`--git-dir=${remote}`, 'rev-parse', 'v2.3.0^{commit}'], fixture.root).trim(),
            command('git', ['rev-parse', 'HEAD'], fixture.output).trim());
    } finally { fs.rmSync(fixture.root, { recursive: true }); }
});

test('publication includes ignored-name additions and commits the exact prepared distributed tree', () => {
    const fixture = publicationFixture({ ignoredNames: true });
    try {
        const remote = path.join(fixture.root, 'remote.git');
        command('git', ['clone', '--quiet', '--bare', fixture.output, remote], fixture.root);
        command('git', ['remote', 'add', 'origin', remote], fixture.output);
        const name = 'plugins/excel-cli/skills/excel-cli/references/guide.test.md';
        const content = 'Required functional content\r\n';
        fs.writeFileSync(path.join(fixture.built, name.slice('plugins/'.length)), content);
        fs.writeFileSync(path.join(fixture.output, 'private.log'), 'Unrelated local note');
        const result = publishFixture(fixture);
        assert.match(result, /"status": "published"/);
        assert.equal(command('git', ['show', `v2.3.0:${name}`], fixture.output), 'Required functional content\n');
        assert.equal(command('git', [`--git-dir=${remote}`, 'show', `v2.3.0:${name}`], fixture.root), 'Required functional content\n');
        assert.equal(command('git', ['ls-tree', '--name-only', 'v2.3.0', '--', 'private.log'], fixture.output), '');
        assert.equal(fs.readFileSync(path.join(fixture.output, 'private.log'), 'utf8'), 'Unrelated local note');
    } finally { fs.rmSync(fixture.root, { recursive: true }); }
});

test('publisher preview of real root changes never writes to destination and removes stale source overlay files', () => {
    const fixture = publicationFixture({ staleOverlay: true });
    try {
        fs.writeFileSync(path.join(fixture.source, '.github', 'plugins', 'marketplace-repo', 'README.md'), 'New root documentation');
        const sourceCommit = commit(fixture.source, 'New overlay');
        command('git', ['tag', 'v2.4.0'], fixture.source);
        for (const name of ['excel-cli', 'excel-mcp']) {
            const manifest = path.join(fixture.built, name, 'plugin.json');
            const value = JSON.parse(fs.readFileSync(manifest));
            value.version = '2.4.0';
            fs.writeFileSync(manifest, JSON.stringify(value));
            fs.writeFileSync(path.join(fixture.built, name, 'version.txt'), '2.4.0');
            fs.writeFileSync(path.join(fixture.built, name, 'skills', name, 'VERSION'), '2.4.0');
        }
        const result = publishFixture(fixture, ['-Preview'], '2.4.0', sourceCommit);
        assert.match(result, /"decision": "published"/);
        assert.match(result, /"changed_plugins": \[\]/);
        assert.match(result, /"handoff": false/);
        assert.match(result, /stale.txt/);
        assert.equal(command('git', ['rev-parse', 'HEAD'], fixture.output).trim(), fixture.original);
        assert.equal(command('git', ['tag', '--list', 'v2.4.0'], fixture.output).trim(), '');
    } finally { fs.rmSync(fixture.root, { recursive: true }); }
});

test('authorized manual repair restores missing stamps/tag and preserves an existing immutable tag', () => {
    const fixture = publicationFixture();
    try {
        const remote = path.join(fixture.root, 'remote.git');
        command('git', ['clone', '--quiet', '--bare', fixture.output, remote], fixture.root);
        command('git', ['remote', 'add', 'origin', remote], fixture.output);
        for (const name of ['excel-cli', 'excel-mcp']) {
            const file = path.join(fixture.built, name, 'plugin.json');
            const manifest = JSON.parse(fs.readFileSync(file));
            manifest.version = '2.0.1';
            fs.writeFileSync(file, JSON.stringify(manifest));
            fs.writeFileSync(path.join(fixture.built, name, 'version.txt'), '2.0.1');
            fs.writeFileSync(path.join(fixture.built, name, 'skills', name, 'VERSION'), '2.0.1');
        }
        const exactSource = command('git', ['rev-parse', 'v2.0.1'], fixture.source).trim();
        command('git', ['checkout', '--quiet', '--detach', exactSource], fixture.source);
        fs.unlinkSync(path.join(fixture.output, 'plugins', 'excel-cli', 'skills', 'excel-cli', 'VERSION'));
        commit(fixture.output, 'Simulate missing known stamp');
        const repaired = publishFixture(fixture, ['-ManualRepair'], '2.0.1', exactSource);
        assert.match(repaired, /"status": "published"/);
        assert.match(repaired, /"handoff": false/);
        assert.equal(command('git', ['rev-parse', 'v2.0.1'], fixture.output).trim(), fixture.original);
        assert.equal(fs.readFileSync(path.join(fixture.output, 'plugins', 'excel-cli', 'skills', 'excel-cli', 'VERSION'), 'utf8'), '2.0.1');
        command('git', ['tag', '--delete', 'v2.0.1'], fixture.output);
        command('git', ['push', 'origin', ':refs/tags/v2.0.1'], fixture.output);
        const restored = publishFixture(fixture, ['-ManualRepair'], '2.0.1', exactSource);
        assert.match(restored, /"status": "published"/);
        assert.equal(command('git', ['tag', '--list', 'v2.0.1'], fixture.output).trim(), 'v2.0.1');
    } finally { fs.rmSync(fixture.root, { recursive: true }); }
});
