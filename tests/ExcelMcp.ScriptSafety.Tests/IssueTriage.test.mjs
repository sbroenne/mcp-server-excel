import assert from 'node:assert/strict';
import { test } from 'node:test';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { prepare, publish, dispatch, validateResult, ownedState, assignmentBody, main, api as githubApi } from '../../scripts/IssueTriage.mjs';
import os from 'node:os';

const repository = 'sbroenne/mcp-server-excel';
const bot = { id: 41898282, login: 'github-actions[bot]', type: 'Bot' };
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const sourcePath = 'src/ExcelMcp.Core/Commands/Range/RangeCommands.cs';
const result = () => ({
    summary: 'A localized range error needs correction.',
    labels: ['bug'], duplicates: [], missing_details: [],
    easy_fix: { eligible: true, reason: 'The failing condition is clear.', paths: [sourcePath],
        validation: ['Add a failing regression test and run Excel E2E.'] },
});

function fixture() {
    const state = {
        issue: { number: 42, title: 'Range bug', body: 'A small reproducible defect.', state: 'open',
            locked: false, labels: [], assignees: [], user: { login: 'reporter' } },
        comments: [], pulls: [], calls: [], files: new Set([sourcePath]),
    };
    const api = async (endpoint, options = {}) => {
        const { method = 'GET', body } = options;
        state.calls.push({ endpoint, method, body });
        if (endpoint === 'user') return { login: 'sbroenne' };
        if (endpoint === `repos/${repository}`) return { full_name: repository, default_branch: 'main' };
        if (endpoint === `repos/${repository}/commits/main`) return { sha: 'a'.repeat(40) };
        if (endpoint === `repos/${repository}/issues/42` && method === 'GET') return structuredClone(state.issue);
        if (endpoint.startsWith(`repos/${repository}/labels?`)) return [{ name: 'bug' }, { name: 'question' }];
        if (endpoint.startsWith(`repos/${repository}/pulls?`)) return structuredClone(state.pulls);
        if (endpoint.startsWith(`repos/${repository}/contents/`)) {
            const file = decodeURIComponent(endpoint.split('/contents/')[1].split('?')[0]);
            if (!state.files.has(file)) throw new Error('Missing source file.');
            return { type: 'file' };
        }
        if (endpoint.startsWith(`repos/${repository}/issues/42/comments?`)) {
            const page = Number(new URLSearchParams(endpoint.split('?')[1]).get('page'));
            return structuredClone(state.comments.slice((page - 1) * 100, page * 100));
        }
        if (endpoint === `repos/${repository}/issues/42/comments` && method === 'POST') {
            const comment = { id: state.comments.length + 1, user: bot, body: body.body };
            state.comments.push(comment);
            return structuredClone(comment);
        }
        const commentMatch = endpoint.match(/\/issues\/comments\/(\d+)$/);
        if (commentMatch && method === 'PATCH') {
            const comment = state.comments.find(item => item.id === Number(commentMatch[1]));
            assert.ok(comment, 'Only an existing comment may be edited.');
            comment.body = body.body;
            return structuredClone(comment);
        }
        if (endpoint === `repos/${repository}/issues/42/assignees` && method === 'POST') {
            state.issue.assignees.push({ login: 'Copilot', id: 198982749, type: 'Bot' });
            return structuredClone(state.issue);
        }
        if (endpoint === `repos/${repository}/issues/7`) return { number: 7, state: 'closed', title: 'Similar issue' };
        throw new Error(`Unexpected API call: ${method} ${endpoint}`);
    };
    return { state, api };
}

async function ready() {
    const fixtureState = fixture();
    const snapshot = await prepare('42', fixtureState.api);
    return { ...fixtureState, snapshot };
}

test('automatic and manual issue numbers use the same validated target', async () => {
    const { api } = fixture();
    assert.equal((await prepare(42, api)).issue_number, 42);
    assert.equal((await prepare('42', api)).issue_number, 42);
    for (const invalid of ['', '0', '-1', '4.2', '42; echo injected', '9007199254740992']) {
        await assert.rejects(prepare(invalid, api), /issue number/i);
    }
});

test('PRs are rejected and closed or locked issues are visibly skipped', async () => {
    const { state, api } = fixture();
    state.issue.pull_request = {};
    await assert.rejects(prepare(42, api), /pull request/i);
    delete state.issue.pull_request;
    state.issue.state = 'closed';
    assert.equal((await prepare(42, api)).skip_reason, 'Issue is closed or locked.');
    state.issue.state = 'open';
    state.issue.locked = true;
    assert.equal((await prepare(42, api)).skip_reason, 'Issue is closed or locked.');
});

test('output rejects unknown fields, fake labels, PR duplicates, and unsafe fix scope', async () => {
    const { api, state } = fixture();
    await assert.rejects(validateResult({ ...result(), assign_to: 'attacker' }, api), /fields/i);
    await assert.rejects(validateResult({ ...result(), labels: ['invented'] }, api), /label/i);
    const duplicate = { ...result(), duplicates: [{ number: 7, reason: 'Same symptom.' }],
        easy_fix: { ...result().easy_fix, eligible: false } };
    assert.equal((await validateResult(duplicate, api)).duplicates[0].number, 7);
    const read = async (endpoint, options) => endpoint.endsWith('/issues/7') ?
        { number: 7, pull_request: {} } : api(endpoint, options);
    await assert.rejects(validateResult(duplicate, read), /duplicate/i);
    for (const unsafe of ['.github/workflows/ci.yml', 'docs/agents/development.md',
        'src/Auth.cs', '../README.md', 'src/Generated.g.cs', 'scripts/update.ps1']) {
        state.files.add(unsafe);
        const value = result();
        value.easy_fix.paths = [unsafe];
        await assert.rejects(validateResult(value, api), /scope/i);
    }
});

test('unclear reports and possible duplicates cannot be eligible easy fixes', async () => {
    const { api } = fixture();
    for (const extra of [{ missing_details: ['What reproduces it?'] },
        { duplicates: [{ number: 7, reason: 'Same condition.' }] }]) {
        await assert.rejects(validateResult({ ...result(), ...extra }, api), /easy fix/i);
    }
    const value = result();
    value.easy_fix.paths = [];
    await assert.rejects(validateResult(value, api), /scope/i);
});

test('publishing creates one bot-owned comment, updates it, and avoids unchanged writes', async () => {
    const { state, api, snapshot } = await ready();
    await publish(snapshot, result(), api);
    assert.equal(state.comments.length, 1);
    assert.match(state.comments[0].body, /Suggested labels:.*bug/);
    assert.match(state.comments[0].body, /No likely duplicates found/);
    assert.equal(ownedState(state.comments[0], 42).handoff, 'none');
    const writes = () => state.calls.filter(call => call.method !== 'GET').length;
    const initialWrites = writes();
    await publish(snapshot, result(), api);
    assert.equal(writes(), initialWrites);
    await publish(snapshot, { ...result(), summary: 'New diagnosis.' }, api);
    assert.equal(state.comments.length, 1);
    assert.match(state.comments[0].body, /New diagnosis/);
});

test('copied markers in human comments are not owned and discovery is paginated', async () => {
    const { state, api, snapshot } = await ready();
    await publish(snapshot, result(), api);
    const owned = structuredClone(state.comments[0]);
    state.comments = Array.from({ length: 100 }, (_, index) =>
        ({ id: index + 10, user: { id: 1, type: 'User' }, body: owned.body }));
    state.comments.push({ ...owned, id: 200 });
    const fresh = await prepare(42, api);
    await publish(fresh, { ...result(), summary: 'Paginated update.' }, api);
    assert.equal(state.comments.length, 101);
    assert.equal(state.comments[0].body, owned.body);
    assert.match(state.comments[100].body, /Paginated update/);
    assert.equal(ownedState(state.comments[0], 42), null);
});

test('stale issue content and failed reads cannot cause publication', async () => {
    const { state, api, snapshot } = await ready();
    state.issue.body = 'Changed during inference.';
    await assert.rejects(publish(snapshot, result(), api), /changed/i);
    assert.equal(state.comments.length, 0);
    await assert.rejects(prepare(42, async () => { throw new Error('API unavailable'); }), /API unavailable/);
});

test('auto-fix off switch does not prevent advisory publication', async () => {
    const { state, api, snapshot } = await ready();
    await publish(snapshot, result(), api);
    assert.equal((await dispatch(snapshot, result(), api, { enabled: false })).status, 'disabled');
    assert.equal(state.comments.length, 1);
    assert.equal(state.issue.assignees.length, 0);
});

test('assignment uses existing Squad guidance without modifying or loading Squad profiles', () => {
    const body = assignmentBody({ default_branch: 'main' }, result(), 'a'.repeat(40));
    assert.deepEqual(body.assignees, ['copilot-swe-agent[bot]']);
    assert.ok(!Object.hasOwn(body.agent_assignment, 'custom_agent'));
    assert.equal(body.agent_assignment.target_repo, repository);
    assert.equal(body.agent_assignment.base_branch, 'main');
    assert.match(body.agent_assignment.custom_instructions, /draft/i);
    assert.match(body.agent_assignment.custom_instructions, /Test-E2E/);
    assert.match(body.agent_assignment.custom_instructions, /\.squad\/routing\.md/);
    assert.match(body.agent_assignment.custom_instructions, /Do not change Squad/);
    assert.ok(body.agent_assignment.custom_instructions.includes(result().easy_fix.validation[0]));
    assert.ok(!Object.hasOwn(body.agent_assignment, 'model'));
});

test('successful dispatch persists state and repeated runs never assign twice', async () => {
    const { state, api, snapshot } = await ready();
    await publish(snapshot, result(), api);
    assert.equal((await dispatch(snapshot, result(), api, { enabled: true })).status, 'assigned');
    assert.equal(ownedState(state.comments[0], 42).handoff, 'assigned');
    const fresh = await prepare(42, api);
    assert.equal((await dispatch(fresh, result(), api, { enabled: true })).status, 'already-started');
    assert.equal(state.calls.filter(call => call.endpoint.endsWith('/assignees')).length, 1);
});

test('assigned issues, linked PRs and security-sensitive issues do not launch fixes', async () => {
    for (const scenario of ['assigned', 'pull', 'pull-title', 'security', 'security-label', 'security-comment']) {
        const { state, api, snapshot } = await ready();
        await publish(snapshot, result(), api);
        if (scenario === 'assigned') state.issue.assignees.push({ login: 'someone' });
        if (scenario === 'pull') state.pulls.push({ number: 88, body: 'Fixes #42', state: 'open' });
        if (scenario === 'pull-title') state.pulls.push({ number: 88, title: 'Fix #42', body: '', state: 'open' });
        if (scenario === 'security') state.issue.body = 'Fix authentication and credential validation.';
        if (scenario === 'security-label') state.issue.labels.push({ name: 'security' });
        if (scenario === 'security-comment') state.comments.push({
            id: 99, user: { id: 1, type: 'User' }, body: 'This is credential validation work.' });
        const fresh = await prepare(42, api);
        const outcome = await dispatch(fresh, result(), api, { enabled: true });
        assert.equal(outcome.status, 'blocked');
        assert.equal(state.calls.filter(call => call.endpoint.endsWith('/assignees')).length, 0);
    }
});

test('uncertain assignment response leaves a pending receipt and is never blindly retried', async () => {
    const { state, api, snapshot } = await ready();
    await publish(snapshot, result(), api);
    const failing = async (endpoint, options) => {
        if (endpoint.endsWith('/assignees')) {
            await api(endpoint, options);
            throw new Error('Connection lost after submit.');
        }
        return api(endpoint, options);
    };
    await assert.rejects(dispatch(snapshot, result(), failing, { enabled: true }), /Connection lost/);
    assert.equal(ownedState(state.comments[0], 42).handoff, 'pending');
    const fresh = await prepare(42, api);
    assert.equal((await dispatch(fresh, result(), api, { enabled: true })).status, 'already-started');
    assert.equal(state.issue.assignees.length, 1);
    assert.equal(state.calls.filter(call => call.endpoint.endsWith('/assignees')).length, 1);
});

test('wrong credential owner fails before assignment', async () => {
    const { state, api, snapshot } = await ready();
    await publish(snapshot, result(), api);
    await assert.rejects(dispatch(snapshot, result(), async (endpoint, options) =>
        endpoint === 'user' ? { login: 'another-owner' } : api(endpoint, options), { enabled: true }), /owner/i);
    assert.equal(state.issue.assignees.length, 0);
});

test('source validation uses the assessed commit and changed default branches cannot launch fixes', async () => {
    const { state, api, snapshot } = await ready();
    await publish(snapshot, result(), api);
    assert.ok(state.calls.filter(call => call.endpoint.includes('/contents/'))
        .every(call => call.endpoint.endsWith(`?ref=${snapshot.source_sha}`)));
    for (const change of ['branch', 'commit']) {
        const changed = async (endpoint, options) => {
            if (change === 'branch' && endpoint === `repos/${repository}`) {
                return { full_name: repository, default_branch: 'next' };
            }
            if (change === 'commit' && endpoint.endsWith('/commits/main')) return { sha: 'b'.repeat(40) };
            return api(endpoint, options);
        };
        await assert.rejects(dispatch(snapshot, result(), changed, { enabled: true }), /branch.*changed/i);
    }
    assert.equal(state.issue.assignees.length, 0);
});

test('failed comment writes never lead to assignment or a success receipt', async () => {
    const { state, api, snapshot } = await ready();
    const denied = async (endpoint, options) => {
        if (options?.method && options.method !== 'GET') throw new Error('Write denied.');
        return api(endpoint, options);
    };
    await assert.rejects(publish(snapshot, result(), denied), /Write denied/);
    assert.equal(state.comments.length, 0);
    await publish(snapshot, result(), api);
    await assert.rejects(dispatch(snapshot, result(), denied, { enabled: true }), /Write denied/);
    assert.equal(state.issue.assignees.length, 0);
    assert.equal(ownedState(state.comments[0], 42).handoff, 'none');
});

test('workflow preserves Squad and existing runner settings', () => {
    const workflow = fs.readFileSync(path.join(root, '.github', 'workflows', 'issue-triage.md'), 'utf8');
    assert.match(workflow, /types: \[opened, reopened\]/);
    assert.match(workflow, /roles: all/);
    assert.match(workflow, /min-integrity: none/);
    assert.match(workflow, /allowed-repos: \[sbroenne\/mcp-server-excel\]/);
    assert.match(workflow, /status-comment: false/);
    assert.match(workflow, /detection_success == 'true'/);
    assert.match(workflow, /ISSUE_AUTO_FIX_ENABLED/);
    assert.doesNotMatch(workflow, /add-labels:|close-issue:|create-pull-request:|issues: \[write\]/);
    for (const profile of ['squad-auto-fix', 'squad-fix-worker', 'squad-fix-reviewer']) {
        assert.equal(fs.existsSync(path.join(root, '.github', 'agents', `${profile}.agent.md`)), false);
    }
    assert.match(fs.readFileSync(path.join(root, '.github', 'workflows', 'copilot-setup-steps.yml'), 'utf8'),
        /vars.EXCEL_COPILOT_ENABLED == 'true'/);
});

test('generated workflow makes the validated target available to the agent and isolates assignment', () => {
    const source = fs.readFileSync(path.join(root, '.github', 'workflows', 'issue-triage.lock.yml'), 'utf8');
    const job = (lock, name) => {
        const match = lock.match(new RegExp(`^  ${name}:\\r?\\n([\\s\\S]*?)(?=^  [\\w-]+:|$(?![\\s\\S]))`, 'm'));
        assert.ok(match, `Missing generated job ${name}`);
        return match[1];
    };
    for (const newline of ['\n', '\r\n']) {
        const lock = source.replace(/\r?\n/g, newline);
        assert.match(job(lock, 'agent'), /needs:[\s\S]*? - pre_activation/);
        assert.match(job(lock, 'agent'), /timeout-minutes: 20/);
        assert.match(job(lock, 'pre_activation'), /issue_number: \$\{\{ steps.target.outputs.issue_number \}\}/);
        assert.match(job(lock, 'publish_triage'), /needs.detection.outputs.detection_success == 'true'/);
        assert.match(job(lock, 'publish_triage'), /GH_TOKEN: \$\{\{ secrets.COPILOT_ASSIGN_TOKEN \}\}/);
        assert.doesNotMatch(job(lock, 'agent'), /COPILOT_ASSIGN_TOKEN|issues: write|contents: write/);
        assert.match(lock, /group: issue-triage-\$\{\{ github.event.issue.number \|\| inputs.issue_number \}\}/);
        const manifest = JSON.parse(lock.split('\n').find(line => line.startsWith('# gh-aw-manifest: ')).slice(18));
        assert.deepEqual(manifest.mcp_servers.find(server => server.name === 'safeoutputs').tools,
            ['missing_data', 'missing_tool', 'noop', 'publish_triage']);
        assert.ok(manifest.actions.every(action => /^[a-f0-9]{40}$/.test(action.sha)));
    }
});

test('CLI prepare, publish and dispatch use the same guards, without persisting issue data', async () => {
    const { state, api } = fixture();
    const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-triage-'));
    try {
        const eventPath = path.join(directory, 'event.json');
        const snapshotPath = path.join(directory, 'target.json');
        const outputPath = path.join(directory, 'outputs.txt');
        const resultPath = path.join(directory, 'result.json');
        fs.writeFileSync(eventPath, JSON.stringify({ issue: { number: 42 } }));
        fs.writeFileSync(resultPath, JSON.stringify({ items: [{ type: 'publish_triage', analysis_json: JSON.stringify(result()) }] }));
        const env = { GITHUB_REPOSITORY: repository, GITHUB_EVENT_PATH: eventPath,
            GITHUB_EVENT_NAME: 'issues', GITHUB_OUTPUT: outputPath, GH_AW_AGENT_OUTPUT: resultPath };
        await main(['prepare', snapshotPath], env, api);
        assert.match(fs.readFileSync(outputPath, 'utf8'), /issue_number=42\neligible=true/);
        assert.doesNotMatch(fs.readFileSync(snapshotPath, 'utf8'), /reproducible defect|reporter|Range bug/);
        await main(['publish', snapshotPath], env, api);
        await main(['dispatch', snapshotPath], env, api);
        assert.equal(state.issue.assignees.length, 0);
        await main(['dispatch', snapshotPath], { ...env, ISSUE_AUTO_FIX_ENABLED: 'true', EXCEL_RUNNER_ENABLED: 'false' }, api);
        assert.match(state.comments[0].body, /runner control is disabled/);
        await main(['prepare', snapshotPath], { ...env, GITHUB_EVENT_NAME: 'workflow_dispatch', ISSUE_NUMBER: '42' }, api);
        fs.writeFileSync(resultPath, JSON.stringify({ items: [{ type: 'publish_triage', analysis_json: JSON.stringify(result()) }, { type: 'close_issue' }] }));
        await assert.rejects(main(['publish', snapshotPath], env, api), /exactly one/i);
        await assert.rejects(main(['prepare', snapshotPath], { ...env, GITHUB_REPOSITORY: 'another/repository' }, api), /repository/i);
        await assert.rejects(main(['prepare', snapshotPath], { ...env, GITHUB_EVENT_NAME: 'issue_comment' }, api), /Unsupported/);
        await assert.rejects(githubApi('repos/another/repository/issues/42'), /destination/i);
    } finally {
        fs.rmSync(directory, { recursive: true });
    }
});
