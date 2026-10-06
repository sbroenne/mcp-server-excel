const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const { test } = require('node:test');

const root = path.resolve(__dirname, '../..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');
const assign = require('../../.github/scripts/assign-squad-copilot.cjs');

function triageScript() {
  const workflow = read('.github/workflows/squad-triage.yml');
  const block = workflow.match(/          script: \|\r?\n([\s\S]*?)(?=\r?\n      - name:|$)/);
  assert.ok(block, 'triage script exists');
  return block[1].split(/\r?\n/).map(line => line.slice(12)).join('\n');
}

async function triage(title, team = read('.squad/team.md')) {
  const labels = [];
  const comments = [];
  const outputs = {};
  const fakeFs = {
    existsSync: () => true,
    readFileSync: file => file === '.squad/team.md' ? team : ''
  };
  const context = { repo: { owner: 'example', repo: 'example' }, payload: { issue: { number: 1, title } } };
  const core = { info() {}, warning() {}, setOutput: (key, value) => { outputs[key] = value; } };
  const github = { rest: { issues: {
    addLabels: async request => labels.push(...request.labels),
    createComment: async request => comments.push(request.body)
  } } };
  const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
  await new AsyncFunction('require', 'context', 'core', 'github', triageScript())(
    name => { assert.equal(name, 'fs'); return fakeFs; }, context, core, github);
  return { labels, comments, outputs };
}

test('copilot-only roster routes suitable issues without requiring a Lead', async () => {
  const result = await triage('bug fix: isolated regression');
  assert.ok(result.labels.includes('squad:copilot'));
  assert.equal(result.outputs['assign-copilot'], 'false');
  assert.match(result.comments[0], /Automatic assignment is disabled/);
});

test('copilot-only roster requests maintainer review for unsuitable work', async () => {
  const result = await triage('security architecture');
  assert.deepEqual(result.labels, []);
  assert.match(result.comments[0], /maintainer must review/);
});

test('configured Lead retains fallback routing', async () => {
  const team = read('.squad/team.md').replace('## Coding Agent', '| Alice | Lead | charter | Active |\n\n## Coding Agent');
  const result = await triage('security architecture', team);
  assert.ok(result.labels.includes('squad:alice'));
});

test('auto-assignment uses a separate PAT step in both workflows', async () => {
  const team = read('.squad/team.md').replace('copilot-auto-assign: false', 'copilot-auto-assign: true');
  const result = await triage('bug fix: isolated regression', team);
  assert.equal(result.outputs['assign-copilot'], 'true');
  for (const file of ['squad-triage.yml', 'squad-issue-assign.yml']) {
    const workflow = read(`.github/workflows/${file}`);
    assert.match(workflow, /github-token: \$\{\{ secrets\.COPILOT_ASSIGN_TOKEN \}\}/);
    assert.match(workflow, /require\('\.\/\.github\/scripts\/assign-squad-copilot\.cjs'\)/);
  }
});

test('shared assignment preserves the coding-agent request and surfaces failure', async () => {
  const requests = [];
  const context = { repo: { owner: 'example', repo: 'example' }, payload: { issue: { number: 7 } } };
  const core = { info() {}, setFailed: message => { requests.push(message); } };
  const github = {
    rest: { repos: { get: async () => ({ data: { default_branch: 'main' } }) } },
    request: async (route, request) => requests.push({ route, request })
  };
  await assign({ github, context, core, token: '' });
  assert.match(requests.pop(), /COPILOT_ASSIGN_TOKEN is required/);
  await assign({ github, context, core, token: 'test-placeholder' });
  assert.deepEqual(requests[0].request.assignees, ['copilot-swe-agent[bot]']);
  assert.equal(requests[0].request.agent_assignment.base_branch, 'main');
  github.request = async () => { throw new Error('assignment rejected'); };
  await assert.rejects(assign({ github, context, core, token: 'test-placeholder' }), /assignment rejected/);
});

test('mesh sync keeps URLs, output paths, and bearer values as literal curl arguments', () => {
  const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'squad-mesh-test-'));
  const bashPath = value => process.platform === 'win32'
    ? value.replace(/\\/g, '/').replace(/^([a-z]):/i, (_, drive) => `/${drive.toLowerCase()}`)
    : value;
  try {
    const source = 'https://example.invalid/?a=1&b=2; touch injected';
    const target = 'output folder; touch target-injected';
    const bearer = 'fixture; touch token-injected';
    const config = path.join(temp, 'mesh.json');
    const output = path.join(temp, 'curl-args.txt');
    fs.writeFileSync(config, JSON.stringify({ squads: { fixture: {
      zone: 'remote-opaque', source, sync_to: target, auth: 'bearer'
    } } }));
    const script = path.join(root, '.squad/templates/skills/distributed-mesh/sync-mesh.sh');
    const result = spawnSync('bash', ['-c',
      'curl() { printf "%s\\n" "$@" > "$CURL_ARGUMENTS"; }; export -f curl; bash "$1" "$2"',
      'test', bashPath(script), bashPath(config)], {
      cwd: temp, encoding: 'utf8',
      env: { ...process.env, FIXTURE_TOKEN: bearer, CURL_ARGUMENTS: bashPath(output) }
    });
    assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
    const args = fs.readFileSync(output, 'utf8').trim().split('\n');
    assert.ok(args.includes(source));
    assert.ok(args.includes(`${target}/SUMMARY.md`));
    assert.ok(args.includes(`Authorization: Bearer ${bearer}`));
    for (const marker of ['injected', 'target-injected', 'token-injected']) {
      assert.equal(fs.existsSync(path.join(temp, marker)), false);
    }
  } finally {
    fs.rmSync(temp, { recursive: true, force: true });
  }
});

function run(command, args, cwd) {
  const result = spawnSync(command, args, { cwd, encoding: 'utf8' });
  assert.equal(result.status, 0, `${command} ${args.join(' ')}\n${result.stdout}\n${result.stderr}`);
  return result.stdout.trim();
}

test('screenshot template publishes from an isolated worktree without overwriting edits', () => {
  const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'squad-screenshots-test-'));
  try {
    const remote = path.join(temp, 'remote.git');
    const repo = path.join(temp, 'repo');
    run('git', ['init', '--bare', remote], temp);
    run('git', ['clone', remote, repo], temp);
    run('git', ['config', 'user.name', 'Test'], repo);
    run('git', ['config', 'user.email', 'test@example.invalid'], repo);
    fs.writeFileSync(path.join(repo, 'work.txt'), 'committed');
    run('git', ['add', 'work.txt'], repo);
    run('git', ['commit', '-m', 'fixture'], repo);
    fs.writeFileSync(path.join(repo, 'work.txt'), 'uncommitted edit');
    fs.mkdirSync(path.join(repo, 'screenshots'));
    fs.writeFileSync(path.join(repo, 'screenshots', 'fixture.png'), 'synthetic image fixture');
    const before = run('git', ['status', '--porcelain'], repo);
    const originalBranch = run('git', ['branch', '--show-current'], repo);
    const skill = read('.squad/templates/skills/pr-screenshots/SKILL.md');
    const code = skill.match(/```powershell\r?\n([\s\S]*?)```/)[1].replaceAll('{PR_NUMBER}', '123');
    run('pwsh', ['-NoProfile', '-Command', code], repo);
    assert.equal(run('git', ['branch', '--show-current'], repo), originalBranch);
    assert.equal(run('git', ['status', '--porcelain'], repo), before);
    assert.equal(fs.readFileSync(path.join(repo, 'work.txt'), 'utf8'), 'uncommitted edit');
    const ref = run('git', ['ls-remote', '--heads', 'origin', 'screenshots-pr-123-*'], repo).split(/\s+/)[1];
    assert.ok(ref);
    run('git', ['fetch', 'origin', ref], repo);
    assert.equal(run('git', ['ls-tree', '-r', '--name-only', 'FETCH_HEAD'], repo), 'screenshots/fixture.png');
    run('git', ['branch', '-D', ref.replace('refs/heads/', '')], repo);
  } finally {
    fs.rmSync(temp, { recursive: true, force: true });
  }
});

test('notes fetch bootstraps and merges divergent notes without losing local entries', () => {
  const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'squad-notes-test-'));
  try {
    const remote = path.join(temp, 'remote.git');
    const left = path.join(temp, 'left');
    const right = path.join(temp, 'right');
    run('git', ['init', '--bare', remote], temp);
    run('git', ['clone', remote, left], temp);
    const config = cwd => {
      run('git', ['config', 'user.name', 'Test'], cwd);
      run('git', ['config', 'user.email', 'test@example.invalid'], cwd);
    };
    config(left);
    run('git', ['commit', '--allow-empty', '-m', 'fixture'], left);
    run('git', ['push', 'origin', 'HEAD'], left);
    run('git', ['notes', '--ref=squad/test', 'add', '-m', 'base', 'HEAD'], left);
    run('git', ['push', 'origin', 'refs/notes/squad/test'], left);
    run('git', ['clone', left, right], temp);
    config(right);
    run('git', ['remote', 'set-url', 'origin', remote], right);
    const fetch = flags => run('pwsh', ['-NoProfile', '-File',
      path.join(root, '.squad/templates/scripts/notes/fetch.ps1'), '-RepoPath', right, ...flags], root);
    // Migrate the legacy refspec as well as initialize local notes.
    run('git', ['config', '--add', 'remote.origin.fetch', 'refs/notes/*:refs/notes/*'], right);
    fetch(['-Setup']);
    assert.equal(run('git', ['notes', '--ref=squad/test', 'show', 'HEAD'], right), 'base');
    assert.match(run('git', ['config', '--get-all', 'remote.origin.fetch'], right), /refs\/notes\/remotes\/origin\/squad\/\*/);
    run('git', ['notes', '--ref=squad/test', 'append', '-m', 'local', 'HEAD'], right);
    run('git', ['notes', '--ref=squad/test', 'append', '-m', 'remote', 'HEAD'], left);
    run('git', ['push', 'origin', 'refs/notes/squad/test'], left);
    fetch([]);
    assert.match(run('git', ['notes', '--ref=squad/test', 'show', 'HEAD'], right), /local/);
    fetch(['-Merge']);
    const note = run('git', ['notes', '--ref=squad/test', 'show', 'HEAD'], right);
    for (const text of ['base', 'local', 'remote']) assert.ok(note.includes(text));
    run('git', ['push', 'origin', 'refs/notes/squad/test'], right);
  } finally {
    fs.rmSync(temp, { recursive: true, force: true });
  }
});
