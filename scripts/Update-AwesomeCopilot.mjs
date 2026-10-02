import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import {
    canonical, hash, git, readGitTree, resolveCommit, assertTag, parseJson, externalCommandEnvironment,
} from './PluginContent.mjs';
import {
    upstreamRepo, forkRepo, prAuthor, allowedPaths, planListings, parseState, stateMarker,
    assertListingPatch, assertAllowedPaths, ownedPulls,
} from './AwesomeCopilotPolicy.mjs';

function run(command, args, cwd, env = externalCommandEnvironment()) {
    return execFileSync(command, args, {
        cwd, env, windowsHide: true, maxBuffer: 64 * 1024 * 1024,
        encoding: 'utf8', shell: process.platform === 'win32' && command === 'npm',
    });
}

const applicationCredential = /^(AWESOME_COPILOT_PR_TOKEN|PLUGINS_REPO_TOKEN|RELEASE_PAT|GH_TOKEN|GITHUB_TOKEN|COPILOT_GITHUB_TOKEN|GH_AW_GITHUB_TOKEN|GH_AW_GITHUB_MCP_SERVER_TOKEN)$/i;
const initialApplicationCredential = Object.entries(process.env).some(([name, value]) => applicationCredential.test(name) && value) ||
    (process.platform === 'linux' && fs.readFileSync('/proc/self/environ', 'utf8').split('\0').some(entry => {
        const equals = entry.indexOf('=');
        return equals > 0 && applicationCredential.test(entry.slice(0, equals)) && entry.length > equals + 1;
    }));

function assertBuildEnvironment(env = process.env) {
    if (initialApplicationCredential || [process.env, env].some(environment =>
        Object.entries(environment).some(([name, value]) => applicationCredential.test(name) && value))) {
        throw new Error('Upstream build blocked: GitHub/write credential is present in the parent process. Use separate discovery and build jobs.');
    }
}

export function validateUpstreamBuild(directory, { env = process.env, execute = run } = {}) {
    assertBuildEnvironment(env);
    const childEnv = externalCommandEnvironment(env);
    execute('npm', ['ci', '--ignore-scripts', '--no-audit', '--no-fund'], directory, childEnv);
    execute('npm', ['run', 'plugin:validate'], directory, childEnv);
    execute('npm', ['run', 'build'], directory, childEnv);
}

export function api(endpoint, { method = 'GET', body, token } = {}) {
    const permittedSearch = method === 'GET' && endpoint.startsWith('search/issues?') &&
        new URLSearchParams(endpoint.split('?')[1]).get('q') ===
            'repo:github/awesome-copilot is:pr is:closed author:sbroenne "excel-plugin-update-state" in:body';
    if (!/^repos\/(github\/awesome-copilot|sbroenne\/awesome-copilot)(\/|$)/.test(endpoint) && endpoint !== 'user' && !permittedSearch) {
        throw new Error('GitHub API destination is not allowlisted.');
    }
    const credential = token ?? process.env.GH_TOKEN ?? process.env.GITHUB_TOKEN;
    if (!credential) {
        if (method !== 'GET' || body) throw new Error('Authenticated API writes require an explicit credential.');
        const text = run('curl', ['--disable', '--fail-with-body', '--silent', '--show-error',
            '--proto', '=https', '-H', 'Accept: application/vnd.github+json', `https://api.github.com/${endpoint}`]);
        return parseJson(Buffer.from(text));
    }
    const args = ['api', endpoint, '--method', method];
    if (body) args.push('--input', '-');
    const text = execFileSync('gh', args, {
        input: body ? JSON.stringify(body) : undefined, encoding: 'utf8', windowsHide: true,
        env: { ...externalCommandEnvironment(),
            GH_TOKEN: credential },
        maxBuffer: 16 * 1024 * 1024,
    });
    return parseJson(Buffer.from(text));
}

export function allPulls(read = api) {
    const result = [];
    for (let page = 1; ; page++) {
        const batch = read(`repos/${upstreamRepo}/pulls?state=open&per_page=100&page=${page}`);
        if (!Array.isArray(batch)) throw new Error('Invalid paginated PR API data.');
        result.push(...batch);
        if (batch.length < 100) break;
    }
    // REST covers recently closed PRs before the search index catches up.
    for (let page = 1; ; page++) {
        const batch = read(`repos/${upstreamRepo}/pulls?state=closed&sort=updated&direction=desc&per_page=100&page=${page}`);
        if (!Array.isArray(batch)) throw new Error('Invalid recently closed PR API data.');
        result.push(...batch);
        if (batch.some(pr => !Number.isFinite(Date.parse(pr.updated_at)))) throw new Error('Missing PR update timestamp.');
        if (batch.length < 100 || batch.some(pr => Date.parse(pr.updated_at) < Date.now() - 24 * 60 * 60 * 1000)) break;
    }
    const query = encodeURIComponent('repo:github/awesome-copilot is:pr is:closed author:sbroenne "excel-plugin-update-state" in:body');
    for (let page = 1; ; page++) {
        const search = read(`search/issues?q=${query}&per_page=100&page=${page}`);
        if (search.incomplete_results !== false || !Number.isSafeInteger(search.total_count) ||
            search.total_count < 0 || search.total_count > 1000 || !Array.isArray(search.items)) {
            throw new Error('Incomplete declined-PR search API data.');
        }
        const remaining = search.total_count - (page - 1) * 100;
        if (remaining < 0 || search.items.length !== Math.min(100, remaining)) throw new Error('Truncated declined-PR search.');
        for (const item of search.items) {
            if (!Number.isSafeInteger(item.number)) throw new Error('Invalid closed PR search result.');
            result.push(read(`repos/${upstreamRepo}/pulls/${item.number}`));
        }
        if (page * 100 >= search.total_count) break;
    }
    return [...new Map(result.map(pr => [pr.number, pr])).values()];
}

function physicalPath(value) {
    if (process.platform !== 'win32') return value;
    if (value.startsWith('\\\\?\\UNC\\')) return `\\\\${value.slice(8)}`;
    return value.startsWith('\\\\?\\') ? value.slice(4) : value;
}

export function safeUpdaterPath(target, { directory = false, allowMissing = false } = {}) {
    const absolute = path.resolve(target), volume = path.parse(absolute).root;
    const parts = absolute.slice(volume.length).split(path.sep).filter(Boolean);
    let current = volume;
    const existing = [];
    for (let i = -1; i < parts.length; i++) {
        if (i >= 0) current = path.join(current, parts[i]);
        const stat = fs.lstatSync(current, { throwIfNoEntry: false });
        if (!stat) {
            if (allowMissing && existing.length > 0) break;
            throw new Error('Unsafe updater path: required path is missing.');
        }
        if (stat.isSymbolicLink()) throw new Error('Unsafe updater path: links are forbidden.');
        const isDirectory = i < parts.length - 1 || directory;
        if (isDirectory ? !stat.isDirectory() : !stat.isFile()) throw new Error('Unsafe updater path type.');
        const resolved = fs.realpathSync.native(current);
        if (process.platform !== 'win32' && resolved !== current) throw new Error('Unsafe updater path: realpath escapes its lexical boundary.');
        existing.push({ path: current, real: resolved });
    }
    if (process.platform === 'win32') {
        const expanded = parseJson(Buffer.from(run('pwsh', ['-NoProfile', '-NonInteractive', '-Command', `
            $ErrorActionPreference="Stop"
            Add-Type 'using System.Text; using System.Runtime.InteropServices; public static class UpdaterLongPaths {
                [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
                public static extern uint GetLongPathName(string p, StringBuilder b, uint n);
            }'
            $names=@(foreach($item in (ConvertFrom-Json $env:EXCEL_UPDATER_PATHS)) {
                if (([IO.File]::GetAttributes($item.path) -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                    throw "Unsafe updater path: reparse points are forbidden."
                }
                $name=[Text.StringBuilder]::new(32768)
                $length=[UpdaterLongPaths]::GetLongPathName($item.path,$name,32768)
                if ($length -eq 0) { throw [ComponentModel.Win32Exception]::new([Runtime.InteropServices.Marshal]::GetLastWin32Error()) }
                if ($length -ge 32768) { throw "Unsafe updater path: expanded path exceeds buffer." }
                $name.ToString()
            })
            ConvertTo-Json -InputObject $names -Compress
        `], undefined, { ...externalCommandEnvironment(), EXCEL_UPDATER_PATHS: JSON.stringify(existing) })));
        if (!Array.isArray(expanded) || expanded.length !== existing.length ||
            expanded.some((value, i) => typeof value !== 'string' ||
                physicalPath(value).toLowerCase() !== physicalPath(existing[i].real).toLowerCase())) {
            throw new Error('Unsafe updater path: realpath escapes its expanded lexical boundary.');
        }
    }
    const last = existing.at(-1);
    return path.join(physicalPath(last.real), path.relative(last.path, absolute));
}

function workspace(directory, allowMissing = false) {
    const absolute = safeUpdaterPath(directory, { directory: true, allowMissing });
    const source = physicalPath(fs.realpathSync.native(path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')));
    const contained = (parent, child) => {
        const relative = path.relative(parent, child);
        return relative === '' || (!relative.startsWith(`..${path.sep}`) && relative !== '..' && !path.isAbsolute(relative));
    };
    if (contained(source, absolute) || contained(absolute, source)) throw new Error('Unsafe updater workspace: must not overlap trusted source.');
    return absolute;
}

function allowedInputs(directory, commit = 'HEAD') {
    workspace(directory);
    for (const name of allowedPaths) safeUpdaterPath(path.join(directory, name), { allowMissing: true });
    const entries = new Map(git(directory, ['ls-tree', '-z', commit, '--', ...allowedPaths]).toString().split('\0')
        .filter(Boolean).map(entry => {
            const match = /^(\d+) (\w+) [a-f0-9]{40}\t(.+)$/.exec(entry);
            if (!match) throw new Error('Invalid allowed-input Git mode data.');
            return [match[3], { mode: match[1], type: match[2] }];
        }));
    for (const name of allowedPaths) {
        const entry = entries.get(name);
        if (!entry || entry.type !== 'blob' || !['100644', '100755'].includes(entry.mode)) {
            throw new Error('Unsafe allowed-input Git mode: tracked links or missing files are forbidden.');
        }
    }
}

function clone(repository, directory) {
    workspace(directory, true);
    if (fs.existsSync(directory)) throw new Error('Use a new disposable updater workspace.');
    run('git', ['clone', '--quiet', '--no-checkout', `https://github.com/${repository}.git`, directory]);
}

function listedFile(directory, commit) {
    allowedInputs(directory, commit);
    return parseJson(git(directory, ['show', `${commit}:plugins/external.json`]));
}

function checkout(directory, commit) {
    allowedInputs(directory, commit);
    git(directory, ['-c', 'core.autocrlf=false', 'checkout', '--quiet', '--detach', commit]);
    git(directory, ['config', 'core.autocrlf', 'false']);
    allowedInputs(directory, commit);
}

function forkHeads(text) {
    if (typeof text !== 'string') throw new Error('Invalid fork head response.');
    return text.trim().split(/\r?\n/).filter(Boolean).map(line => {
        const match = /^([a-f0-9]{40})\trefs\/heads\/(excel-plugin-updates-[a-f0-9]{12})$/.exec(line);
        if (!match) throw new Error('Invalid fork head response.');
        return { sha: match[1], branch: match[2] };
    });
}

function readForkHeads() {
    return run('git', ['ls-remote', '--heads', `https://github.com/${forkRepo}.git`, 'refs/heads/excel-plugin-updates-*']);
}

export function assertForkHeads(plan, text, pulls) {
    const heads = forkHeads(text);
    if (plan.action === 'create') {
        const associated = new Set(ownedPulls(pulls).map(pr => pr.head.ref));
        if (heads.some(head => !associated.has(head.branch))) {
            throw new Error('An owned-prefix orphan branch exists; reconcile it before creating any new proposal.');
        }
        if (heads.some(head => head.branch === plan.branch)) {
            throw new Error('Proposed branch already exists; no replacement or overwrite is allowed.');
        }
    } else if (heads.find(head => head.branch === plan.branch)?.sha !== plan.expectedHead) {
        throw new Error('Fork branch head changed.');
    }
}

export function assertPendingTree(directory, plan) {
    if (plan.action !== 'update') return;
    const state = plan.previousState;
    if (!/^[a-f0-9]{40}$/.test(state.base ?? '') || state.head !== plan.expectedHead) throw new Error('Missing expected PR base/head.');
    const before = listedFile(directory, state.base), after = listedFile(directory, state.head);
    const affected = Object.keys(state.entries);
    assertListingPatch(before, after, affected);
    const paths = git(directory, ['diff', '--name-only', state.base, state.head]).toString().trim().split(/\r?\n/);
    assertAllowedPaths(paths);
    for (const name of affected) {
        if (canonical(after.find(entry => entry.name === name)) !== canonical(state.entries[name].entry)) {
            throw new Error('Pending branch contains human listing changes.');
        }
    }
}

export function patch(directory, plan, validatedFiles) {
    allowedInputs(directory);
    const file = path.join(directory, 'plugins', 'external.json');
    const before = parseJson(fs.readFileSync(safeUpdaterPath(file)));
    const after = before.map(entry => plan.state.entries[entry.name]?.entry ?? entry);
    assertListingPatch(before, after, plan.changedPlugins);
    if (validatedFiles !== undefined) {
        if (!validatedFiles || typeof validatedFiles !== 'object' || Array.isArray(validatedFiles)) {
            throw new Error('Missing validated marketplace files.');
        }
        assertAllowedPaths(Object.keys(validatedFiles));
        if (Object.values(validatedFiles).some(text => typeof text !== 'string') ||
            Buffer.byteLength(canonical(validatedFiles)) > 2 * 1024 * 1024 ||
            !Object.hasOwn(validatedFiles, 'plugins/external.json') ||
            canonical(parseJson(Buffer.from(validatedFiles['plugins/external.json']))) !== canonical(after)) {
            throw new Error('Validated files do not match the freshly verified listing proposal.');
        }
        for (const [name, text] of Object.entries(validatedFiles)) fs.writeFileSync(safeUpdaterPath(path.join(directory, name), { allowMissing: true }), text);
    } else {
        fs.writeFileSync(safeUpdaterPath(file), `${JSON.stringify(after, null, 2)}\n`);
        validateUpstreamBuild(directory);
    }
    allowedInputs(directory);
    const paths = git(directory, ['diff', '--name-only']).toString().trim().split(/\r?\n/);
    assertAllowedPaths(paths);
    const external = parseJson(fs.readFileSync(safeUpdaterPath(file)));
    assertListingPatch(before, external, plan.changedPlugins);
    const oldMarket = parseJson(git(directory, ['show', 'HEAD:.github/plugin/marketplace.json']));
    const market = parseJson(fs.readFileSync(safeUpdaterPath(path.join(directory, '.github', 'plugin', 'marketplace.json'))));
    assertListingPatch(oldMarket.plugins, market.plugins, plan.changedPlugins);
    const oldRoot = { ...oldMarket }, newRoot = { ...market };
    delete oldRoot.plugins;
    delete newRoot.plugins;
    if (canonical(oldRoot) !== canonical(newRoot)) throw new Error('Generated marketplace root metadata changed.');
    for (const name of plan.changedPlugins) {
        if (canonical(market.plugins.find(entry => entry.name === name)) !== canonical(plan.state.entries[name].entry)) {
            throw new Error('Generated marketplace does not match the verified proposed entry.');
        }
    }
    const files = Object.fromEntries(paths.map(name => [name, fs.readFileSync(safeUpdaterPath(path.join(directory, name)), 'utf8')]));
    if (Buffer.byteLength(canonical(files)) > 2 * 1024 * 1024) throw new Error('Marketplace patch exceeds two MiB.');
    if (validatedFiles !== undefined && canonical(files) !== canonical(validatedFiles)) {
        throw new Error('Validated patch files differ from the exact destination diff.');
    }
    return files;
}

export function prepare({ tag, workDirectory }) {
    assertBuildEnvironment();
    return preparePlan({ tag, workDirectory });
}

export function discover({ tag, workDirectory }) {
    return preparePlan({ tag, workDirectory, discoveryOnly: true });
}

export function buildDiscoveredPlan({ discovery, workDirectory }) {
    if (!discovery || !Array.isArray(discovery.pulls) ||
        discovery.discoveryFingerprint !== discoveryFingerprint(discovery)) {
        throw new Error('Invalid public discovery artifact/digest.');
    }
    assertBuildEnvironment();
    return preparePlan({ tag: discovery.tag, workDirectory, discovery });
}

function discoveryFingerprint(plan) {
    const inputs = { ...plan };
    delete inputs.discoveryFingerprint;
    return hash(canonical(inputs));
}

function preparePlan({ tag, workDirectory, validatedFiles, discoveryOnly = false, discovery }) {
    assertTag(tag);
    workspace(workDirectory, true);
    fs.mkdirSync(workDirectory, { recursive: true });
    const published = path.join(workDirectory, 'published');
    const upstream = path.join(workDirectory, 'upstream');
    clone('sbroenne/mcp-server-excel-plugins', published);
    clone(upstreamRepo, upstream);
    const commit = resolveCommit(published, `refs/tags/${tag}`);
    const upstreamCommit = resolveCommit(upstream, 'refs/remotes/origin/main');
    const cache = new Map();
    const getTree = sha => {
        if (!cache.has(sha)) cache.set(sha, readGitTree(published, sha));
        return cache.get(sha);
    };
    const resolveTag = value => resolveCommit(published, `refs/tags/${assertTag(value)}`);
    const pulls = discovery?.pulls ?? allPulls();
    const plan = planListings({
        listings: listedFile(upstream, upstreamCommit), tag, commit, getTree, resolveTag, pulls,
    });
    plan.tag = tag;
    plan.commit = commit;
    plan.upstreamCommit = upstreamCommit;
    if (plan.action !== 'noop') plan.state.base = plan.previousState?.base ?? upstreamCommit;
    if (discoveryOnly || discovery) {
        const snapshot = { ...plan, pulls };
        snapshot.discoveryFingerprint = discoveryFingerprint(snapshot);
        if (discovery && snapshot.discoveryFingerprint !== discovery.discoveryFingerprint) {
            throw new Error('Public discovery baseline changed; retry discovery before building.');
        }
        if (discoveryOnly) return snapshot;
    }
    if (plan.action === 'noop') return plan;
    if (plan.action === 'update') {
        git(upstream, ['fetch', '--quiet', `https://github.com/${forkRepo}.git`, plan.expectedHead]);
        assertPendingTree(upstream, plan);
    }
    assertForkHeads(plan, readForkHeads(), pulls);
    checkout(upstream, plan.expectedHead ?? upstreamCommit);
    plan.files = patch(upstream, plan, validatedFiles);
    plan.patchFingerprint = hash(canonical(plan.files));
    plan.guardFingerprint = guardFingerprint(plan);
    return plan;
}

function guardFingerprint(plan) {
    return hash(canonical({
        action: plan.action, tag: plan.tag, commit: plan.commit, upstreamCommit: plan.upstreamCommit,
        state: plan.state, branch: plan.branch, expectedHead: plan.expectedHead, pullNumber: plan.pullNumber,
        expectedBody: plan.expectedBody,
        patchFingerprint: plan.patchFingerprint,
    }));
}

export function recheckPreparedPlan({ trustedPlan, workDirectory }) {
    if (!trustedPlan.files || trustedPlan.patchFingerprint !== hash(canonical(trustedPlan.files)) ||
        trustedPlan.guardFingerprint !== guardFingerprint(trustedPlan)) {
        throw new Error('Trusted build plan fingerprint mismatch; no writes performed.');
    }
    // Recheck current public inputs and exact built bytes without executing upstream code in the writer.
    const fresh = preparePlan({ tag: trustedPlan.tag, workDirectory, validatedFiles: trustedPlan.files });
    if (fresh.action !== 'noop' && fresh.guardFingerprint !== trustedPlan.guardFingerprint) {
        throw new Error('Precheck changed; retry preview. No writes performed.');
    }
    return fresh;
}

export function validateRequest(output, trustedPlan) {
    if (!output || !Array.isArray(output.items) || output.items.length !== 1) throw new Error('Exactly one safe-output proposal is required.');
    const item = output.items[0];
    if (item.type !== 'submit_marketplace_update' ||
        Object.keys(item).some(key => !['type', 'proposal_fingerprint', 'body'].includes(key)) ||
        item.proposal_fingerprint !== trustedPlan.guardFingerprint ||
        typeof item.body !== 'string' || item.body.length < 40 || item.body.length > 50000 ||
        item.body.includes('excel-plugin-update-state:')) throw new Error('Invalid or conflicting safe-output request.');
    return item;
}

export function writesEnabled(env) {
    return env.AWESOME_COPILOT_UPDATES_ENABLED === 'true' &&
        env.PREVIEW === 'false' && env.GH_AW_SAFE_OUTPUTS_STAGED !== 'true';
}

export function assertTemplate(template, body) {
    template = template.replace(/\r\n/g, '\n');
    body = body.replace(/\r\n/g, '\n').replace(/^(- \[)[xX](\] .*)$/gm, '$1 $2');
    const elements = template.match(/^#{1,6} .+$|^- \[[ xX]\] .+$|<!--[\s\S]*?-->/gm) ?? [];
    let position = 0;
    for (const element of elements) {
        const found = body.indexOf(element.replace(/^(- \[)[xX](\] .*)$/gm, '$1 $2'), position);
        if (found < 0) throw new Error('PR body must preserve upstream template headings, comments, checklist items and order.');
        position = found + element.length;
    }
}

function verifyLive(plan) {
    const pulls = allPulls();
    const owned = ownedPulls(pulls).filter(pr => pr.state === 'open');
    if (plan.action === 'create') {
        if (owned.length) throw new Error('Another automated PR appeared; no duplicate will be created.');
    } else {
        if (owned.length !== 1 || owned[0].number !== plan.pullNumber ||
            owned[0].head.sha !== plan.expectedHead || hash(owned[0].body) !== plan.expectedBody ||
            parseState(owned[0].body).head !== plan.expectedHead) {
            throw new Error('PR state/head changed; refusing to overwrite or replace.');
        }
    }
    assertForkHeads(plan, readForkHeads(), pulls);
}

export function refreshPullAfterPush(plan, head, body, { readPull, readHead, writeBody,
    wait = () => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 2000) }) {
    const assertHead = () => {
        if (readHead() !== head) throw new Error('Fork head changed after push; manual reconciliation required.');
    };
    let lastError;
    for (let attempt = 0; attempt < 8; attempt++) {
        assertHead();
        let live;
        try { live = readPull(); }
        catch (error) { lastError = error; }
        if (live) {
            if (live.state !== 'open' || typeof live.body !== 'string' || ![plan.expectedHead, head].includes(live.head?.sha) ||
                (hash(live.body) !== plan.expectedBody && live.body !== body)) {
                throw new Error('PR changed after push; manual reconciliation required, no replacement PR.');
            }
            if (live.head.sha === head && live.body === body) {
                assertHead();
                return live;
            }
            if (live.head.sha === head) {
                let response;
                try {
                    response = writeBody();
                }
                catch (error) { lastError = error; }
                if (response?.head?.sha === head && response.body === body && response.state === 'open') {
                    assertHead();
                    return response;
                }
                if (response) lastError = new Error('PR update response not yet consistent; rechecking the exact transition.');
            }
        }
        if (attempt < 7) {
            console.error(`PR propagation/update retry ${attempt + 1}/8; expected fork head and body remain protected.`);
            wait();
        }
    }
    throw new Error('PR propagation/body update did not complete; use the submission receipt for manual reconciliation. No replacement PR.', {
        cause: lastError,
    });
}

export function createSubmissionReceipt(plan, head, body) {
    return {
        schema: 1, status: 'prepared', action: plan.action, upstream: upstreamRepo, fork: forkRepo, base: plan.state.base,
        branch: plan.branch, pullNumber: plan.pullNumber, previousHead: plan.expectedHead,
        expectedBody: plan.expectedBody, head, body, files: plan.files,
    };
}

export function submit({ trustedPlan, output, workDirectory, env = process.env }) {
    const request = validateRequest(output, trustedPlan);
    if (!writesEnabled(env)) return { status: 'preview', reason: 'Write gate disabled or staged; no GitHub mutations.' };
    const token = env.AWESOME_COPILOT_PR_TOKEN;
    if (!token) throw new Error('AWESOME_COPILOT_PR_TOKEN is required in the safe-output job.');
    if (api('user', { token }).login !== prAuthor) throw new Error('PR token owner must match the configured author.');
    const fork = api(`repos/${forkRepo}`, { token });
    if (!fork.fork || fork.parent?.full_name !== upstreamRepo || fork.permissions?.push !== true) {
        throw new Error('Destination must be the writable allowlisted upstream fork.');
    }
    const fresh = recheckPreparedPlan({ trustedPlan, workDirectory });
    if (fresh.action === 'noop') return { status: 'skipped', reason: fresh.reason };
    const upstream = path.join(workDirectory, 'upstream');
    const templatePath = path.join(upstream, '.github', 'pull_request_template.md');
    safeUpdaterPath(templatePath);
    const templateMode = git(upstream, ['ls-tree', 'HEAD', '--', '.github/pull_request_template.md']).toString();
    if (!/^100(644|755) blob /.test(templateMode)) throw new Error('Unsafe upstream PR template Git mode.');
    assertTemplate(fs.readFileSync(safeUpdaterPath(templatePath), 'utf8'), request.body);
    git(upstream, ['config', 'user.name', 'Excel Plugin Updates']);
    git(upstream, ['config', 'user.email', '3026464+sbroenne@users.noreply.github.com']);
    git(upstream, ['add', '--', ...allowedPaths]);
    git(upstream, ['commit', '--quiet', '-m', `Update Excel plugin listings (${fresh.state.proposalFingerprint})`]);
    const head = resolveCommit(upstream, 'HEAD');
    const state = { ...fresh.state, head, bodyFingerprint: hash(request.body.trimEnd()) };
    const body = `${request.body.trimEnd()}\n\n${stateMarker(state)}`;
    verifyLive(fresh);
    const receipt = createSubmissionReceipt(fresh, head, body);
    const receiptFile = path.join(workDirectory, 'submission-receipt.json');
    const saveReceipt = () => fs.writeFileSync(safeUpdaterPath(receiptFile, { allowMissing: true }), `${JSON.stringify(receipt, null, 2)}\n`);
    saveReceipt();
    const auth = Buffer.from(`x-access-token:${token}`).toString('base64');
    run('git', ['push', `https://github.com/${forkRepo}.git`, `HEAD:refs/heads/${fresh.branch}`], upstream, {
        ...externalCommandEnvironment(), GIT_CONFIG_COUNT: '1', GIT_CONFIG_KEY_0: 'http.https://github.com/.extraheader',
        GIT_CONFIG_VALUE_0: `AUTHORIZATION: basic ${auth}`, GIT_TERMINAL_PROMPT: '0',
    });
    receipt.status = 'pushed';
    saveReceipt();
    let response;
    if (fresh.action === 'create') {
        // No fallback issue, alternate branch, replacement PR or upstream labels.
        if (ownedPulls(allPulls()).some(pr => pr.state === 'open')) throw new Error('Automated PR appeared after push; do not create a duplicate.');
        response = api(`repos/${upstreamRepo}/pulls`, {
            method: 'POST', token,
            body: { title: 'Update Excel plugin listings', body, head: `sbroenne:${fresh.branch}`, base: 'main', maintainer_can_modify: true },
        });
    } else {
        response = refreshPullAfterPush(fresh, head, body, {
            readPull: () => api(`repos/${upstreamRepo}/pulls/${fresh.pullNumber}`),
            readHead: () => forkHeads(readForkHeads()).find(item => item.branch === fresh.branch)?.sha,
            writeBody: () => api(`repos/${upstreamRepo}/pulls/${fresh.pullNumber}`, { method: 'PATCH', token, body: { body } }),
        });
    }
    if (!response?.html_url || response.head?.sha !== head) throw new Error('Invalid PR write response; inspect the existing branch/PR before retrying.');
    receipt.status = 'completed';
    receipt.pullNumber = response.number;
    saveReceipt();
    return { status: fresh.action === 'create' ? 'created' : 'updated', url: response.html_url, number: response.number, head };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
    const [mode, tag, workDirectory, planFile] = process.argv.slice(2);
    let result;
    if (['prepare', 'discover', 'build'].includes(mode)) {
        result = mode === 'discover' ? discover({ tag, workDirectory }) :
            mode === 'build' ? buildDiscoveredPlan({ discovery: parseJson(fs.readFileSync(safeUpdaterPath(tag))), workDirectory }) :
                prepare({ tag, workDirectory });
        fs.writeFileSync(safeUpdaterPath(planFile, { allowMissing: true }), `${JSON.stringify(result, null, 2)}\n`);
        if (process.env.GITHUB_OUTPUT) {
            fs.appendFileSync(safeUpdaterPath(process.env.GITHUB_OUTPUT, { allowMissing: true }),
                `actionable=${result.action !== 'noop' && process.env.PREVIEW === 'false' && process.env.AWESOME_COPILOT_UPDATES_ENABLED === 'true'}\n` +
                `upstream_commit=${result.upstreamCommit}\n`);
        }
    } else if (mode === 'submit') {
        const trustedPlan = parseJson(fs.readFileSync(safeUpdaterPath(planFile)));
        const output = parseJson(fs.readFileSync(safeUpdaterPath(process.env.GH_AW_AGENT_OUTPUT)));
        result = submit({ trustedPlan, output, workDirectory });
    } else throw new Error('Use discover, build, prepare or submit.');
    const summary = { action: result.action, status: result.status, reason: result.reason,
        changedPlugins: result.changedPlugins, tag: result.tag, publishedCommit: result.commit,
        upstreamCommit: result.upstreamCommit, url: result.url };
    console.log(JSON.stringify(summary, null, 2));
    if (process.env.GITHUB_STEP_SUMMARY) fs.appendFileSync(safeUpdaterPath(process.env.GITHUB_STEP_SUMMARY, { allowMissing: true }), `\n\`\`\`json\n${JSON.stringify(summary, null, 2)}\n\`\`\`\n`);
}
