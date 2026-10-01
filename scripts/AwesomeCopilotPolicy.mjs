import {
    pluginNames, publishedRepo, canonical, hash, assertTag, assertVersion,
    compareVersions, compareTrees, validatePlugin, validatePublication, parseJson,
} from './PluginContent.mjs';

export const upstreamRepo = 'github/awesome-copilot';
export const forkRepo = 'sbroenne/awesome-copilot';
export const prAuthor = 'sbroenne';
export const allowedPaths = ['plugins/external.json', '.github/plugin/marketplace.json'];
const marker = 'excel-plugin-update-state:';
const metadata = ['description', 'author', 'repository', 'homepage', 'license', 'keywords'];
const shaPattern = /^[a-f0-9]{40}$/;

export function stateMarker(state) {
    return `<!-- ${marker}b64url:${Buffer.from(canonical(state)).toString('base64url')} -->`;
}

export function parseState(body) {
    if (typeof body !== 'string') throw new Error('Missing owned PR state.');
    const matches = [...body.matchAll(/<!-- excel-plugin-update-state:(.*?) -->/gs)];
    if (matches.length !== 1) throw new Error('Expected exactly one owned PR state marker.');
    const payload = matches[0][1];
    let text;
    if (payload.startsWith('b64url:')) {
        const encoded = payload.slice('b64url:'.length);
        if (!/^[A-Za-z0-9_-]+$/.test(encoded)) throw new Error('Invalid owned state encoding.');
        const bytes = Buffer.from(encoded, 'base64url');
        if (bytes.toString('base64url') !== encoded) throw new Error('Noncanonical owned state encoding.');
        text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    } else {
        if (!payload.startsWith('{') || /--|[<>]/.test(payload)) {
            throw new Error('Unsafe legacy owned state; manual verified migration required.');
        }
        text = payload;
    }
    const state = parseJson(Buffer.from(text));
    if (canonical(state) !== text) throw new Error('Noncanonical owned state JSON.');
    if (state.schema !== 1 || !state.entries || typeof state.entries !== 'object' || Array.isArray(state.entries) ||
        !Object.keys(state.entries).length || Object.keys(state.entries).some(name => !pluginNames.includes(name))) {
        throw new Error('Invalid owned PR proposal entries.');
    }
    for (const [name, proposed] of Object.entries(state.entries)) {
        validateListing(proposed.entry, name);
        if (!/^[a-f0-9]{64}$/.test(proposed.fingerprint)) throw new Error('Invalid proposal fingerprint.');
    }
    if (state.proposalFingerprint !== proposalFingerprint(state.entries)) throw new Error('Invalid combined proposal fingerprint.');
    if (!/^[a-f0-9]{64}$/.test(state.bodyFingerprint ?? '')) {
        throw new Error('Missing or invalid PR body fingerprint; human edits are protected.');
    }
    if (state.bodyFingerprint !== hash(body.replace(/<!-- excel-plugin-update-state:.*? -->/gs, '').trimEnd())) {
        throw new Error('PR body changed since the last automated write; human edits are protected.');
    }
    return state;
}

export function validateListing(entry, name) {
    if (!entry || entry.name !== name) throw new Error('Invalid listed plugin identity.');
    assertVersion(entry.version);
    const source = entry.source;
    if (!source || source.source !== 'github' || source.repo !== publishedRepo || source.path !== `plugins/${name}` ||
        (!source.ref && !source.sha) || (source.sha && !shaPattern.test(source.sha))) {
        throw new Error('Invalid listing published source locator.');
    }
    if (source.ref) assertTag(source.ref);
}

function proposalFingerprint(entries) {
    return hash(canonical(Object.fromEntries(Object.entries(entries).map(([name, item]) => [name, item.fingerprint]))));
}

export function ownedPulls(pulls) {
    if (!Array.isArray(pulls)) throw new Error('Invalid pull request API response.');
    return pulls.filter(pr => {
        if (!pr || !Number.isSafeInteger(pr.number) || !['open', 'closed'].includes(pr.state) ||
            !pr.user?.login || (pr.body !== null && typeof pr.body !== 'string')) {
            throw new Error('Incomplete pull request API data.');
        }
        const ours = pr.head?.repo?.full_name === forkRepo &&
            /^excel-plugin-updates-[a-f0-9]{12}$/.test(pr.head?.ref);
        if (!ours) {
            if (pr.user.login === prAuthor && pr.body?.includes(marker)) {
                throw new Error('Automated PR ownership data is missing or inconsistent.');
            }
            return false;
        }
        if (!shaPattern.test(pr.head.sha) || pr.user.login !== prAuthor ||
            pr.base?.repo?.full_name !== upstreamRepo || pr.base.ref !== 'main' || !pr.body?.includes(marker)) {
            throw new Error('Automated branch PR ownership mismatch.');
        }
        parseState(pr.body);
        return true;
    });
}

export function planListings({ listings, tag, commit, getTree, resolveTag, pulls }) {
    assertTag(tag);
    if (!shaPattern.test(commit) || resolveTag(tag) !== commit) throw new Error('Missing or mismatched published candidate tag.');
    const candidateTree = getTree(commit);
    if (validatePublication(candidateTree).version !== tag.slice(1)) throw new Error('Candidate tag/version mismatch.');
    if (!Array.isArray(listings)) throw new Error('Invalid listing API data.');
    const owned = ownedPulls(pulls);
    const open = owned.filter(pr => pr.state === 'open');
    if (open.length > 1) throw new Error('At most one open automated update PR is allowed.');
    const pending = open[0];
    const previous = pending ? parseState(pending.body) : null;
    if (pending && previous.head !== pending.head.sha) throw new Error('PR head changed since the last automated write; human changes are protected.');
    const entries = structuredClone(previous?.entries ?? {});
    const listed = {};
    const evidence = {};
    for (const name of pluginNames) {
        const matches = listings.filter(entry => entry.name === name);
        if (matches.length !== 1) throw new Error(`Expected exactly one existing ${name} listing.`);
        const entry = matches[0];
        validateListing(entry, name);
        if (compareVersions(tag.slice(1), entry.version) < 0 ||
            (entries[name] && compareVersions(tag.slice(1), entries[name].entry.version) < 0)) {
            throw new Error(`Downgrade listing/pending update blocked: ${name}.`);
        }
        if (entries[name]) {
            const proposed = entries[name];
            const source = proposed.entry.source;
            const sha = source.ref ? resolveTag(source.ref) : source.sha;
            if (!shaPattern.test(sha ?? '') || (source.sha && source.sha !== sha) ||
                validatePlugin(getTree(sha), name).version !== proposed.entry.version) {
                throw new Error('Pending published locator mismatch.');
            }
            const fingerprint = compareTrees(getTree(sha), getTree(sha), { publication: false, plugin: name }).candidateFingerprint;
            if (fingerprint !== proposed.fingerprint) throw new Error('Pending content fingerprint mismatch.');
        }
        const baselineCommit = entry.source.ref ? resolveTag(entry.source.ref) : entry.source.sha;
        if (!shaPattern.test(baselineCommit ?? '') || (entry.source.sha && baselineCommit !== entry.source.sha)) {
            throw new Error(`Listed ref/SHA mismatch: ${name}.`);
        }
        const baselineTree = getTree(baselineCommit);
        const baseline = validatePlugin(baselineTree, name);
        if (baseline.version !== entry.version) throw new Error('Listed manifest/version mismatch.');
        if (previous?.listed?.[name] && canonical(previous.listed[name]) !== canonical(entry)) {
            throw new Error(`Upstream changed the pending ${name} listing; manual conflict resolution required.`);
        }
        listed[name] = structuredClone(entry);
        const comparison = compareTrees(baselineTree, candidateTree, { publication: false, plugin: name });
        evidence[name] = { baselineCommit, candidateCommit: commit, ...comparison };
        if (!comparison.changedPaths.length) {
            if (entries[name]) {
                throw new Error(`Pending plugin proposal was reverted: ${name}; manual resolution required.`);
            }
            continue;
        }
        if (entries[name]?.fingerprint === comparison.candidateFingerprint) continue;
        const candidate = validatePlugin(candidateTree, name).manifest;
        const desired = structuredClone(entry);
        desired.version = candidate.version;
        desired.source = { ...entry.source, ref: tag, sha: commit };
        for (const field of metadata) {
            if (canonical(baseline.manifest[field]) === canonical(candidate[field])) continue;
            if (candidate[field] === undefined) delete desired[field];
            else desired[field] = structuredClone(candidate[field]);
        }
        entries[name] = { entry: desired, fingerprint: comparison.candidateFingerprint };
    }
    const state = { schema: 1, entries, listed, proposalFingerprint: proposalFingerprint(entries) };
    if (!Object.keys(entries).length) return { action: 'noop', reason: 'Currently listed plugin content is equivalent.', evidence, changedPlugins: [] };
    if (pending && state.proposalFingerprint === previous.proposalFingerprint) {
        return { action: 'noop', reason: 'Pending proposal already contains equivalent content.', evidence, changedPlugins: [] };
    }
    for (const pr of owned.filter(pr => pr.state === 'closed' && !pr.merged_at)) {
        if (parseState(pr.body).proposalFingerprint === state.proposalFingerprint) {
            throw new Error('Identical proposal was declined (closed without merge); human decision required.');
        }
    }
    return {
        action: pending ? 'update' : 'create',
        changedPlugins: Object.keys(entries).sort(), evidence, state,
        pullNumber: pending?.number ?? null,
        expectedHead: pending?.head.sha ?? null,
        expectedBody: pending ? hash(pending.body) : null,
        branch: pending?.head.ref ?? `excel-plugin-updates-${state.proposalFingerprint.slice(0, 12)}`,
        previousState: previous,
        previousBody: pending?.body ?? null,
    };
}

export function assertListingPatch(before, after, affected) {
    if (!Array.isArray(before) || !Array.isArray(after) || before.length !== after.length ||
        new Set(before.map(entry => entry.name)).size !== before.length ||
        new Set(after.map(entry => entry.name)).size !== after.length) {
        throw new Error('Listing patch must preserve all entries, identities and ordering.');
    }
    if (!affected.length || affected.some(name => !pluginNames.includes(name))) throw new Error('Unsupported affected entries.');
    for (let index = 0; index < before.length; index++) {
        const old = before[index], next = after[index];
        if (old.name !== next.name) throw new Error('Listing entries cannot be reordered.');
        if (!affected.includes(old.name) && canonical(old) !== canonical(next)) throw new Error('Patch changes an unaffected listing.');
        if (affected.includes(old.name)) validateListing(next, old.name);
    }
}

export function assertAllowedPaths(paths) {
    if (!Array.isArray(paths) || !paths.length) throw new Error('Listing patch is empty.');
    if (paths.length > 2 || paths.some(file => !allowedPaths.includes(file))) throw new Error('Patch includes a file outside the allowed marketplace files.');
}
