// Shared deterministic rules for publication and marketplace listings. No network or writes.
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

export const pluginNames = ['excel-cli', 'excel-mcp'];
export const publishedRepo = 'sbroenne/mcp-server-excel-plugins';
const schema = 'https://agent-plugins.org/schemas/1.0.0/plugin.schema.json';
const decoder = new TextDecoder('utf-8', { fatal: true });

export function canonical(value) {
    if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`;
    if (value !== null && typeof value === 'object') {
        return `{${Object.keys(value).sort().map(key => `${JSON.stringify(key)}:${canonical(value[key])}`).join(',')}}`;
    }
    return JSON.stringify(value);
}

export function parseJson(bytes) {
    const text = decoder.decode(Buffer.from(bytes)).replace(/^\uFEFF/, '');
    const value = JSON.parse(text);
    // Reject values JSON.parse would silently discard or round before comparison.
    const tokens = text.match(/"(?:\\.|[^"\\])*"|-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?|true|false|null|[{}\[\]:,]/g);
    let index = 0;
    function decimal(token) {
        const [, sign, whole, fraction = '', power = '0'] = /^(-?)(\d+)(?:\.(\d+))?(?:[eE]([+-]?\d+))?$/.exec(token);
        let digits = (whole + fraction).replace(/^0+/, '');
        if (!digits) return '0';
        let exponent = BigInt(power) - BigInt(fraction.length);
        while (digits.endsWith('0')) { digits = digits.slice(0, -1); exponent++; }
        return `${sign}${digits}e${exponent}`;
    }
    function visit() {
        const token = tokens[index++];
        if (token === '{') {
            const keys = new Set();
            while (tokens[index] !== '}') {
                const key = JSON.parse(tokens[index++]);
                if (keys.has(key)) throw new Error(`Duplicate JSON key: ${key}`);
                keys.add(key);
                index++; // colon
                visit();
                if (tokens[index] !== ',') break;
                index++;
            }
            index++;
        } else if (token === '[') {
            while (tokens[index] !== ']') {
                visit();
                if (tokens[index] !== ',') break;
                index++;
            }
            index++;
        } else if (/^-?\d/.test(token)) {
            const number = Number(token);
            if (!Number.isFinite(number) || decimal(token) !== decimal(String(number))) {
                throw new Error('JSON numeric precision would be lost during canonical comparison.');
            }
        }
    }
    visit();
    return value;
}

export function canonicalJson(text) { return canonical(parseJson(Buffer.from(text))); }
export function hash(bytes) { return createHash('sha256').update(bytes).digest('hex'); }

export function assertVersion(version) {
    if (typeof version !== 'string' || !/^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$/.test(version)) {
        throw new Error(`Invalid release version: ${version}`);
    }
    return version;
}

export function compareVersions(left, right) {
    const a = assertVersion(left).split('.').map(BigInt), b = assertVersion(right).split('.').map(BigInt);
    for (let i = 0; i < 3; i++) if (a[i] !== b[i]) return a[i] > b[i] ? 1 : -1;
    return 0;
}

export function assertTag(tag) {
    if (typeof tag !== 'string' || !tag.startsWith('v')) throw new Error('An existing published vX.Y.Z tag is required.');
    assertVersion(tag.slice(1));
    return tag;
}

export function git(repo, args, options = {}) {
    return execFileSync('git', ['-C', repo, ...args], {
        maxBuffer: 64 * 1024 * 1024, windowsHide: true, ...options,
    });
}

export function resolveCommit(repo, ref) {
    if (!/^[a-f0-9]{40}$/.test(ref) && !/^refs\/tags\/v(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$/.test(ref) &&
        !['HEAD', 'refs/remotes/origin/main'].includes(ref)) throw new Error('Unsupported commit locator.');
    const sha = git(repo, ['rev-parse', '--verify', `${ref}^{commit}`]).toString().trim();
    if (!/^[a-f0-9]{40}$/.test(sha)) throw new Error('Invalid resolved commit.');
    return sha;
}

export function readGitTree(repo, ref) {
    if (!/^[a-f0-9]{40}$/.test(ref) && ref !== 'HEAD') throw new Error('An exact commit/tree is required.');
    const entries = git(repo, ['ls-tree', '-r', '-z', ref]).toString('utf8').split('\0').filter(Boolean);
    if (!entries.length) throw new Error('Publication tree is empty.');
    const result = new Map();
    const blobs = [];
    for (const entry of entries) {
        const match = /^(\d+) blob ([a-f0-9]{40})\t(.+)$/.exec(entry);
        if (!match || !['100644', '100755'].includes(match[1])) throw new Error('Links/submodules are not allowed in a publication.');
        const [, mode, sha, name] = match;
        if (name.split('/').some(part => !part || part === '.' || part === '..' || part === '.git') || /[\r\n\\]/.test(name)) {
            throw new Error('Unsupported publication path.');
        }
        blobs.push({ mode, sha, name });
    }
    const content = git(repo, ['cat-file', '--batch'], { input: blobs.map(blob => blob.sha).join('\n') + '\n' });
    let offset = 0;
    for (const { mode, sha, name } of blobs) {
        const end = content.indexOf(10, offset);
        const header = content.subarray(offset, end).toString();
        const match = /^([a-f0-9]{40}) blob (\d+)$/.exec(header);
        if (!match || match[1] !== sha) throw new Error('Invalid Git blob response.');
        const length = Number(match[2]);
        offset = end + 1;
        if (offset + length >= content.length || content[offset + length] !== 10) throw new Error('Truncated Git blob response.');
        result.set(name, { mode, bytes: content.subarray(offset, offset + length) });
        offset += length + 1;
    }
    if (offset !== content.length) throw new Error('Unexpected Git blob response data.');
    return result;
}

function requireFile(tree, name) {
    const file = tree.get(name);
    if (!file || !['100644', '100755'].includes(file.mode)) throw new Error(`Missing or unsupported file: ${name}`);
    return file.bytes;
}

export function validatePlugin(tree, name, { expectedVersion, repairStamps = false } = {}) {
    if (!pluginNames.includes(name)) throw new Error('Unsupported plugin identity.');
    const prefix = `plugins/${name}/`;
    const manifest = parseJson(requireFile(tree, `${prefix}plugin.json`));
    if (!manifest || manifest.name !== name || manifest.$schema !== schema) throw new Error(`Invalid ${name} manifest identity/schema.`);
    assertVersion(manifest.version);
    const version = assertVersion(expectedVersion ?? manifest.version);
    const repairs = [];
    if (manifest.version !== version) {
        if (!repairStamps) throw new Error(`Wrong manifest version in ${name}.`);
        repairs.push(`${prefix}plugin.json`);
    }
    for (const field of ['description', 'repository', 'license']) {
        if (typeof manifest[field] !== 'string' || !manifest[field].trim()) throw new Error(`Invalid ${name} manifest ${field}.`);
    }
    if (!manifest.author || typeof manifest.author.name !== 'string' || !manifest.author.name.trim() ||
        !Array.isArray(manifest.keywords) || manifest.keywords.some(value => typeof value !== 'string')) {
        throw new Error(`Invalid ${name} manifest metadata.`);
    }
    for (const file of ['README.md', `skills/${name}/SKILL.md`, `skills/${name}/references/range.md`,
        name === 'excel-cli' ? 'bin/start-cli.ps1' : 'mcp.json']) requireFile(tree, prefix + file);
    for (const stamp of ['version.txt', `skills/${name}/VERSION`]) {
        const value = tree.get(prefix + stamp);
        if (!value || decoder.decode(value.bytes).trim() !== version) {
            if (!repairStamps) throw new Error(`Missing or mismatched ${name}/${stamp}.`);
            repairs.push(prefix + stamp);
        }
    }
    for (const [file, value] of tree) {
        if (!file.startsWith(prefix)) continue;
        if (/\.(exe|dll|pdb)$/i.test(file) || /\.(deps|runtimeconfig)\.json$/i.test(file)) throw new Error('Bundled runtimes are prohibited.');
        if (file.endsWith('.json')) parseJson(value.bytes);
    }
    return { manifest, version, repairs };
}

export function validatePublication(tree, options = {}) {
    for (const file of tree.keys()) {
        if (file.startsWith('plugins/') && !pluginNames.some(name => file.startsWith(`plugins/${name}/`))) {
            throw new Error('Unsupported publication plugin layout.');
        }
    }
    const marketplacePath = tree.has('.github/plugin/marketplace.json') ? '.github/plugin/marketplace.json' : 'marketplace.json';
    const marketplace = parseJson(requireFile(tree, marketplacePath));
    if (!Array.isArray(marketplace.plugins) || marketplace.plugins.length !== 2) throw new Error('Expected exactly two marketplace plugins.');
    const versions = new Set();
    const repairs = [];
    for (const name of pluginNames) {
        const matches = marketplace.plugins.filter(plugin => plugin.name === name);
        if (matches.length !== 1 || matches[0].source !== `./plugins/${name}`) throw new Error('Invalid marketplace plugin identity/source.');
        const version = assertVersion(matches[0].version);
        const validated = validatePlugin(tree, name, { expectedVersion: version, ...options });
        versions.add(version);
        repairs.push(...validated.repairs);
    }
    if (versions.size !== 1) throw new Error('Published marketplace must have one release version.');
    return { version: [...versions][0], marketplacePath, repairs };
}

function normalized(tree, { publication = true, plugin, repairStamps = false } = {}) {
    if (publication) validatePublication(tree, { repairStamps });
    else validatePlugin(tree, plugin);
    const result = new Map();
    for (const [name, file] of tree) {
        if (plugin && !name.startsWith(`plugins/${plugin}/`)) continue;
        if (pluginNames.some(value => name === `plugins/${value}/version.txt` || name === `plugins/${value}/skills/${value}/VERSION`)) continue;
        let bytes = file.bytes;
        if (name.endsWith('.json')) {
            const json = parseJson(bytes);
            if (pluginNames.some(value => name === `plugins/${value}/plugin.json`)) delete json.version;
            if (publication && ['.github/plugin/marketplace.json', 'marketplace.json'].includes(name)) {
                for (const entry of json.plugins) if (pluginNames.includes(entry.name)) delete entry.version;
            }
            bytes = Buffer.from(canonical(json));
        }
        result.set(name, `${file.mode}:${hash(bytes)}`);
    }
    return result;
}

export function compareTrees(baseline, candidate, options = {}) {
    const before = normalized(baseline, options);
    const after = normalized(candidate, { ...options, repairStamps: false });
    const changedPaths = [...new Set([...before.keys(), ...after.keys()])].sort()
        .filter(name => before.get(name) !== after.get(name));
    return {
        changedPaths,
        changedPlugins: pluginNames.filter(name => changedPaths.some(file => file.startsWith(`plugins/${name}/`))),
        baselineFingerprint: hash(canonical(Object.fromEntries(before))),
        candidateFingerprint: hash(canonical(Object.fromEntries(after))),
    };
}

export function publicationDecision(comparison, { currentVersion, version, tagExists = false, manualRepair = false, rawChanged = false }) {
    if (compareVersions(currentVersion, version) > 0) throw new Error('Downgrade publish blocked.');
    if (tagExists && currentVersion !== version) throw new Error('Existing tag conflicts with marketplace version.');
    const write = manualRepair
        ? comparison.changedPaths.length > 0 || rawChanged || !tagExists
        : comparison.changedPaths.length > 0;
    return { status: write ? 'published' : 'skipped', write, handoff: write && comparison.changedPlugins.length > 0 };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
    const [baselineRepo, candidateRepo, candidateTree, version, manual, exists] = process.argv.slice(2);
    const baseline = readGitTree(baselineRepo, 'HEAD'), candidate = readGitTree(candidateRepo, candidateTree);
    const current = validatePublication(baseline, { repairStamps: manual === 'true' });
    const prepared = validatePublication(candidate);
    if (prepared.version !== assertVersion(version)) throw new Error('Candidate version mismatch.');
    const comparison = compareTrees(baseline, candidate, { repairStamps: manual === 'true' });
    const rawChanged = canonical([...baseline].map(([name, file]) => [name, file.mode, hash(file.bytes)]).sort()) !==
        canonical([...candidate].map(([name, file]) => [name, file.mode, hash(file.bytes)]).sort());
    console.log(JSON.stringify({
        ...comparison, currentVersion: current.version, version, rawChanged, repairs: current.repairs,
        ...publicationDecision(comparison, {
            currentVersion: current.version, version, manualRepair: manual === 'true', tagExists: exists === 'true', rawChanged,
        }),
    }));
}
