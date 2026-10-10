import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
process.chdir(root);
const read = f => JSON.parse(fs.readFileSync(f, 'utf8'));
const request = read('audio_request.json');
const audio = read('audio_meta.json');
const footage = read('footage.json');
const ev = read('evidence.json');
const esc = s => String(s).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');
const fix = s => s.replaceAll('M C P', 'MCP').replaceAll('C L I', 'CLI').replaceAll('A I', 'AI');

const duration = 161;
const minutes = Math.round(ev.timeline.totalMinutes);
const calls = ev.toolCalls;
const words = ev.humanInput.promptWords;
const prompt = fs.readFileSync('prompt.txt', 'utf8').trim().split(/\r?\n\s*\r?\n/);
if (prompt.join(' ').split(/\s+/).length !== words) throw new Error('prompt.txt word count does not match evidence.json');

// Scenes and where each clip sits on the video timeline.
const scenes = [
  { id: 's1', title: 'Hook', start: 0, end: 13 },
  { id: 's2', title: 'The ask', start: 13, end: 27 },
  { id: 's3', title: 'Research', start: 27, end: 42 },
  { id: 's4', title: 'Excel MCP Server', start: 42, end: 72 },
  { id: 's5', title: 'Hand-off', start: 72, end: 82 },
  { id: 's6', title: 'PowerPoint MCP Server', start: 82, end: 110 },
  { id: 's7', title: 'The briefing', start: 110, end: 130 },
  { id: 's8', title: 'Scoreboard', start: 130, end: 147 },
  { id: 's9', title: 'How this video was made', start: 147, end: duration },
];
const clipStart = { ask: 13, research: 27, excel: 42, 'ppt-a': 82, 'ppt-b': 103.5 };
const clips = footage.clips.map(c => ({ ...c, start: clipStart[c.id], speed: (c.srcEnd - c.srcStart) / c.duration }));
for (const c of clips) {
  if (!fs.existsSync(`assets/footage/${c.id}.mp4`)) throw new Error(`Missing assets/footage/${c.id}.mp4. Footage clips are not committed (about 41 MB); regenerate them from the original OBS recording with npm run footage -- <recording>. The published video is https://youtu.be/_z-twdXG2fA.`);
}
const elapsed = (c, t) => Math.max(0, c.srcStart - footage.promptAt + (t - c.start) * c.speed);
const mmss = s => `${String(Math.floor(s / 60)).padStart(2, '0')}:${String(Math.floor(s % 60)).padStart(2, '0')}`;
const pptA = clips.find(c => c.id === 'ppt-a'), pptB = clips.find(c => c.id === 'ppt-b');
const skipFrom = mmss(elapsed(pptA, pptA.start + pptA.duration)), skipTo = mmss(elapsed(pptB, pptB.start));
if (skipFrom !== ev.skippedInVideo.fromElapsed || skipTo !== ev.skippedInVideo.toElapsed) throw new Error(`Skip marker ${skipFrom}-${skipTo} disagrees with evidence.json`);

// Voice line start times (seconds); each line must end before the next starts and inside its scene.
const voiceAt = {
  '01a': .3, '01b': 2.4, '01c': 9.0,
  '02a': 13.6, '02b': 19.2,
  '03a': 27.6, '03b': 30, '03c': 37.9,
  '04a': 42.6, '04b': 47.8, '04c': 56.6, '04d': 62.5,
  '05a': 72.6, '05b': 78,
  '06a': 82.6, '06b': 85.2, '06c': 92.4, '06d': 104.2,
  '07a': 110.5, '07b': 113.4, '07c': 117, '07d': 120.8,
  '08a': 130.6, '08b': 137,
  '09a': 147.6, '09b': 151.4,
};
const cues = [];
let lastEnd = 0;
for (const line of request.lines) {
  const v = audio.voices.find(x => x.id === line.id);
  if (!v || !fs.existsSync(v.path)) throw new Error(`Missing voice ${line.id}`);
  const start = voiceAt[line.id], end = +(start + v.duration_s).toFixed(3);
  const scene = scenes.find(s => start >= s.start && start < s.end);
  if (start < lastEnd + .2) throw new Error(`Voice ${line.id} overlaps the previous line`);
  if (end > scene.end - .2) throw new Error(`Voice ${line.id} runs past scene ${scene.title}`);
  cues.push({ id: line.id, scene: scene.id, text: fix(line.text), start, end, path: v.path, duration: v.duration_s });
  lastEnd = end;
}

const kicker = (n, name) => `<div class="kicker"><b>${n} /</b> ${name}</div>`;
const at = t => `data-at="${t}"`;
const stage = (t, label) => `<li class="stage" data-on="${t}">${label}</li>`;
const video = (id, extra = '') => {
  const c = clips.find(x => x.id === id);
  return `<video id="v-${id}" src="assets/footage/${id}.mp4" data-start="${c.start}" data-duration="${c.duration}" data-track-index="1" muted playsinline${extra}></video>`;
};
const badges = (cls, ids) => `<div class="badges ${cls}"><div class="badge"><span>Elapsed</span><strong class="clock" data-clips="${ids}">00:00</strong></div><div class="badge"><span>Speed</span><strong class="speed" data-clips="${ids}">1x</strong></div></div>`;
const slideImg = (n, size = '') => `assets/slides/${size}Slide${n}.png`;

const fan = [6, 5, 4, 3, 2, 1, 0].map(i => `<div class="fan" id="fan${i + 1}" style="left:${960 + (6 - i) * 24}px;top:${130 + (6 - i) * 44}px;z-index:${7 - i}"><img src="${slideImg(i + 1, 'fan/')}" alt="Slide ${i + 1} of the agent's briefing"></div>`).join('');

const html_scenes = {
  s1: `${fan}
 <div class="hook-copy"><div class="hook-label">Agentic workflow</div><h1 id="hook-title">One request<br>to an <em>AI agent.</em></h1>
 <p class="lead" ${at(2.4)}>It researched World Bank data, analyzed it in Excel, then built this PowerPoint briefing.</p></div>
 <div class="hook-flow" ${at(4.6)}><span>Research</span><i>→</i><span class="xl">Excel MCP Server</span><i>→</i><span class="pp">PowerPoint MCP Server</span></div>
 <div class="hook-stat" ${at(9.0)}>1 request · ${minutes} minutes · ${ev.humanInput.interventions} follow-ups</div>`,
  s2: `${kicker('01', 'The ask')}
 <div class="copy"><h2>One request,<br><em>in plain English.</em></h2></div>
 <div class="prompt" ${at(13.6)}><span class="prompt-label">The exact prompt</span>${prompt.map(p => `<p>${esc(p)}</p>`).join('')}</div>
 <div class="facts" style="top:850px"><div class="mono" ${at(19.2)} style="font-size:26px;color:#22382b">${words} words · sent once · ${ev.humanInput.interventions} follow-ups</div></div>
 <div class="card card-tall">${video('ask')}</div><div class="card-label">Real footage · GitHub Copilot CLI</div>
 ${badges('top-badges', 'ask')}`,
  s3: `${kicker('02', 'Research')}
 <div class="copy"><h2>First, find<br><em>the right data.</em></h2></div>
 <ul class="facts">
  <li class="fact" ${at(30)}><b>01</b>World Bank data service, official indicators</li>
  <li class="fact" ${at(32.4)}><b>02</b>Income, growth, life expectancy, internet</li>
  <li class="fact" ${at(34.8)}><b>03</b>25 economies × 25 years</li>
  <li class="fact" ${at(37.9)}><b>04</b>Coverage checked for gaps</li>
 </ul>
 <div class="card card-tall">${video('research')}</div><div class="card-label">Real footage · GitHub Copilot CLI</div>
 ${badges('top-badges', 'research')}`,
  s4: `${kicker('03', 'Excel MCP Server')}
 <div class="rail"><h2>Build the<br><em>analysis.</em></h2><ul class="stages">
  ${stage(42.9, 'Power Query')}${stage(47.6, 'Data Model + DAX')}${stage(50.3, 'Analysis + checks')}${stage(56.3, 'Charts')}${stage(62.9, 'PivotTable')}${stage(69.9, 'Refresh all')}
 </ul></div>
 <div class="card card-wide">${video('excel')}</div><div class="card-label">Real footage · Copilot CLI + Excel</div>
 ${badges('rail-badges', 'excel')}`,
  s5: `${kicker('04', 'Hand-off')}
 <div class="handoff"><h2>Every number on a slide<br><em>comes from Excel.</em></h2>
 <div class="numbers">
  <div class="number" ${at(73.6)}><strong>5.9x</strong><span>China's income per person, 2000 to 2024</span></div>
  <div class="number" ${at(74.4)}><strong>23x → 11x</strong><span>richest ÷ poorest income, of the 25</span></div>
  <div class="number" ${at(75.2)}><strong>+4.7 yrs</strong><span>median life expectancy</span></div>
  <div class="number" ${at(76)}><strong>7% → 90%</strong><span>median share of people online</span></div>
 </div>
 <div class="flow" ${at(78)}><span class="file xl">${esc(ev.outputs.workbook)}</span><i>the agent reads its results back</i><span class="file pp">${esc(ev.outputs.deck.replace(/ \(.*\)/, ''))}</span></div></div>`,
  s6: `${kicker('05', 'PowerPoint MCP Server')}
 <div class="rail"><h2>Design the<br><em>briefing.</em></h2><ul class="stages">
  ${stage(85.2, 'Template slide')}${stage(94.3, 'Export and look')}${stage(95.8, 'Copy the design')}${stage(101.6, 'Native charts')}${stage(103.5, 'Recover + check')}
 </ul></div>
 <div class="card card-wide">${video('ppt-a')}${video('ppt-b')}</div><div class="card-label">Real footage · Copilot CLI + PowerPoint</div>
 <div class="skip" id="skip"><strong>Skipped ${skipFrom} → ${skipTo} elapsed</strong>The chart data window got stuck (PowerPoint MCP issue #108), then PowerPoint crashed. The agent reopened the deck and rebuilt the lost charts.</div>
 ${badges('rail-badges', 'ppt-a,ppt-b')}`,
  s7: `${kicker('06', 'The briefing')}
 ${[[1, 110, 113.2, 'Executive summary'], [2, 113.2, 116.8, 'Prosperity'], [3, 116.8, 120.6, 'The income gap'], [6, 120.6, 124.6, 'Connectivity']].map(([n, a, b, name]) =>
    `<div class="slide" id="slide${n}" data-in="${a}" data-out="${b}"><div class="frame"><img src="${slideImg(n)}" alt="Slide ${n}: ${name}"></div><div class="slide-note">Slide ${n} of 7 · ${name} · exported from the agent's deck</div></div>`).join('')}
 <div class="grid" id="grid">${[1, 2, 3, 4, 5, 6, 7].map(n => `<img src="${slideImg(n, 'grid/')}" alt="Slide ${n}">`).join('')}<div class="grid-note">7 slides. Native, editable charts. Sources and speaker notes.</div></div>`,
  s8: `${kicker('07', 'Scoreboard')}
 <div class="score"><h2>One request.<br><em>The heavy lifting, done.</em></h2>
 <div class="stats">
  <div class="stat" ${at(130.8)}><strong>1</strong><span>request, ${words} words, ${ev.humanInput.interventions} follow-ups</span></div>
  <div class="stat" ${at(131.3)}><strong>${minutes} min</strong><span>from prompt to finished deck</span></div>
  <div class="stat" ${at(132.2)}><strong>${calls.total}</strong><span>tool calls by the agent</span></div>
  <div class="stat xl" ${at(133)}><strong>${calls.excel}</strong><span>Excel MCP Server calls</span></div>
  <div class="stat pp" ${at(133.6)}><strong>${calls.powerpoint}</strong><span>PowerPoint MCP Server calls</span></div>
  <div class="stat" ${at(134.6)}><strong>2 + 1</strong><span>MCP servers and one AI model</span></div>
 </div>
 <div class="review" ${at(137.2)}>Headline numbers checked against the World Bank. <em>You still review the result.</em></div>
 <div class="links" ${at(140)}><span class="xl">excelmcpserver.dev</span><span class="pp">powerpointmcpserver.dev</span><span class="plain">Open source · MIT</span></div></div>`,
  s9: `<div class="roll" id="roll"><h2>How this video<br><em>was made.</em></h2>
 ${[
    ['The run', [
      ['Agent', 'GitHub Copilot CLI, autopilot'],
      ['AI model', 'Claude Opus 5.5'],
      ['Tools', `${ev.servers.excel} · ${ev.servers.powerpoint}`],
      ['Human input', `${ev.humanInput.prompts} request, ${words} words, ${ev.humanInput.interventions} follow-ups`],
      ['Footage', 'One screen recording; speed-ups and the cut are labelled'],
    ]],
    ['This video', [
      ['Recording, editing, script', 'An AI agent: GitHub Copilot CLI with Claude Opus 5.5'],
      ['Video engine', 'HyperFrames, open source'],
      ['Voice', 'AI voice (Kokoro), generated on the same PC'],
      ['Human role', 'Gave direction and approved the result'],
    ]],
    ['Data', [
      ['Source', 'World Bank World Development Indicators, CC BY 4.0'],
      ['Workbook and deck', 'excelmcpserver.dev/samples/world-bank-briefing'],
    ]],
  ].map(([head, rows]) => `<div class="roll-group"><h3>${head}</h3>${rows.map(([k, v]) => `<div class="roll-row"><span>${esc(k)}</span><b>${esc(v)}</b></div>`).join('')}</div>`).join('')}</div>`,
};

const clipData = clips.map(({ id, start, duration: d, speed, srcStart }) => ({ id, start, end: start + d, speed: +speed.toFixed(4), srcStart }));
const tracks = cues.map(c => `<audio id="voice-${c.id}" src="${c.path}" data-start="${c.start}" data-duration="${c.duration}" data-track-index="10" data-volume="1"></audio>`);
const css = fs.readFileSync('style.css', 'utf8');
const html = `<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=1920,height=1080">
<title>Agentic workflow: World Bank data to Excel and PowerPoint</title><script src="assets/vendor/gsap.min.js"></script><style>${css}</style></head>
<body><main id="root" data-composition-id="main" data-width="1920" data-height="1080" data-start="0" data-duration="${duration}">
${scenes.map(s => `<section id="${s.id}" class="scene${s.id === 's6' ? ' ppt' : s.id === 's9' ? ' credits' : ''}" data-layout-allow-overflow>${html_scenes[s.id]}</section>`).join('\n')}
<div id="progress" data-layout-ignore></div>
${tracks.join('\n')}
</main><script>
window.__timelines = window.__timelines || {};
const tl = gsap.timeline({ paused: true });
window.__timelines["main"] = tl;
const scenes = ${JSON.stringify(scenes.map(({ id, start, end }) => ({ id, start, end })))};
const clips = ${JSON.stringify(clipData)};
const promptAt = ${footage.promptAt};
tl.fromTo("#progress", { scaleX: 0 }, { scaleX: 1, duration: ${duration}, ease: "none" }, 0);
scenes.forEach(({ id, start, end }, i) => {
  tl.fromTo("#" + id, { autoAlpha: 0 }, { autoAlpha: 1, duration: i ? .45 : .01, ease: "power2.out", immediateRender: true }, start);
  if (end < ${duration}) tl.to("#" + id, { autoAlpha: 0, duration: .35, ease: "power2.in" }, end - .35);
});
document.querySelectorAll("[data-at]").forEach(el => {
  tl.fromTo(el, { autoAlpha: 0, y: 26 }, { autoAlpha: 1, y: 0, duration: .6, ease: "power3.out" }, +el.dataset.at);
});
tl.fromTo("#hook-title", { autoAlpha: 0, y: 40 }, { autoAlpha: 1, y: 0, duration: .8, ease: "power3.out" }, .2);
[1, 2, 3, 4, 5, 6, 7].reverse().forEach((n, i) => {
  tl.fromTo("#fan" + n, { autoAlpha: 0, y: 90, rotation: (n % 2 ? -4 : 4) }, { autoAlpha: 1, y: 0, rotation: (n - 4) * .8, duration: .7, ease: "power3.out" }, .3 + i * .22);
});
document.querySelectorAll(".slide").forEach(el => {
  const a = +el.dataset.in, b = +el.dataset.out;
  tl.fromTo(el, { autoAlpha: 0, scale: 1.03 }, { autoAlpha: 1, scale: 1, duration: .5, ease: "power2.out" }, a);
  tl.to(el, { autoAlpha: 0, duration: .35, ease: "power2.in" }, b - .2);
});
tl.fromTo("#grid", { autoAlpha: 0, y: 30 }, { autoAlpha: 1, y: 0, duration: .6, ease: "power3.out" }, 124.6);
tl.fromTo("#grid img", { autoAlpha: 0 }, { autoAlpha: 1, duration: .35, stagger: .12 }, 124.7);
tl.fromTo("#skip", { autoAlpha: 0, y: 20 }, { autoAlpha: 1, y: 0, duration: .45, ease: "power3.out" }, ${pptB.start});
tl.fromTo("#roll", { y: 80 }, { y: -40, duration: ${duration - scenes.at(-1).start}, ease: "none" }, ${scenes.at(-1).start});
const pad = n => String(Math.floor(n)).padStart(2, "0");
const clocks = [...document.querySelectorAll(".clock")], speeds = [...document.querySelectorAll(".speed")], stages = [...document.querySelectorAll(".stage")];
function activeClip(ids, t) {
  const list = clips.filter(c => ids.includes(c.id));
  return list.find(c => t >= c.start && t < c.end) || (t < list[0].start ? list[0] : list[list.length - 1]);
}
function render() {
  const t = tl.time();
  clocks.forEach(el => {
    const c = activeClip(el.dataset.clips.split(","), t);
    const e = Math.max(0, c.srcStart - promptAt + (Math.min(Math.max(t, c.start), c.end) - c.start) * c.speed);
    el.textContent = pad(e / 60) + ":" + pad(e % 60);
  });
  speeds.forEach(el => {
    const c = activeClip(el.dataset.clips.split(","), t);
    el.textContent = c.speed < 1.05 ? "real time" : (c.speed >= 10 ? Math.round(c.speed) : c.speed.toFixed(1)) + "x faster";
  });
  stages.forEach(el => el.classList.toggle("on", t >= +el.dataset.on));
}
tl.eventCallback("onUpdate", render);
render();
</script></body></html>`;
fs.writeFileSync('index.html', html);

const clock = t => new Date(Math.round(t * 1000)).toISOString().slice(11, 23);
fs.writeFileSync('agentic-world-bank-briefing.vtt', 'WEBVTT\n\n' + cues.map((c, i) => `${i + 1}\n${clock(c.start)} --> ${clock(c.end)}\n${c.text}\n`).join('\n'));
fs.writeFileSync('captions.json', JSON.stringify(cues.map(({ id, text, start, end }) => ({ id, text, start, end })), null, 2) + '\n');
fs.writeFileSync('schedule.json', JSON.stringify({ duration, scenes, clips: clipData, voice: cues.map(({ id, start, end }) => ({ id, start, end })) }, null, 2) + '\n');
const lines = id => cues.filter(c => c.scene === id).map(c => c.text).join(' ');
fs.writeFileSync('SCRIPT.md', `# Agentic workflow: World Bank data to Excel and PowerPoint\n\n**Voice:** Local Kokoro / af_heart (warm female English). No music; footage muted.\n\n**Duration:** ${duration} seconds. Captions align to individually synthesized lines.\n\n` +
  scenes.map((s, i) => `## ${i + 1}. ${s.title}\n\n**Time:** ${s.start}–${s.end}s\n\n    ${lines(s.id)}\n`).join('\n'));
const visuals = {
  s1: 'The seven exported slides fan in; "One request to an AI agent." under an Agentic workflow label, with the three steps and the measured effort.',
  s2: 'The exact prompt, word count, and real footage of it being sent in Copilot CLI.',
  s3: 'Research steps beside sped-up terminal footage; elapsed clock and speed.',
  s4: 'Sped-up footage of Copilot CLI and Excel; Excel steps light up as they happen.',
  s5: 'Headline values from the workbook and the hand-off from workbook to deck.',
  s6: 'Sped-up footage of Copilot CLI and PowerPoint; labelled skip over the stuck chart window and crash.',
  s7: 'Four exported slides full-screen, then all seven.',
  s8: 'Measured effort from the session record, review reminder, links.',
  s9: 'Dark end-credits roll: how the run and this video were made, and the data source.',
};
fs.writeFileSync('STORYBOARD.md', `---\nformat: 1920x1080\nduration: ${duration}s\nmessage: "An agentic workflow: one request, two MCP servers, public data to a board-ready briefing."\narc: Hook → Ask → Research → Excel → Hand-off → PowerPoint → Briefing → Scoreboard → Credits\naudience: people curious about AI agents doing real office work\nmode: autonomous\n---\n\n` +
  scenes.map((s, i) => `## Frame ${i + 1} — ${s.title}\n\n- status: animated\n- src: index.html\n- duration: ${s.end - s.start}s\n- poster: ${s.start + 3}s\n- transition_in: fade\n- scene: ${visuals[s.id]}\n- voiceover: ${lines(s.id)}\n`).join('\n'));
console.log(`Built ${duration}s composition: ${scenes.length} scenes, ${clips.length} clips, ${cues.length} voice lines (last ends ${lastEnd}s).`);
