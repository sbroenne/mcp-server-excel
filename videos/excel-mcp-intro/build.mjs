import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
process.chdir(root);
const request = JSON.parse(fs.readFileSync('audio_request.json','utf8'));
const audio = JSON.parse(fs.readFileSync('audio_meta.json','utf8'));
const esc = s => String(s).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;');
const duration = 120;
const starts = [0,9.5,27.5,39.5,48,60,72,84,93,103,112];
const ends = [...starts.slice(1),duration];
const voiceKeys = ['intro','compare','fit','02','03','04','05','06','07','08','09'];
const titles = ['What is Excel MCP Server?','File editing or the real Excel engine?','When should you use it?','One workbook, two ways','Power Query','Data Model and DAX','PivotTables and charts','VBA workflows','Python in Excel','Real calculation','Two equal entry points'];
const scenes = [
`<div class="kicker">Introducing Excel MCP Server</div>
 <div class="hero-copy"><h1 id="hero-title">Your AI.<br><em>Real Excel.</em></h1><p class="lead reveal">Ask in plain English.</p><p class="intro-definition reveal">Connect your AI assistant to<br>Microsoft Excel on your desktop.</p></div>
 <div class="hero-evidence reveal"><div class="proof-label">Real Excel · synthetic sales data</div><img src="assets/sales-data-detail.png" alt="Actual Excel sales table with twelve synthetic orders"></div>
 <div class="foot-label">Claude, Copilot, and other MCP-compatible assistants</div>`,
`<div class="kicker">Understand the difference</div>
 <h2>Same workbook format.<br><em>Different engines.</em></h2>
 <div class="comparison-row">
   <div class="comparison-card reveal"><div class="mono">File-based workflows</div><h3>Anthropic's spreadsheet skill</h3><p>openpyxl edits the file.<br>LibreOffice calculates formulas.</p><h3 class="secondary-heading">openpyxl-based MCP servers</h3><p>Read and write workbook files.</p></div>
   <div class="comparison-card excel-engine reveal"><div class="mono">Excel MCP Server</div><h3>The real Excel application</h3><p>Excel opens, calculates,<br>refreshes, and saves the workbook.</p><div class="engine-features"><span>Power Query</span><span>DAX</span><span>VBA</span></div></div>
 </div><div class="comparison-footer reveal">Not just a different interface. A different engine doing the work.</div>`,
`<div class="kicker">Choose it for the work you actually do</div>
 <h2>Do you need<br><em>real Excel?</em></h2>
 <div class="fit-list"><div class="fit-item">Refreshable business reports</div><div class="fit-item">Existing models, macros, and PivotTables</div><div class="fit-item">Live results from Excel's own engine</div></div>
 <div class="fit-alternative reveal"><div class="mono">A file-based tool may be enough</div><h3>Simple exports.<br>No desktop Excel.</h3><p>Choose a file-based workflow when you need portability, or just a new spreadsheet.</p></div>
 <div class="fit-requirement reveal">Excel MCP requires Windows + desktop Microsoft Excel.</div>`,
`<div class="kicker">01 / Start with a question</div>
 <h2>One workbook.<br><em>Two ways to work.</em></h2>
 <div class="interface-row">
   <div class="interface reveal"><div class="mono">MCP / your AI assistant</div><p class="request">“Turn these sales orders<br>into a regional report.”</p><div class="tag">Conversational workflow</div></div>
   <div class="interface reveal"><div class="mono">excelcli / your terminal</div><p class="request">The same Excel operations.<br>Scripted. Repeatable.</p><div class="tag">Command-line workflow</div></div>
 </div><div class="workbook-chip reveal">sales-workbook.xlsx <span>→</span> real desktop Excel</div>`,
`<div class="kicker">02 / Power Query</div>
 <div class="split-copy"><h2>Refreshable.<br><em>Not disposable.</em></h2><p class="lead">From raw orders<br>to analysis-ready data.</p>
 <div class="steps"><div class="step">1 <span>Read the source</span></div><div class="step">2 <span>Set column types</span></div><div class="step">3 <span>Clean the regions</span></div></div></div>
 <div class="query-proof reveal"><div class="proof-label">Actual Excel output / CleanSales</div><img src="assets/sales-query-detail.png" alt="Power Query output with typed sales data"></div>
 <div class="query-code reveal"><span class="mono">M / excerpt from executed query</span><pre>Table.TransformColumnTypes(Source,<br>  {{"Revenue", Currency.Type}})</pre></div>`,
`<div class="kicker light">03 / Data Model + DAX</div>
 <h2 class="light">One definition.<br><em>Every decision.</em></h2>
 <div class="dax-code reveal"><div class="mono">Persisted measure / CleanSales</div><pre>Total Revenue :=<br>  SUM(CleanSales[Revenue])</pre></div>
 <div class="data-flow reveal"><span>Sales</span><b>→</b><span>CleanSales</span><b>→</b><span>Data Model</span></div>
 <div class="results"><div class="stat"><span>North</span><strong>$194k</strong></div><div class="stat"><span>South</span><strong>$165k</strong></div><div class="stat"><span>West</span><strong>$225k</strong></div></div>
 <div class="foot-label light">Verified DAX query result · 12 synthetic orders</div>`,
`<div class="kicker">04 / PivotTables + charts</div>
 <h2 class="chart-title">See the story.<em> Keep the objects.</em></h2>
 <div class="report-proof reveal"><div class="proof-label">Actual Excel capture / native PivotTable + linked PivotChart</div><img src="assets/sales-report-detail.png" alt="Actual Excel sales report with native PivotTable and green PivotChart"></div>
 <div class="report-callout reveal"><strong>$225k</strong><span>West leads revenue</span></div>`,
`<div class="kicker">05 / VBA</div>
 <div class="split-copy"><h2>Keep your<br><em>proven<br>workflows.</em></h2><p class="lead">Read, edit, and run VBA.<br>Only run macros you trust.</p></div>
 <div class="vba-code reveal"><div class="mono">Illustrative VBA / not executed</div><pre><span class="syntax">Sub</span> RefreshReport()<br>  Worksheets("Report") _<br>    .PivotTables("SalesByRegion") _<br>    .RefreshTable<br>  Application.Calculate<br><span class="syntax">End Sub</span></pre><div class="code-note">Editing VBA projects needs manual trust access.</div></div>`,
`<div class="kicker">06 / Python in Excel</div>
 <h2>Analysis, right<br><em>beside your data.</em></h2>
 <div class="python-proof reveal"><div class="proof-label">Genuine repository screenshot / existing product example</div><img src="assets/excel-demo-python.png" alt="Real Excel formula bar showing Python reading a sales revenue column"></div>
 <div class="python-note reveal"><span class="tag">Microsoft-hosted Python</span><p>Supported Microsoft 365.</p><p>Python enabled.<br>Internet required.</p></div>`,
`<div class="kicker">07 / Real calculation</div>
 <h2>Change an input.<br><em>Excel does the math.</em></h2>
 <div class="calc-formula reveal"><span class="mono">Report!B3 / verified in Excel</span><pre>=SUM(Sales[Revenue])</pre></div>
 <div class="calc-result"><div id="before"><span>Original total</span><strong>$584k</strong></div><b id="calc-arrow">→</b><div id="after"><span>After +$10k to one order</span><strong>$594k</strong></div></div>
 <div class="preserve reveal">Formulas <i>·</i> Queries <i>·</i> Charts <i>·</i> Workbook features</div>`,
`<div class="kicker">Choose your way in</div>
 <h2 class="closing">Two equal ways in.<br><em>One real Excel engine.</em></h2>
 <div class="closing-interfaces"><div class="closing-card reveal"><strong>MCP</strong><span>Conversational assistants</span></div><div class="closing-card reveal"><strong>excelcli</strong><span>Can use fewer tokens for coding agents</span></div></div>
 <div class="cta reveal">excelmcpserver.dev</div>
 <div class="requirements">Windows + desktop Excel required</div>`
];
const cues = [];
const tracks = [];
const schedule = [];
for (let scene=0;scene<starts.length;scene++) {
  const id = String(scene+1).padStart(2,'0');
  const voiceKey = voiceKeys[scene];
  const phrases = request.lines.filter(x => x.id.startsWith(voiceKey));
  const voices = phrases.map(p => {
    const v=audio.voices.find(v=>v.id===p.id);
    if(!v || !fs.existsSync(v.path)) throw new Error(`Missing voice ${p.id}`);
    return v;
  });
  const total = voices.reduce((a,v)=>a+v.duration_s,0);
  const available = ends[scene]-starts[scene];
  const gap = .1;
  if(total+gap*(voices.length-1) > available-.3) throw new Error(`Voice ${id} exceeds scene: ${total}/${available}`);
  let t = starts[scene] + Math.max(.15,(available-total-gap*(voices.length-1))*.38);
  phrases.forEach((p,i)=>{
    const v=voices[i], end=t+v.duration_s;
    const caption=p.text.replaceAll('excel C L I','excelcli').replaceAll('Excel M C P','Excel MCP').replaceAll('Microsoft three sixty five','Microsoft 365').replaceAll('open pie excel','openpyxl');
    cues.push({id:p.id,text:caption,start:+t.toFixed(3),end:+end.toFixed(3)});
    tracks.push(`<audio id="voice-${p.id}" class="clip" src="${v.path}" data-start="${t.toFixed(3)}" data-duration="${v.duration_s}" data-track-index="10"></audio>`);
    t=end+gap;
  });
  schedule.push({id,voiceKey,start:starts[scene],end:ends[scene],title:titles[scene],voiceDuration:total});
}
const clock=t=>new Date(Math.round(t*1000)).toISOString().slice(11,23);
fs.writeFileSync('excel-mcp-intro.vtt','WEBVTT\n\n'+cues.map((c,i)=>`${i+1}\n${clock(c.start)} --> ${clock(c.end)}\n${c.text}\n`).join('\n'));
fs.writeFileSync('captions.json',JSON.stringify(cues,null,2)+'\n');
fs.writeFileSync('schedule.json',JSON.stringify(schedule,null,2)+'\n');
fs.writeFileSync('SCRIPT.md',`# Excel MCP introduction\n\n**Voice:** Local Kokoro / af_heart (warm female English)\n\n**Voice direction:** Warm, articulate, confident. No music.\n\n**Duration:** Exactly ${duration} seconds. Captions align to individually synthesized phrases, not guessed word timestamps.\n\n`+schedule.map((s,i)=>`## Line ${i+1} — ${s.title} (Frame ${i+1})\n\n**Time:** ${s.start}–${s.end}s\n\n    ${request.lines.filter(p=>p.id.startsWith(s.voiceKey)).map(p=>p.text).join(' ')}\n`).join('\n'));
fs.writeFileSync('STORYBOARD.md',`---\nformat: 1920x1080\nduration: ${duration}s\nmessage: "Understand what Excel MCP does, how it differs, and whether you need it."\narc: Introduce → Compare → Choose → Request → Prepare → Model → Visualize → Extend → Calculate → Start\naudience: people who have never heard of Excel MCP Server\nmode: autonomous\n---\n\n`+schedule.map((s,i)=>`## Frame ${i+1} — ${s.title}\n\n- status: animated\n- src: index.html\n- duration: ${s.end-s.start}s\n- poster: ${s.start+2.5}s\n- transition_in: ${[1,4,6,10].includes(i)?'zoom-through':'cut-the-curve left'}\n- scene: ${titles[i]}\n- voiceover: ${request.lines.filter(p=>p.id.startsWith(s.voiceKey)).map(p=>p.text).join(' ')}\n\nThe scene uses live Excel evidence where applicable; VBA is explicitly illustrative and Python is existing repository evidence.\n`).join('\n'));
const css=fs.readFileSync('style.css','utf8');
const captions=cues.map(c=>`<div id="caption-${c.id}" class="caption" data-layout-allow-caption-zone>${esc(c.text)}</div>`).join('\n');
const revealData=schedule.map(s=>({id:s.id,start:s.start}));
const html=`<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=1920,height=1080">
<title>Excel MCP — Real Excel</title><script src="assets/vendor/gsap.min.js"></script><style>${css}</style></head>
<body><main id="root" data-composition-id="main" data-width="1920" data-height="1080" data-start="0" data-duration="${duration}">
${scenes.map((content,i)=>`<section id="s${String(i+1).padStart(2,'0')}" class="scene ${i===5?'dark':''}" data-layout-allow-overflow>${content}<div class="scene-number">${String(i+1).padStart(2,'0')} / ${scenes.length}</div></section>`).join('\n')}
<div class="caption-overlay">${captions}</div>
<div id="progress" data-layout-ignore></div>
${tracks.join('\n')}
</main><script>
window.__timelines=window.__timelines||{};
const tl=gsap.timeline({paused:true});
window.__timelines["main"]=tl;
tl.fromTo("#progress",{scaleX:0},{scaleX:1,duration:${duration},ease:"none"},0);
const sceneData=${JSON.stringify(revealData)};
sceneData.forEach(({id,start})=>{
  tl.fromTo("#s"+id+" .reveal",{autoAlpha:0,x:36},{autoAlpha:1,x:0,duration:.7,stagger:.14,ease:"power3.out"},start+.4);
});
tl.fromTo("#hero-title",{autoAlpha:0,y:40},{autoAlpha:1,y:0,duration:.8,ease:"power3.out"},.1);
tl.fromTo(".fit-item",{autoAlpha:0,x:35},{autoAlpha:1,x:0,duration:.55,stagger:1.5,ease:"power3.out"},28.5);
tl.fromTo(".step",{autoAlpha:0,x:35},{autoAlpha:1,x:0,duration:.55,stagger:1.8,ease:"power3.out"},49.5);
tl.fromTo(".stat",{autoAlpha:0,y:35},{autoAlpha:1,y:0,duration:.65,stagger:.55,ease:"power3.out"},65.4);
tl.fromTo("#before",{autoAlpha:0},{autoAlpha:1,duration:.35},104);
tl.fromTo("#calc-arrow",{autoAlpha:0,x:20},{autoAlpha:1,x:0,duration:.35},105.2);
tl.fromTo("#after",{autoAlpha:0,x:40},{autoAlpha:1,x:0,duration:.7,ease:"power3.out"},105.7);
${cues.map(c=>`tl.set("#caption-${c.id}",{autoAlpha:1},${c.start});tl.set("#caption-${c.id}",{autoAlpha:0},${c.end});`).join('\n')}
// <seams:auto>
// </seams:auto>
</script></body></html>`;
fs.writeFileSync('index.html',html);
const ledger={fps:30,seams:starts.slice(1).map((cut,i)=>{
  const z=[0,3,5,9].includes(i);
  return {id:`${i+1}-to-${i+2}`,cut,technique:z?'zoom-through':'cut-the-curve LEFT',exit:{selector:`#s${String(i+1).padStart(2,'0')}`,axis:z?'z':'x',dir:z?1:-1,dur:.55},entry:{selector:`#s${String(i+2).padStart(2,'0')}`,axis:z?'z':'x',dir:z?1:-1,dur:.8,travel:12},blur:18};
})};
fs.writeFileSync('ledger.json',JSON.stringify(ledger,null,2)+'\n');
console.log(`Built ${duration}s composition, ${scenes.length} scenes, ${cues.length} precise phrase captions.`);
