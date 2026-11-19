import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
const configPath = path.join(root, 'footage.json');
const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
const outputDir = path.join(root, 'assets', 'substitutes');
const fontPath = path.join(process.env.WINDIR ?? 'C:\\Windows', 'Fonts', 'segoeui.ttf')
  .replaceAll('\\', '/')
  .replace(':', '\\:');
const stages = ['Prompt', 'Research', 'Excel', 'PowerPoint'];
const stageForClip = {
  ask: 'Prompt',
  research: 'Research',
  excel: 'Excel',
  'ppt-a': 'PowerPoint',
  'ppt-b': 'PowerPoint',
};

if (!fs.existsSync(fontPath.replace('\\:', ':'))) {
  throw new Error(`Required Windows font was not found: ${fontPath}`);
}

fs.mkdirSync(outputDir, { recursive: true });

for (const clip of config.clips) {
  const activeStage = stageForClip[clip.id];
  if (!activeStage) throw new Error(`No illustrative stage is defined for ${clip.id}`);

  const filters = [
    'drawbox=x=80:y=70:w=1760:h=940:color=0x5FD08F@0.3:t=2',
    'drawbox=x=100:y=175:w=1720:h=3:color=0x5FD08F@0.5:t=fill',
    `drawtext=fontfile='${fontPath}':text='ILLUSTRATIVE SUBSTITUTE':fontcolor=0x5FD08F:fontsize=30:x=100:y=112`,
    `drawtext=fontfile='${fontPath}':text='${activeStage}':fontcolor=0xF4F7F5:fontsize=82:x=(w-text_w)/2:y=285`,
    `drawtext=fontfile='${fontPath}':text='Illustrative substitute':fontcolor=0xB4C4BA:fontsize=34:x=(w-text_w)/2:y=420`,
    `drawtext=fontfile='${fontPath}':text='Not original footage':fontcolor=0xB4C4BA:fontsize=26:x=(w-text_w)/2:y=468`,
    ...stages.map((stage, index) => {
      const x = 120 + index * 440;
      const selected = stage === activeStage;
      const border = selected ? '0x5FD08F' : '0x496251';
      const text = selected ? '0x5FD08F' : '0xB4C4BA';
      return [
        `drawbox=x=${x}:y=600:w=360:h=112:color=${border}@0.85:t=3`,
        `drawtext=fontfile='${fontPath}':text='${stage}':fontcolor=${text}:fontsize=30:x=${x}+(360-text_w)/2:y=640`,
      ].join(',');
    }),
    'drawbox=x=120:y=790:w=1680:h=3:color=0x496251:t=fill',
    "drawbox=x='120+mod(t*260\\,1560)':y=779:w=120:h=24:color=0x5FD08F:t=fill",
  ].join(',');
  const output = path.join(outputDir, `${clip.id}.mp4`);
  const args = [
    '-v', 'error', '-y',
    '-f', 'lavfi', '-i', `color=c=0x0D1117:s=1920x1080:r=30:d=${clip.duration}`,
    '-vf', filters,
    '-an', '-c:v', 'libx264', '-preset', 'veryfast', '-crf', '28',
    '-pix_fmt', 'yuv420p', '-movflags', '+faststart', '-t', String(clip.duration),
    output,
  ];

  console.log(`${clip.id}: ${activeStage} illustrative card (${clip.duration}s)`);
  const result = spawnSync('ffmpeg', args, { stdio: 'inherit' });
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
}

config.sourceMode = 'illustrative-substitutes';
fs.writeFileSync(configPath, `${JSON.stringify(config, null, 2)}\n`);
