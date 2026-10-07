// Cuts, crops and speeds up the raw run recording into the clips used by the composition.
// Usage: node prepare-footage.mjs <path-to-recording.mp4>
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
const config = JSON.parse(fs.readFileSync(path.join(root, 'footage.json'), 'utf8'));
const recording = process.argv[2];
if (!recording || !fs.existsSync(recording)) {
  console.error(`Pass the raw recording (${config.recording}) as the first argument.`);
  process.exit(1);
}
const outDir = path.join(root, 'assets', 'footage');
fs.mkdirSync(outDir, { recursive: true });

for (const clip of config.clips) {
  const speed = (clip.srcEnd - clip.srcStart) / clip.duration;
  const out = path.join(outDir, `${clip.id}.mp4`);
  const filter = `crop=${config.crops[clip.crop]},setpts=(PTS-STARTPTS)/${speed.toFixed(6)},fps=30`;
  const args = ['-v', 'error', '-y', '-ss', String(clip.srcStart), '-t', String(clip.srcEnd - clip.srcStart), '-i', recording,
    '-vf', filter, '-t', String(clip.duration), '-an', '-c:v', 'libx264', '-preset', 'slow', '-crf', '20',
    '-pix_fmt', 'yuv420p', '-movflags', '+faststart', out];
  console.log(`${clip.id}: ${clip.srcStart}-${clip.srcEnd}s at ${speed.toFixed(2)}x -> ${clip.duration}s`);
  const result = spawnSync('ffmpeg', args, { stdio: 'inherit' });
  if (result.status !== 0) process.exit(result.status ?? 1);
}
