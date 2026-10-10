// Writes the smaller full-colour slide copies used by the opening fan and closing grid.
// PowerPoint saves some exports with a 256-colour palette; shrinking those without
// converting to full colour first leaves coloured fringes on text.
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const slides = path.join(path.dirname(fileURLToPath(import.meta.url)), 'assets', 'slides');
for (const [folder, width] of [['fan', 1600], ['grid', 880]]) {
  fs.mkdirSync(path.join(slides, folder), { recursive: true });
  for (let n = 1; n <= 7; n++) {
    const args = ['-v', 'error', '-y', '-i', path.join(slides, `Slide${n}.png`),
      '-vf', `format=rgb24,scale=${width}:-1:flags=lanczos`, '-pix_fmt', 'rgb24',
      path.join(slides, folder, `Slide${n}.png`)];
    const result = spawnSync('ffmpeg', args, { stdio: 'inherit' });
    if (result.status !== 0) process.exit(result.status ?? 1);
  }
}
