# Agentic World Bank briefing video

Follow the [repository rules](../../AGENTS.md). Invoke `/hyperframes` before
changing the composition, then load the skills it selects. Framework rules and
media treatments belong to those skills, not copied instructions here.

## What this video must stay true to

Every screen is real: footage from one recorded, unedited agent run, or a slide
exported from the deck that run produced. Do not add invented screens, results
or numbers.

- Sped-up footage shows its speed; the clock shows real time elapsed since the
  prompt was sent. Skipped stretches are labelled with what was skipped.
- Effort figures (prompt, words, minutes, tool calls, interventions) come from
  `evidence.json`, which was measured from the Copilot CLI session record.
- Data claims match the deck and the World Bank WDI source (CC BY 4.0).
- Excel MCP Server and PowerPoint MCP Server are named and shown equally.

## Sources and generated files

`build.mjs` is the source of truth. It reads `audio_request.json`,
`audio_meta.json`, `footage.json` and `evidence.json`, then writes `index.html`,
`captions.json`, `agentic-world-bank-briefing.vtt`, `schedule.json`,
`SCRIPT.md` and `STORYBOARD.md`. Edit the sources and run `npm run build`; do
not hand-edit the generated files.

Narration text spells out acronyms for the voice (`M C P`, `C L I`, `A I`);
`build.mjs` converts them back for captions. Regenerate voice lines with the
media-use audio script and the local Kokoro `af_heart` voice.

## Footage

The raw OBS recording stays outside the repository. `assets/footage/` is
git-ignored on purpose: the prepared clips are about 41 MB, and committing them
would add that size to every clone. A clean checkout therefore cannot build or
render the video; `npm run build` stops with a missing-footage error. The
published video is https://youtu.be/_z-twdXG2fA. Regenerate the clips from the
recording:

```powershell
npm run footage -- <path-to-recording.mp4>
```

`footage.json` holds each clip's source range, crop and output length. Keep
crops free of window title bars, the taskbar and notifications. Do not use the
recording after 32:40, because another window covers the screen.

## Slides

`assets/slides/Slide1.png` to `Slide7.png` are the deck's own 4K exports. The
opening fan and closing grid use smaller copies in `assets/slides/fan/` and
`assets/slides/grid/`; regenerate them after replacing an export:

```powershell
npm run slides
```

## Checks and preview

Run `npm run build` and then `npm run check` after every change. Fix all errors
and review warnings before handoff. For Studio review use the persistent
preview. PowerShell drops the `--` in `npm run dev -- --background`, so call
the pinned CLI directly:

```powershell
npx --yes hyperframes@0.8.139 preview --background
npx --yes hyperframes@0.8.139 preview --status
npx --yes hyperframes@0.8.139 preview --stop
```

## Render and publication

`npm run render` writes `renders/agentic-world-bank-briefing-4k.mp4` at
`--quality=high --crf=18 --resolution=landscape-4k`. Use `npm run render:draft`
only for previews. Verify the output's dimensions and encoding before upload.
YouTube upload and website embedding need the owner's approval; follow the
publication steps in [the introduction video guide](../excel-mcp-intro/AGENTS.md#youtube-publication).
