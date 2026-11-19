# Agentic World Bank briefing video

Follow the [repository rules](../../AGENTS.md). Invoke `/hyperframes` before
changing the composition, then load the skills it selects. Framework rules and
media treatments belong to those skills, not copied instructions here.

## What this video must stay true to

Slides are exports from the deck produced by the recorded agent run. Screen
footage comes from that run when available. If it is unavailable, use abstract,
clearly labeled illustrative substitutes; do not imitate real app screens or
present substitutes as footage. Do not invent results or numbers.

- Label every substitute in-frame and in its card caption. Hide elapsed-time
  and speed labels when showing substitutes; those labels describe real
  recorded footage only.
- When using recorded footage, show its speed and real time elapsed since the
  prompt was sent. Label skipped stretches with what was skipped.
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
git-ignored; the original clips are about 41 MB. The five small, clearly labeled
illustrative substitutes in `assets/substitutes/` are tracked so a clean
checkout can build and check the composition. Generate them with
`npm run footage:substitutes`. To use the original recording instead, run
`npm run footage -- <path-to-recording.mp4>`; after all clips are prepared,
that command sets `footage.json` to recorded mode. Keep substitutes abstract
and clearly labeled; never make them look like captured screens or invent
results.

The published video is https://youtu.be/_z-twdXG2fA. Regenerate the clips from
the recording:

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

After composition or source edits, run `npm run build` and then `npm run check`.
For a HyperFrames update, use the documented project-upgrade flow and keep the
new pin only if `npm run check` passes. Fix other errors and review warnings
before handoff. PowerShell drops the `--` in `npm run dev -- --background`,
so call the pinned CLI directly:

```powershell
npx --yes hyperframes@0.9.70 preview --background
npx --yes hyperframes@0.9.70 preview --status
npx --yes hyperframes@0.9.70 preview --stop
```

## Render and publication

`npm run render` writes `renders/agentic-world-bank-briefing-4k.mp4` at
`--quality=high --crf=18 --resolution=landscape-4k`. Use `npm run render:draft`
only for previews. Verify the output's dimensions and encoding before upload.
YouTube upload and website embedding need the owner's approval; follow the
publication steps in [the introduction video guide](../excel-mcp-intro/AGENTS.md#youtube-publication).
