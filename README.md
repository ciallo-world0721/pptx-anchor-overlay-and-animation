# PPTX Native Animation

Native object animation and Morph transitions for `.pptx` files — written as real
OOXML `p:timing` / `p:transition`, not embedded video, GIFs, or page screenshots.
Everything stays editable in PowerPoint after the fact.

A skill for agentic coding assistants (Hermes Agent / Claude-style skill layout):
`SKILL.md` is the operating manual, `references/` holds the specifications,
`scripts/` holds the engine.

## What it actually does

- **State contract first.** You describe start frame, end frame, object identity,
  and where the audience should dwell — before any XML is touched.
- **Two engines, chosen by intent.** Page-to-page multi-object moves, scales and
  rotations → Morph double-state pages. In-slide click sequences, emphasis, motion
  paths, staggered batches → the plan engine that writes `p:timing` from a JSON
  plan.
- **203 native PowerPoint object effects** (53 entrance / 33 emphasis / 64 motion
  path / 53 exit) in a queryable preset table with real behavior nodes.
- **Six-layer verification, no skipping.** Package integrity → semantic read-back
  → static render → PowerPoint COM audit (an oracle that shares no code with the
  engine) → live slideshow.
- **Strict, loud failures.** Unknown effect names, unknown plan fields, ambiguous
  shape targets and illegal trigger/duration combinations raise instead of being
  silently replaced. Writes are atomic: a failed plan never lands on disk.

## Repository layout

```
SKILL.md                        # when to use, four invariants, design layer, verification
references/
  design-states.md              # state contracts, Morph identity, final geometry, OOXML
  timing-plan-spec.md           # full plan schema, scheduling formulas, path grammar
  verification.md               # six-layer checks, slideshow review, COM audit
  com-timing.md                 # COM process safety, AddEffect, behaviors
  layout-and-integrity.md       # anchoring, overlays, package and relationship checks
  provenance-licensing.md       # preset provenance and reuse terms
scripts/
  native_ppt_anim.py            # plan engine: inspect / list-effects / apply / validate
  pptx_com_audit.py             # PowerPoint read-back
  requirements.txt              # lxml (+ pywin32 on Windows)
  vendor/                       # ppt-master preset table (MIT, see vendor/LICENSE)
```

## Setup

```bash
uv venv .venv --python 3.11
uv pip install --python .venv/Scripts/python.exe -r scripts/requirements.txt   # Windows
# macOS/Linux: uv pip install --python .venv/bin/python -r scripts/requirements.txt
PY=.venv/Scripts/python.exe      # use $PY for every command below
```

`lxml` is the only hard dependency. `pywin32` is needed for the COM audit
(Office on Windows) and is skipped elsewhere.

## Quick start

```bash
# 1. Read the deck: slide order, shape names/IDs, existing timelines
$PY scripts/native_ppt_anim.py inspect deck.pptx

# 2. Query the preset table instead of guessing names
$PY scripts/native_ppt_anim.py list-effects --category emphasis
$PY scripts/native_ppt_anim.py describe emphasis_spin
$PY scripts/native_ppt_anim.py list-transitions

# 3. Write a plan (atomic replace; a failure leaves the input untouched)
$PY scripts/native_ppt_anim.py apply deck.pptx plan.json -o animated.pptx --report report.json

# 4. Validate the result against the plan, then ask PowerPoint itself
$PY scripts/native_ppt_anim.py validate animated.pptx --spec plan.json
$PY scripts/pptx_com_audit.py animated.pptx --report com-audit.json
```

Minimal plan:

```json
{
  "version": 1,
  "slides": [
    {
      "slide": 1,
      "animations": [
        {
          "target": "hero",
          "effect": "custom_motion",
          "path_px": "M -240 80 C -160 -120 -60 -100 0 0 E",
          "trigger": "on-click",
          "duration": 1.0,
          "accelerate": 0.1,
          "bounce_end": 0.18
        }
      ],
      "transition": { "effect": "fade", "duration": 0.4 }
    }
  ]
}
```

`target` and `trigger_shape` accept a unique shape name or a positive shape ID.
Durations are seconds. Scheduling, repeat, auto-reverse and bounce semantics are
specified in `references/timing-plan-spec.md`.

## Honest limits

No animation audio, no arbitrary mask morphing or frame-by-frame physics, no
COM *writing* wrapper (COM is a manual route and a verification oracle only).
Custom paths support `M`/`L`/`C`/`E` — no `Q`, arcs or implicit relative
commands. When something is out of scope it says so instead of faking success.

Structure passing is not the same as looking right: `validate` green means the
deck is well-formed and semantically matches the plan — it never means the
motion has been visually confirmed.

## License

MIT — see [LICENSE](LICENSE).

`scripts/vendor/` is derived from [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master)
(MIT, Copyright (c) 2025–2026 Hugo He); its license is preserved at
`scripts/vendor/LICENSE-ppt-master.txt`. Preset provenance and the terms of every
reference implementation consulted are recorded in `references/provenance-licensing.md`.
