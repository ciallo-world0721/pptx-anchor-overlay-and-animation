#!/usr/bin/env python
"""Read back a PPTX timing tree through Microsoft PowerPoint (COM).

This is the independent oracle: it asks PowerPoint's own object model what it
parsed, not our reader. Structural evidence only -- it cannot judge rhythm,
overlap, or smoothness.

Usage:
    python pptx_com_audit.py deck.pptx [--report audit.json]

Exit codes: 0 ok, 2 file missing, 3 PowerPoint already running, 4 COM failure.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import NoReturn


def _fail(message: str, code: int) -> NoReturn:
    print(f"error: {message}", file=sys.stderr)
    raise SystemExit(code)


def audit(path: Path) -> list[dict]:
    import win32com.client as win32

    # PowerPoint is single-instance for automation. Never attach to a session
    # the user is working in: a Quit() there closes their windows.
    try:
        win32.GetActiveObject("PowerPoint.Application")
    except Exception:
        pass
    else:
        _fail("PowerPoint is already running; close it before an automated read-back", 3)

    app = win32.DispatchEx("PowerPoint.Application")
    pres = None
    try:
        pres = app.Presentations.Open(str(path), ReadOnly=True, Untitled=False, WithWindow=False)
        slides = []
        for index, slide in enumerate(pres.Slides, start=1):
            seq = slide.TimeLine.MainSequence
            effects = []
            for i in range(1, seq.Count + 1):
                effects.append(_effect_row(i, seq.Item(i)))
            interactive = []
            for s in range(1, slide.TimeLine.InteractiveSequences.Count + 1):
                iseq = slide.TimeLine.InteractiveSequences.Item(s)
                rows = []
                for i in range(1, iseq.Count + 1):
                    fx = iseq.Item(i)
                    rows.append(
                        {
                            "shape": _shape_name(fx),
                            "effect_type": fx.EffectType,
                            "trigger": fx.Timing.TriggerType,
                            "duration_s": round(fx.Timing.Duration, 4),
                        }
                    )
                interactive.append(rows)
            trans = slide.SlideShowTransition
            slides.append(
                {
                    "slide": index,
                    "shapes": slide.Shapes.Count,
                    "main_sequence": seq.Count,
                    "interactive_sequences": slide.TimeLine.InteractiveSequences.Count,
                    "interactive": interactive,
                    "effects": effects,
                    "transition": {
                        "entry_effect": trans.EntryEffect,
                        "duration_s": round(trans.Duration, 4),
                        "advance_on_click": bool(trans.AdvanceOnClick),
                        "advance_on_time": bool(trans.AdvanceOnTime),
                        "advance_time_s": round(trans.AdvanceTime, 4),
                    },
                }
            )
        return slides
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        try:
            app.Quit()
        except Exception:
            pass


def _safe(getter, default=None):
    try:
        return getter()
    except Exception:
        return default


def _shape_name(effect) -> str:
    try:
        return effect.Shape.Name
    except Exception:
        return "<none>"


def _effect_row(index: int, effect) -> dict:
    timing = effect.Timing
    return {
        "index": index,
        "shape": _shape_name(effect),
        "effect_type": effect.EffectType,
        "trigger": timing.TriggerType,
        "exit": effect.Exit,
        "duration_s": round(timing.Duration, 4),
        "delay_s": round(timing.TriggerDelayTime, 4),
        "paragraph": _safe(lambda: int(effect.Paragraph)),
        "build_by_level": _safe(lambda: int(timing.BuildByLevelEffect)),
        "repeat_count": _safe(lambda: round(float(timing.RepeatCount), 4)),
        "auto_reverse": _safe(lambda: bool(timing.AutoReverse)),
        "rewind": _safe(lambda: bool(timing.Rewind)),
        "behaviors": _safe(lambda: int(effect.Behaviors.Count)),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Read a PPTX timing tree via PowerPoint COM")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--report", type=Path, help="write the JSON report to this path")
    args = parser.parse_args()

    path = args.pptx.resolve()
    if not path.is_file():
        _fail(f"input PPTX does not exist: {path}", 2)

    try:
        slides = audit(path)
    except SystemExit:
        raise
    except Exception as exc:  # COM surface produces very noisy exceptions
        _fail(f"PowerPoint COM read-back failed: {exc}", 4)

    payload = {"path": str(path), "slide_count": len(slides), "slides": slides}
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.report:
        args.report.write_text(text, encoding="utf-8")
        print(f"report: {args.report}")
    else:
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
        print(text)


if __name__ == "__main__":
    main()
