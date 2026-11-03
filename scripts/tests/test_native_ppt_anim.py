"""Regression tests for the native PPTX animation plan engine.

Run from the skill directory:
    python -m unittest discover -s scripts/tests -v
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from lxml import etree

TESTS_DIR = Path(__file__).resolve().parent
ENGINE = TESTS_DIR.parent / "native_ppt_anim.py"
FIXTURE = TESTS_DIR / "fixtures" / "base.pptx"
PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
SLIDE_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
)
ENGINE_DIR = TESTS_DIR.parent
if str(ENGINE_DIR) not in sys.path:
    sys.path.insert(0, str(ENGINE_DIR))
import native_ppt_anim  # noqa: E402


def qn(name: str) -> str:
    return f"{{{PML_NS}}}{name}"


def read_parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name) for name in package.namelist()}


def write_parts(path: Path, parts: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        for name, data in parts.items():
            package.writestr(name, data)


def rewrite_slide(parts: dict[str, bytes], part: str, mutate) -> None:
    root = etree.fromstring(parts[part])
    mutate(root)
    parts[part] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def make_two_slide_deck(
    source: Path,
    target: Path,
    *,
    forced_names: bool = False,
) -> Path:
    """Clone the single fixture slide into a second slide."""
    parts = read_parts(source)
    parts["ppt/slides/slide2.xml"] = parts["ppt/slides/slide1.xml"]
    rels_part = "ppt/slides/_rels/slide1.xml.rels"
    if rels_part in parts:
        parts["ppt/slides/_rels/slide2.xml.rels"] = parts[rels_part]

    presentation = etree.fromstring(parts["ppt/presentation.xml"])
    slide_ids = presentation.find(f".//{qn('sldIdLst')}")
    assert slide_ids is not None
    existing_ids = [int(element.get("id", "0")) for element in slide_ids]
    existing_rids = [
        int(element.get(f"{{{R_NS}}}id", "rId0").lstrip("rId"))
        for element in slide_ids
    ]
    new_slide_id = max(existing_ids) + 1
    new_rid = max(existing_rids) + 1
    element = etree.SubElement(slide_ids, qn("sldId"))
    element.set("id", str(new_slide_id))
    element.set(f"{{{R_NS}}}id", f"rId{new_rid}")
    parts["ppt/presentation.xml"] = etree.tostring(
        presentation, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    rels = etree.fromstring(parts["ppt/_rels/presentation.xml.rels"])
    etree.SubElement(
        rels,
        f"{{{PKG_REL_NS}}}Relationship",
        Id=f"rId{new_rid}",
        Type=f"{R_NS}/slide",
        Target="slides/slide2.xml",
    )
    parts["ppt/_rels/presentation.xml.rels"] = etree.tostring(
        rels, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    content_types = parts["[Content_Types].xml"].decode("utf-8")
    content_types = content_types.replace(
        "</Types>",
        f'<Override PartName="/ppt/slides/slide2.xml" ContentType="{SLIDE_CONTENT_TYPE}"/></Types>',
    )
    parts["[Content_Types].xml"] = content_types.encode("utf-8")

    if forced_names:
        def force(root: etree._Element) -> None:
            for prop in root.iter(qn("cNvPr")):
                if prop.get("name") == "hero":
                    prop.set("name", "!!hero")
                if prop.get("name") == "star":
                    prop.set("name", "!!star")

        rewrite_slide(parts, "ppt/slides/slide1.xml", force)
        rewrite_slide(parts, "ppt/slides/slide2.xml", force)
    write_parts(target, parts)
    return target


def inject_audio_timing(parts: dict[str, bytes], part: str, spid: int = 4) -> None:
    def mutate(root: etree._Element) -> None:
        timing = etree.fromstring(
            f'<p:timing xmlns:p="{PML_NS}"><p:tnLst><p:par>'
            f'<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
            f'<p:childTnLst><p:audio isNarration="1"><p:cMediaNode vol="80000">'
            f'<p:cTn id="2" dur="indefinite" fill="hold"/>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
            f'</p:cMediaNode></p:audio></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
        )
        root.append(timing)

    rewrite_slide(parts, part, mutate)


def run_engine(*args: object) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(ENGINE), *[str(arg) for arg in args]],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


def slide_xml(path: Path) -> bytes:
    with zipfile.ZipFile(path) as package:
        return package.read("ppt/slides/slide1.xml")


def timing_root(path: Path) -> etree._Element:
    root = etree.fromstring(slide_xml(path))
    timing = root.find(qn("timing"))
    assert timing is not None, "slide has no timing tree"
    return timing


class NativeAnimationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="pptx-anim-test-")
        self.dir = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def plan(self, name: str, spec: dict) -> Path:
        path = self.dir / f"{name}.json"
        path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
        return path

    def apply(self, spec: dict, name: str) -> tuple[Path, dict]:
        plan_path = self.plan(name, spec)
        output = self.dir / f"{name}.pptx"
        report = self.dir / f"{name}-report.json"
        result = run_engine(
            "apply", FIXTURE, plan_path, "-o", output, "--report", report
        )
        self.assertEqual(result.returncode, 0, result.stderr or result.stdout)
        return output, json.loads(report.read_text(encoding="utf-8"))

    def validate(self, spec: dict, pptx: Path) -> None:
        plan_path = self.plan("validate", spec)
        result = run_engine("validate", pptx, "--spec", plan_path)
        self.assertEqual(result.returncode, 0, result.stderr or result.stdout)

    def assert_apply_fails(self, spec: dict, message: str) -> None:
        plan_path = self.plan("negative", spec)
        output = self.dir / "negative.pptx"
        output.unlink(missing_ok=True)
        result = run_engine("apply", FIXTURE, plan_path, "-o", output)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn(message, result.stderr)
        self.assertFalse(output.exists(), "failed apply must not publish an output file")

    def test_legacy_plan_unchanged(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [
                    {
                        "target": "hero",
                        "effect": "custom_motion",
                        "path_px": "M -240 80 C -160 -120 -60 -100 0 0 E",
                        "trigger": "on-click",
                        "duration": 1.0,
                        "accelerate": 0.1,
                        "bounce_end": 0.18,
                    },
                    {
                        "target": "hero",
                        "effect": "emphasis_spin",
                        "trigger": "with-previous",
                        "duration": 1.0,
                        "accelerate": 0,
                        "decelerate": 0,
                    },
                ],
            }],
        }
        output, report = self.apply(spec, "legacy")
        self.assertEqual(report["slides_changed"][0]["animation_rows"], 2)
        self.validate(spec, output)

    def test_canvas_px_equivalence(self) -> None:
        canvas_spec = {
            "version": 1,
            "canvas_px": [1920, 1080],
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "hero",
                    "effect": "custom_motion",
                    "path_px": "M 0 0 L 300 0 E",
                    "trigger": "on-click",
                    "duration": 1.0,
                }],
            }],
        }
        default_spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "hero",
                    "effect": "custom_motion",
                    "path_px": "M 0 0 L 200 0 E",
                    "trigger": "on-click",
                    "duration": 1.0,
                }],
            }],
        }
        canvas_output, _ = self.apply(canvas_spec, "canvas")
        default_output, _ = self.apply(default_spec, "default")
        canvas_path = re.search(
            r'<p:animMotion[^>]*path="([^"]*)"',
            slide_xml(canvas_output).decode("utf-8"),
        )
        default_path = re.search(
            r'<p:animMotion[^>]*path="([^"]*)"',
            slide_xml(default_output).decode("utf-8"),
        )
        self.assertIsNotNone(canvas_path)
        self.assertIsNotNone(default_path)
        self.assertEqual(canvas_path.group(1), default_path.group(1))
        self.assertEqual(canvas_path.group(1), "M 0 0 L 0.15625 0 E")

    def test_canvas_px_aspect_mismatch(self) -> None:
        spec = {
            "version": 1,
            "canvas_px": [1920, 1200],
            "slides": [{
                "slide": 1,
                "animations": [{"target": "hero", "effect": "entrance_fade"}],
            }],
        }
        self.assert_apply_fails(spec, "aspect ratio")

    def test_from_offset_position_animation(self) -> None:
        spec = {
            "version": 1,
            "canvas_px": [1920, 1080],
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "hero",
                    "effect": "entrance_fade",
                    "trigger": "after-previous",
                    "delay": 0.4,
                    "duration": 0.6,
                    "from_offset_px": [0, 24],
                }],
            }],
        }
        output, report = self.apply(spec, "from-offset")
        self.assertEqual(report["slides_changed"][0]["requested_rows"], 1)
        self.assertEqual(report["slides_changed"][0]["expanded_rows"], 1)
        timing = timing_root(output)
        animations = list(timing.iter(qn("anim")))
        self.assertEqual(len(animations), 1)
        values = animations[0].findall(f"{qn('tavLst')}/{qn('tav')}")
        start = values[0].find(f"{qn('val')}/{qn('strVal')}").get("val")
        end = values[1].find(f"{qn('val')}/{qn('strVal')}").get("val")
        self.assertEqual(start, "#ppt_y+0.022222222222")
        self.assertEqual(end, "#ppt_y")
        self.assertEqual(
            animations[0].find(f"{qn('cBhvr')}/{qn('cTn')}").get("decel"),
            "85000",
        )
        self.validate(spec, output)

    def test_from_offset_rejects_non_entrance(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "hero",
                    "effect": "emphasis_spin",
                    "from_offset_px": [0, 24],
                }],
            }],
        }
        self.assert_apply_fails(spec, "requires an entrance effect")

    def test_from_offset_rejects_zero(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "hero",
                    "effect": "entrance_fade",
                    "from_offset_px": [0, 0],
                }],
            }],
        }
        self.assert_apply_fails(spec, "cannot be [0, 0]")

    def test_targets_stagger_matches_manual_rows(self) -> None:
        sugar = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "targets": ["step-*"],
                    "stagger": 0.14,
                    "effect": "entrance_fade",
                    "trigger": "after-previous",
                    "delay": 0.2,
                    "duration": 0.5,
                }],
            }],
        }
        manual = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [
                    {
                        "target": "step-00",
                        "effect": "entrance_fade",
                        "trigger": "after-previous",
                        "delay": 0.2,
                        "duration": 0.5,
                    },
                    {
                        "target": "step-01",
                        "effect": "entrance_fade",
                        "trigger": "with-previous",
                        "delay": 0.34,
                        "duration": 0.5,
                    },
                    {
                        "target": "step-02",
                        "effect": "entrance_fade",
                        "trigger": "with-previous",
                        "delay": 0.48,
                        "duration": 0.5,
                    },
                    {
                        "target": "step-03",
                        "effect": "entrance_fade",
                        "trigger": "with-previous",
                        "delay": 0.62,
                        "duration": 0.5,
                    },
                ],
            }],
        }
        sugar_output, _ = self.apply(sugar, "targets")
        manual_output, _ = self.apply(manual, "targets-manual")
        self.assertEqual(slide_xml(sugar_output), slide_xml(manual_output))
        self.validate(sugar, sugar_output)

    def test_targets_errors(self) -> None:
        base = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "targets": ["missing-*"],
                    "effect": "entrance_fade",
                }],
            }],
        }
        self.assert_apply_fails(base, "matched no shape")
        base["slides"][0]["animations"][0]["target"] = "hero"
        self.assert_apply_fails(base, "mutually exclusive")
        base["slides"][0]["animations"][0].pop("target")
        base["slides"][0]["animations"][0]["targets"] = ["step-*", "step-00"]
        self.assert_apply_fails(base, "same shape more than once")
        base["slides"][0]["animations"][0]["targets"] = ["step-*"]
        base["slides"][0]["animations"][0]["stagger"] = -0.1
        self.assert_apply_fails(base, "non-negative")

    def test_natural_sort_key(self) -> None:
        sys.path.insert(0, str(ENGINE.parent))
        try:
            import native_ppt_anim
        finally:
            sys.path.pop(0)
        self.assertLess(
            native_ppt_anim.natural_sort_key("step-2"),
            native_ppt_anim.natural_sort_key("step-10"),
        )
        self.assertLess(
            native_ppt_anim.natural_sort_key("step-02"),
            native_ppt_anim.natural_sort_key("step-2a"),
        )

    def test_repeat_until_next_click(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "star",
                    "effect": "emphasis_spin",
                    "trigger": "on-click",
                    "duration": 2.0,
                    "repeat": "until-next-click",
                }],
            }],
        }
        output, _ = self.apply(spec, "repeat-next")
        timing = timing_root(output)
        rows = [
            row for row in timing.iter(qn("cTn"))
            if row.get("presetClass") == "emph"
        ]
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].get("repeatCount"), "indefinite")
        self.assertIsNotNone(rows[0].find(f"{qn('endCondLst')}"))
        self.assertEqual(
            rows[0].find(f"{qn('endCondLst')}/{qn('cond')}").get("evt"),
            "onNext",
        )
        self.validate(spec, output)

    def test_repeat_until_end_of_slide(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "star",
                    "effect": "emphasis_spin",
                    "trigger": "on-click",
                    "duration": 2.0,
                    "repeat": "until-end-of-slide",
                }],
            }],
        }
        output, _ = self.apply(spec, "repeat-end")
        timing = timing_root(output)
        rows = [
            row for row in timing.iter(qn("cTn"))
            if row.get("presetClass") == "emph"
        ]
        self.assertEqual(rows[0].get("repeatCount"), "indefinite")
        self.assertIsNone(rows[0].find(f"{qn('endCondLst')}"))
        self.validate(spec, output)

    def test_repeat_conflict(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "star",
                    "effect": "emphasis_spin",
                    "repeat": "until-next-click",
                    "repeat_count": 2,
                }],
            }],
        }
        self.assert_apply_fails(spec, "mutually exclusive")

    def test_after_previous_after_indefinite(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [
                    {
                        "target": "star",
                        "effect": "emphasis_spin",
                        "trigger": "on-click",
                        "repeat": "until-next-click",
                    },
                    {
                        "target": "hero",
                        "effect": "entrance_fade",
                        "trigger": "after-previous",
                    },
                ],
            }],
        }
        self.assert_apply_fails(spec, "after-previous cannot follow")

    def test_iterate_letter_and_word(self) -> None:
        for unit, expected in (("letter", "lt"), ("word", "wd")):
            spec = {
                "version": 1,
                "slides": [{
                    "slide": 1,
                    "animations": [{
                        "target": "quote",
                        "effect": "entrance_fade",
                        "trigger": "on-click",
                        "iterate": {"by": unit, "delay": 0.04},
                    }],
                }],
            }
            output, _ = self.apply(spec, f"iterate-{unit}")
            timing = timing_root(output)
            iterations = list(timing.iter(qn("iterate")))
            self.assertEqual(len(iterations), 1)
            self.assertEqual(iterations[0].get("type"), expected)
            self.assertEqual(
                iterations[0].find(qn("tmAbs")).get("val"), "40"
            )
            self.validate(spec, output)

    def test_iterate_default_delay_uses_tm_pct(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "quote",
                    "effect": "entrance_fade",
                    "iterate": {"by": "letter"},
                }],
            }],
        }
        output, _ = self.apply(spec, "iterate-default")
        iterate = list(timing_root(output).iter(qn("iterate")))[0]
        self.assertEqual(iterate.find(qn("tmPct")).get("val"), "10000")
        self.validate(spec, output)

    def test_iterate_requires_text(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "hero",
                    "effect": "entrance_fade",
                    "iterate": {"by": "letter"},
                }],
            }],
        }
        self.assert_apply_fails(spec, "text-bearing")

    def test_build_paragraph(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "body",
                    "effect": "entrance_fade",
                    "trigger": "on-click",
                    "build": "paragraph",
                }],
            }],
        }
        output, _ = self.apply(spec, "build")
        timing = timing_root(output)
        builds = list(timing.iter(qn("bldP")))
        self.assertEqual(len(builds), 1)
        self.assertEqual(builds[0].get("build"), "p")
        paragraphs = []
        for pRg in timing.iter(qn("pRg")):
            start = pRg.get("st")
            if not paragraphs or paragraphs[-1] != start:
                paragraphs.append(start)
        self.assertEqual(paragraphs, ["0", "1"])
        rows = [
            row for row in timing.iter(qn("cTn"))
            if row.get("presetClass") == "entr"
        ]
        self.assertEqual(len(rows), 2)
        self.validate(spec, output)

    def test_build_rejects_non_text(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "hero",
                    "effect": "entrance_fade",
                    "build": "paragraph",
                }],
            }],
        }
        self.assert_apply_fails(spec, "text-bearing")

    def test_build_and_iterate_conflict(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "quote",
                    "effect": "entrance_fade",
                    "build": "paragraph",
                    "iterate": {"by": "letter"},
                }],
            }],
        }
        self.assert_apply_fails(spec, "mutually exclusive")

    def test_repeat_count_still_works(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "star",
                    "effect": "emphasis_spin",
                    "trigger": "on-click",
                    "duration": 1.0,
                    "auto_reverse": True,
                    "repeat_count": 2,
                }],
            }],
        }
        output, _ = self.apply(spec, "repeat-count")
        result = run_engine("timeline", output, "--slide", 1)
        self.assertEqual(result.returncode, 0, result.stderr)
        row = json.loads(result.stdout)["slides"][0]["groups"][0]["rows"][0]
        self.assertEqual(row["start_s"], 0.0)
        self.assertEqual(row["end_s"], 4.0)
        self.assertEqual(row["repeat"], "count:2")
        self.validate(spec, output)

    def test_repeat_duration_still_works(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "star",
                    "effect": "emphasis_spin",
                    "trigger": "on-click",
                    "duration": 1.0,
                    "repeat_duration": 3.0,
                }],
            }],
        }
        output, _ = self.apply(spec, "repeat-duration")
        result = run_engine("timeline", output, "--slide", 1)
        self.assertEqual(result.returncode, 0, result.stderr)
        row = json.loads(result.stdout)["slides"][0]["groups"][0]["rows"][0]
        self.assertEqual(row["end_s"], 3.0)
        self.assertEqual(row["repeat"], "duration:3s")
        self.validate(spec, output)

    def test_targets_and_from_offset_report(self) -> None:
        spec = {
            "version": 1,
            "canvas_px": [1920, 1080],
            "slides": [{
                "slide": 1,
                "animations": [{
                    "targets": ["step-*"],
                    "stagger": 0.14,
                    "effect": "entrance_fade",
                    "trigger": "after-previous",
                    "delay": 0.4,
                    "duration": 0.5,
                    "from_offset_px": [0, 26],
                }],
            }],
        }
        output, report = self.apply(spec, "targets-offset")
        slide = report["slides_changed"][0]
        self.assertEqual(slide["requested_rows"], 1)
        self.assertEqual(slide["expanded_rows"], 4)
        self.assertEqual(slide["expansions"][0]["sugar"], "targets+from_offset_px")
        self.assertEqual(slide["expansions"][0]["rows"], 4)
        self.validate(spec, output)

    def test_timeline_schedule(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [
                    {
                        "target": "hero",
                        "effect": "entrance_fade",
                        "trigger": "on-click",
                        "duration": 0.6,
                    },
                    {
                        "target": "step-00",
                        "effect": "entrance_fade",
                        "trigger": "with-previous",
                        "delay": 0.8,
                        "duration": 0.5,
                    },
                ],
            }],
        }
        output, _ = self.apply(spec, "timeline")
        result = run_engine("timeline", output, "--slide", 1)
        self.assertEqual(result.returncode, 0, result.stderr)
        timeline = json.loads(result.stdout)
        rows = timeline["slides"][0]["groups"][0]["rows"]
        self.assertEqual(rows[0]["start_s"], 0.0)
        self.assertEqual(rows[0]["end_s"], 0.6)
        self.assertEqual(rows[1]["start_s"], 0.8)
        self.assertEqual(rows[1]["end_s"], 1.3)


class NativeTransitionAndSafetyTests(unittest.TestCase):
    """Transition, Morph, kiosk, replace_existing and atomic-publish coverage."""

    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="pptx-anim-safety-")
        self.dir = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def plan(self, name: str, spec: dict) -> Path:
        path = self.dir / f"{name}.json"
        path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
        return path

    def apply_plan(
        self, source: Path, spec: dict, name: str
    ) -> tuple[Path, dict, subprocess.CompletedProcess[str]]:
        plan_path = self.plan(name, spec)
        output = self.dir / f"{name}.pptx"
        report = self.dir / f"{name}-report.json"
        result = run_engine(
            "apply", source, plan_path, "-o", output, "--report", report
        )
        self.assertEqual(result.returncode, 0, result.stderr or result.stdout)
        return output, json.loads(report.read_text(encoding="utf-8")), result

    def assert_apply_fails(
        self, source: Path, spec: dict, message: str, name: str = "negative"
    ) -> None:
        plan_path = self.plan(name, spec)
        output = self.dir / f"{name}.pptx"
        output.unlink(missing_ok=True)
        result = run_engine("apply", source, plan_path, "-o", output)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn(message, result.stderr)
        self.assertFalse(output.exists(), "failed apply must not publish an output file")

    def test_transition_apply_and_validate(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "transition": {"effect": "fade", "duration": 0.4, "advance_after": 1.5},
            }],
        }
        output, report, _ = self.apply_plan(FIXTURE, spec, "transition")
        self.assertEqual(
            report["slides_changed"][0]["transition"]["canonical_effect"], "fade"
        )
        result = run_engine("validate", output, "--spec", self.plan("v", spec))
        self.assertEqual(result.returncode, 0, result.stderr)
        parts = read_parts(output)
        self.assertIn("ppt/presProps.xml", parts)
        self.assertIn("useTimings", parts["ppt/presProps.xml"].decode("utf-8"))

    def test_transition_requires_effect_for_duration(self) -> None:
        spec = {"version": 1, "slides": [{"slide": 1, "transition": {"duration": 0.5}}]}
        self.assert_apply_fails(FIXTURE, spec, "require an explicit effect", "inert-t")

    def test_transition_rejects_advance_after_zero(self) -> None:
        spec = {
            "version": 1,
            "slides": [{"slide": 1, "transition": {"effect": "fade", "advance_after": 0}}],
        }
        self.assert_apply_fails(FIXTURE, spec, "positive", "zero-advance")

    def test_morph_rejected_on_first_slide(self) -> None:
        spec = {"version": 1, "slides": [{"slide": 1, "transition": {"effect": "morph"}}]}
        self.assert_apply_fails(FIXTURE, spec, "preceding source slide", "morph-1")

    def test_morph_requires_shared_keys(self) -> None:
        deck = make_two_slide_deck(FIXTURE, self.dir / "two.pptx")
        spec = {"version": 1, "slides": [{"slide": 2, "transition": {"effect": "morph"}}]}
        self.assert_apply_fails(deck, spec, "no shared !!key", "morph-keys")

    def test_morph_scope_ignores_untouched_slide(self) -> None:
        deck = make_two_slide_deck(FIXTURE, self.dir / "two.pptx")
        parts = read_parts(deck)
        slide2, _ = native_ppt_anim.apply_transition(
            parts["ppt/slides/slide2.xml"], {"effect": "morph"}, "slide 2"
        )
        parts["ppt/slides/slide2.xml"] = slide2
        write_parts(deck, parts)
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{"target": "hero", "effect": "entrance_fade"}],
            }],
        }
        output, _, _ = self.apply_plan(deck, spec, "morph-scope")
        out_parts = read_parts(output)
        transition = native_ppt_anim.read_slide_transition_xml(
            out_parts["ppt/slides/slide2.xml"]
        )
        self.assertEqual(transition.canonical_effect, "morph")

    def test_morph_with_shared_keys_applies(self) -> None:
        deck = make_two_slide_deck(
            FIXTURE, self.dir / "two-forced.pptx", forced_names=True
        )
        spec = {"version": 1, "slides": [{"slide": 2, "transition": {"effect": "morph"}}]}
        output, report, _ = self.apply_plan(deck, spec, "morph-ok")
        names = {pair["name"] for pair in report["morph_pairs"]}
        self.assertEqual(names, {"!!hero", "!!star"})
        result = run_engine("validate", output, "--spec", self.plan("v2", spec))
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_kiosk_applied_and_validated(self) -> None:
        spec = {
            "version": 1,
            "kiosk": True,
            "slides": [{
                "slide": 1,
                "animations": [{"target": "hero", "effect": "entrance_fade"}],
                "transition": {"effect": "fade"},
            }],
        }
        output, report, _ = self.apply_plan(FIXTURE, spec, "kiosk")
        self.assertTrue(report["validation"]["opc_xml_and_relationships"] == "passed")
        result = run_engine("validate", output, "--spec", self.plan("vk", spec))
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue(json.loads(result.stdout)["kiosk"])

    def test_validate_catches_kiosk_mismatch(self) -> None:
        spec = {"version": 1, "kiosk": True, "slides": [{"slide": 1}]}
        result = run_engine("validate", FIXTURE, "--spec", self.plan("kiosk-miss", spec))
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("not a kiosk show", result.stderr)

    def test_replace_existing_clears_timeline(self) -> None:
        animated, _, _ = self.apply_plan(
            FIXTURE,
            {
                "version": 1,
                "slides": [{"slide": 1, "animations": [{"target": "hero", "effect": "entrance_fade"}]}],
            },
            "pre-clear",
        )
        spec = {"version": 1, "slides": [{"slide": 1, "replace_existing": True, "animations": []}]}
        cleared, report, _ = self.apply_plan(animated, spec, "cleared")
        root = etree.fromstring(read_parts(cleared)["ppt/slides/slide1.xml"])
        self.assertIsNone(root.find(qn("timing")))
        self.assertEqual(report["slides_changed"][0]["animation_rows"], 0)

    def test_replace_existing_empty_keeps_transition(self) -> None:
        animated, _, _ = self.apply_plan(
            FIXTURE,
            {
                "version": 1,
                "slides": [{
                    "slide": 1,
                    "animations": [{"target": "hero", "effect": "entrance_fade"}],
                    "transition": {"effect": "fade", "duration": 0.4},
                }],
            },
            "pre-transition",
        )
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "replace_existing": True,
                "animations": [],
                "transition": {"effect": "fade", "duration": 0.4},
            }],
        }
        cleared, _, _ = self.apply_plan(animated, spec, "kept-transition")
        summary = native_ppt_anim.read_slide_transition_xml(
            read_parts(cleared)["ppt/slides/slide1.xml"]
        )
        self.assertEqual(summary.canonical_effect, "fade")
        self.assertEqual(summary.duration_ms, 400)

    def test_replace_existing_preserves_audio(self) -> None:
        deck = self.dir / "media.pptx"
        deck.write_bytes(FIXTURE.read_bytes())
        parts = read_parts(deck)
        inject_audio_timing(parts, "ppt/slides/slide1.xml")
        write_parts(deck, parts)
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "replace_existing": True,
                "animations": [{"target": "hero", "effect": "entrance_fade"}],
            }],
        }
        output, report, _ = self.apply_plan(deck, spec, "media-replace")
        self.assertEqual(report["slides_changed"][0]["preserved_media"], 1)
        root = etree.fromstring(read_parts(output)["ppt/slides/slide1.xml"])
        audio = list(root.iter(qn("audio")))
        self.assertEqual(len(audio), 1)
        ids = [
            int(ctn.get("id"))
            for ctn in root.iter(qn("cTn"))
            if (ctn.get("id") or "").isdigit()
        ]
        self.assertEqual(len(ids), len(set(ids)), "cTn ids must stay unique")

    def test_in_place_failure_preserves_original(self) -> None:
        deck = self.dir / "in-place.pptx"
        deck.write_bytes(FIXTURE.read_bytes())
        before = deck.read_bytes()
        spec = {
            "version": 1,
            "slides": [{"slide": 1, "animations": [{"target": "hero", "effect": "entrance_fade"}]}],
        }
        with mock.patch.object(
            native_ppt_anim,
            "validate_pptx_animation_package",
            side_effect=ValueError("forced post-write failure"),
        ):
            with self.assertRaises(ValueError):
                native_ppt_anim.apply_spec(deck, deck, spec)
        self.assertEqual(deck.read_bytes(), before)
        self.assertEqual(list(self.dir.glob("*.tmp")), [])

    def test_in_place_apply_succeeds(self) -> None:
        deck = self.dir / "in-place-ok.pptx"
        deck.write_bytes(FIXTURE.read_bytes())
        before = deck.read_bytes()
        spec = {
            "version": 1,
            "slides": [{"slide": 1, "animations": [{"target": "hero", "effect": "entrance_fade"}]}],
        }
        report = native_ppt_anim.apply_spec(deck, deck, spec)
        self.assertNotEqual(deck.read_bytes(), before)
        self.assertEqual(report["sha256"], native_ppt_anim.hashlib.sha256(deck.read_bytes()).hexdigest())
        root = etree.fromstring(read_parts(deck)["ppt/slides/slide1.xml"])
        self.assertIsNotNone(root.find(qn("timing")))

    def test_trigger_shape_runs_in_interactive_sequence(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "hero",
                    "effect": "entrance_fade",
                    "trigger": "on-click",
                    "trigger_shape": "star",
                }],
            }],
        }
        output, _, _ = self.apply_plan(FIXTURE, spec, "trigger-shape")
        root = etree.fromstring(read_parts(output)["ppt/slides/slide1.xml"])
        interactive = [
            ctn for ctn in root.iter(qn("cTn"))
            if ctn.get("nodeType") == "interactiveSeq"
        ]
        self.assertEqual(len(interactive), 1)
        timeline = run_engine("timeline", output, "--slide", 1)
        self.assertEqual(timeline.returncode, 0, timeline.stderr)
        group = json.loads(timeline.stdout)["slides"][0]["groups"][0]
        self.assertEqual(group["kind"], "shape")
        self.assertEqual(group["trigger_shape"], "star")

    def test_timeline_indefinite_row_has_null_end(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "star",
                    "effect": "emphasis_spin",
                    "trigger": "on-click",
                    "repeat": "until-end-of-slide",
                }],
            }],
        }
        output, _, _ = self.apply_plan(FIXTURE, spec, "indefinite")
        timeline = run_engine("timeline", output, "--slide", 1)
        row = json.loads(timeline.stdout)["slides"][0]["groups"][0]["rows"][0]
        self.assertIsNone(row["end_s"])

    def test_timeline_slide_out_of_range(self) -> None:
        result = run_engine("timeline", FIXTURE, "--slide", 99)
        self.assertEqual(result.returncode, 2)
        self.assertIn("out of range", result.stderr)

    def test_validate_spec_catches_mutated_duration(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "hero",
                    "effect": "entrance_fade",
                    "duration": 0.5,
                }],
            }],
        }
        output, _, _ = self.apply_plan(FIXTURE, spec, "mutate-duration")
        parts = read_parts(output)
        root = etree.fromstring(parts["ppt/slides/slide1.xml"])
        changed = 0
        for ctn in root.iter(qn("cTn")):
            if ctn.get("dur") == "500":
                ctn.set("dur", "900")
                changed += 1
        self.assertGreater(changed, 0)
        parts["ppt/slides/slide1.xml"] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        write_parts(output, parts)
        result = run_engine("validate", output, "--spec", self.plan("mutated", spec))
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("duration", result.stderr.lower())

    def test_validate_without_spec_rejects_unreadable_rows(self) -> None:
        output, _, _ = self.apply_plan(
            FIXTURE,
            {
                "version": 1,
                "slides": [{"slide": 1, "animations": [{"target": "hero", "effect": "entrance_fade"}]}],
            },
            "broken",
        )
        parts = read_parts(output)
        root = etree.fromstring(parts["ppt/slides/slide1.xml"])
        row = next(
            ctn for ctn in root.iter(qn("cTn")) if ctn.get("presetClass") == "entr"
        )
        row.set("nodeType", "bogus")
        parts["ppt/slides/slide1.xml"] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        write_parts(output, parts)
        result = run_engine("validate", output)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("unsupported object-animation nodeType", result.stderr)

    def test_untouched_parts_byte_identical(self) -> None:
        before = read_parts(FIXTURE)
        output, _, _ = self.apply_plan(
            FIXTURE,
            {
                "version": 1,
                "slides": [{"slide": 1, "animations": [{"target": "hero", "effect": "entrance_fade"}]}],
            },
            "untouched",
        )
        after = read_parts(output)
        changed = {name for name in before if before[name] != after.get(name)}
        self.assertEqual(changed, {"ppt/slides/slide1.xml"})

    def test_ctn_child_order(self) -> None:
        spec = {
            "version": 1,
            "slides": [{
                "slide": 1,
                "animations": [{
                    "target": "quote",
                    "effect": "entrance_fade",
                    "iterate": {"by": "letter", "delay": 0.04},
                    "repeat": "until-next-click",
                }],
            }],
        }
        output, _, _ = self.apply_plan(FIXTURE, spec, "order")
        root = etree.fromstring(read_parts(output)["ppt/slides/slide1.xml"])
        row = next(
            ctn
            for ctn in root.iter(qn("cTn"))
            if ctn.get("presetClass") == "entr" and ctn.find(qn("iterate")) is not None
        )
        tags = [etree.QName(child).localname for child in row]
        for earlier, later in (
            ("stCondLst", "endCondLst"),
            ("endCondLst", "iterate"),
            ("iterate", "childTnLst"),
        ):
            self.assertLess(tags.index(earlier), tags.index(later), tags)

    def test_describe_exposes_sugar_fields(self) -> None:
        result = run_engine("describe", "entrance_fade")
        self.assertEqual(result.returncode, 0, result.stderr)
        described = json.loads(result.stdout)
        for field in (
            "repeat", "iterate", "build", "targets", "stagger",
            "from_offset_px", "offset_decelerate",
        ):
            self.assertIn(field, described["timing"])
        self.assertIn("canvas_px", described)
        motion = json.loads(run_engine("describe", "custom_motion").stdout)
        self.assertIn("canvas_px", motion["path_px"])


if __name__ == "__main__":
    unittest.main()
