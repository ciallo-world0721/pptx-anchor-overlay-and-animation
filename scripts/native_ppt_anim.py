#!/usr/bin/env python3
"""Apply and verify native PowerPoint object animations and transitions.

This is the stable, file-level interface of the pptx-anchor-overlay-and-animation
skill. It resolves
slide order and shape names, writes native OOXML timing trees, and reads every
requested animation back before publishing the output package.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import math
import os
import posixpath
import re
import sys
import tempfile
import zipfile
from dataclasses import asdict
from fnmatch import fnmatchcase
from functools import lru_cache
from pathlib import Path
from typing import Any, Mapping

from lxml import etree

VENDOR_DIR = Path(__file__).resolve().parent / "vendor"
sys.path.insert(0, str(VENDOR_DIR))

from pptx_animations import (  # noqa: E402
    NATIVE_ANIMATIONS,
    PML_NS,
    create_sequence_timing_xml,
    describe_animation_effect,
    get_available_animations,
    get_available_transitions,
    read_slide_animation_sequence,
    validate_generated_animation_xml,
    validate_pptx_animation_package,
)
from pptx_transitions import (  # noqa: E402
    MC_NS,
    NATIVE_TRANSITIONS,
    AdvanceUpdate,
    EnterUpdate,
    apply_slide_motion_xml,
    describe_transition_effect,
    normalize_transition_effect_request,
    read_slide_transition_xml,
    set_package_kiosk_show,
    set_package_use_timings,
    validate_pptx_transition_package,
)

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
PRESET_CLASSES = {"entr", "emph", "path", "exit"}
CUSTOM_MOTION = "custom_motion"
TOP_FIELDS = {"version", "slides", "kiosk", "canvas_px"}
SLIDE_FIELDS = {"slide", "replace_existing", "animations", "transition"}
ANIMATION_FIELDS = {
    "target", "targets", "stagger", "effect", "duration", "delay", "trigger", "trigger_shape",
    "effect_options", "repeat_count", "repeat_duration", "repeat", "auto_reverse",
    "rewind", "accelerate", "decelerate", "bounce_end", "restart",
    "after_effect", "path", "path_px", "from_offset_px", "offset_decelerate",
    "iterate", "build", "paragraph",
}
TRANSITION_FIELDS = {
    "effect", "duration", "effect_options", "advance_after", "advance_on_click"
}
CORE_TIMING_FIELDS = {
    "repeat_count", "repeat_duration", "auto_reverse", "rewind", "accelerate",
    "decelerate", "bounce_end", "restart", "after_effect",
}


class SpecError(ValueError):
    """An author-facing specification or package consistency error."""


def qn(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


def parse_xml(data: bytes | str) -> etree._Element:
    if isinstance(data, str):
        data = data.encode("utf-8")
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    try:
        return etree.fromstring(data, parser)
    except etree.XMLSyntaxError as exc:
        raise SpecError(f"invalid XML: {exc}") from exc


def serialize_slide(root: etree._Element) -> bytes:
    return etree.tostring(
        root, encoding="UTF-8", xml_declaration=True, standalone=True
    )


def reject_unknown(value: Mapping[str, Any], allowed: set[str], where: str) -> None:
    unknown = sorted(set(value) - allowed)
    if unknown:
        raise SpecError(f"{where} has unknown field(s): {', '.join(unknown)}")


def require_mapping(value: Any, where: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise SpecError(f"{where} must be an object")
    return value


def positive_int(value: Any, where: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise SpecError(f"{where} must be a positive integer")
    return value


def seconds_to_ms(value: Any, where: str, *, allow_zero: bool = True) -> int:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise SpecError(f"{where} must be a number in seconds")
    if not math.isfinite(value) or value < 0 or (value == 0 and not allow_zero):
        qualifier = "non-negative" if allow_zero else "positive"
        raise SpecError(f"{where} must be finite and {qualifier}")
    result = round(value * 1000)
    if result > 0xFFFFFFFF:
        raise SpecError(f"{where} exceeds the OOXML millisecond limit")
    return result


def load_parts(path: Path) -> tuple[dict[str, bytes], list[zipfile.ZipInfo]]:
    if not path.is_file():
        raise SpecError(f"input PPTX does not exist: {path}")
    parts: dict[str, bytes] = {}
    infos: list[zipfile.ZipInfo] = []
    try:
        with zipfile.ZipFile(path) as package:
            bad = package.testzip()
            if bad:
                raise SpecError(f"input PPTX has a corrupt ZIP member: {bad}")
            for info in package.infolist():
                if info.filename in parts:
                    raise SpecError(f"input PPTX has duplicate package part: {info.filename}")
                parts[info.filename] = package.read(info.filename)
                infos.append(copy.copy(info))
    except zipfile.BadZipFile as exc:
        raise SpecError(f"input is not a valid PPTX ZIP package: {path}") from exc
    for required in ("[Content_Types].xml", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"):
        if required not in parts:
            raise SpecError(f"input PPTX is missing required part: {required}")
    return parts, infos


def resolve_part(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def slide_parts(parts: Mapping[str, bytes]) -> list[str]:
    presentation = parse_xml(parts["ppt/presentation.xml"])
    rels = parse_xml(parts["ppt/_rels/presentation.xml.rels"])
    relationships = {
        rel.get("Id"): rel
        for rel in rels.findall(qn(PKG_REL_NS, "Relationship"))
    }
    result: list[str] = []
    for index, slide_id in enumerate(
        presentation.findall(f".//{qn(PML_NS, 'sldIdLst')}/{qn(PML_NS, 'sldId')}"), 1
    ):
        rel_id = slide_id.get(qn(R_NS, "id"))
        rel = relationships.get(rel_id)
        if rel is None:
            raise SpecError(f"slide {index} has unresolved relationship {rel_id!r}")
        part = resolve_part("ppt/presentation.xml", rel.get("Target") or "")
        if part not in parts:
            raise SpecError(f"slide {index} part is missing: {part}")
        result.append(part)
    if not result:
        raise SpecError("presentation contains no slides")
    if len(set(result)) != len(result):
        raise SpecError("presentation order references the same slide part more than once")
    return result


def slide_size_emu(parts: Mapping[str, bytes]) -> tuple[int, int]:
    root = parse_xml(parts["ppt/presentation.xml"])
    size = root.find(qn(PML_NS, "sldSz"))
    if size is None:
        raise SpecError("ppt/presentation.xml is missing p:sldSz")
    try:
        width, height = int(size.get("cx", "")), int(size.get("cy", ""))
    except ValueError as exc:
        raise SpecError("presentation slide size is not numeric") from exc
    if width <= 0 or height <= 0:
        raise SpecError("presentation slide size must be positive")
    return width, height


def shape_records(slide_xml: bytes | str) -> list[dict[str, Any]]:
    root = parse_xml(slide_xml)
    records: list[dict[str, Any]] = []
    for prop in root.iter(qn(PML_NS, "cNvPr")):
        parent = prop.getparent()
        shape = parent.getparent() if parent is not None else None
        if shape is None or shape.tag == qn(PML_NS, "spTree"):
            continue
        try:
            shape_id = int(prop.get("id", ""))
        except ValueError:
            continue
        tx_body = shape.find(qn(PML_NS, "txBody"))
        text = "".join(
            node.text or "" for node in tx_body.iter(qn(A_NS, "t"))
        ) if tx_body is not None else ""
        paragraphs = 0
        if tx_body is not None:
            for paragraph in tx_body.findall(qn(A_NS, "p")):
                if "".join(node.text or "" for node in paragraph.iter(qn(A_NS, "t"))).strip():
                    paragraphs += 1
        records.append({
            "id": shape_id,
            "name": prop.get("name") or "",
            "type": etree.QName(shape).localname,
            "has_text": bool(text.strip()),
            "paragraphs": paragraphs,
        })
    ids = [record["id"] for record in records]
    if len(ids) != len(set(ids)):
        raise SpecError("slide contains duplicate shape IDs")
    return records


def resolve_shape(reference: Any, shapes: list[dict[str, Any]], where: str) -> int:
    if isinstance(reference, bool):
        raise SpecError(f"{where} must be a shape name or positive shape ID")
    if isinstance(reference, int):
        positive_int(reference, where)
        matches = [shape for shape in shapes if shape["id"] == reference]
    elif isinstance(reference, str) and reference:
        matches = [shape for shape in shapes if shape["name"] == reference]
    else:
        raise SpecError(f"{where} must be a non-empty shape name or positive shape ID")
    if not matches:
        raise SpecError(f"{where} does not resolve to a shape: {reference!r}")
    if len(matches) != 1:
        raise SpecError(
            f"{where} name {reference!r} is ambiguous; "
            f"matching IDs: {', '.join(str(shape['id']) for shape in matches)}"
        )
    return matches[0]["id"]


def timing_carriers(root: etree._Element) -> list[etree._Element]:
    carriers: list[etree._Element] = []
    for child in root:
        if child.tag == qn(PML_NS, "timing"):
            carriers.append(child)
        elif child.tag == qn(MC_NS, "AlternateContent") and child.find(f".//{qn(PML_NS, 'timing')}") is not None:
            carriers.append(child)
    return carriers


_PATH_TOKEN = re.compile(r"[MLCE]|[-+]?(?:\d+(?:\.\d*)?|\.\d+)")


def tokenize_path(path: Any, where: str) -> list[str]:
    if not isinstance(path, str) or not path.strip():
        raise SpecError(f"{where} must be a non-empty native SVG-style path string")
    compact = re.sub(r"[\s,]+", "", path)
    tokens = _PATH_TOKEN.findall(path)
    if "".join(tokens) != compact:
        raise SpecError(
            f"{where} contains unsupported syntax; use only M, L, C, E and decimal numbers"
        )
    return tokens


def parse_path(path: Any, where: str) -> list[tuple[str, list[float]]]:
    tokens = tokenize_path(path, where)
    commands: list[tuple[str, list[float]]] = []
    index = 0
    current: tuple[str, list[float]] | None = None
    arity = {"M": 2, "L": 2, "C": 6}
    while index < len(tokens):
        token = tokens[index]
        if token in {"M", "L", "C"}:
            if token == "M" and commands:
                raise SpecError(f"{where} supports only one initial M command")
            if not commands and token != "M":
                raise SpecError(f"{where} must start with M")
            current = (token, [])
            commands.append(current)
            index += 1
            continue
        if token == "E":
            if current is None or index != len(tokens) - 1:
                raise SpecError(f"{where} must end with one E after complete coordinates")
            commands.append(("E", []))
            index += 1
            continue
        if current is None or current[0] == "E":
            raise SpecError(f"{where} has coordinates outside a path command")
        try:
            number = float(token)
        except ValueError as exc:
            raise SpecError(f"{where} contains an invalid number: {token}") from exc
        if not math.isfinite(number):
            raise SpecError(f"{where} numbers must be finite")
        current[1].append(number)
        index += 1
    if not commands or commands[-1][0] != "E":
        raise SpecError(f"{where} must end with E")
    for command, values in commands:
        if command == "E":
            continue
        if not values or len(values) % arity[command]:
            raise SpecError(f"{where} has incomplete {command} coordinates")
    if commands[0][0] != "M" or len(commands[0][1]) != 2:
        raise SpecError(f"{where} must start with one M x y pair")
    return commands


def format_number(value: float) -> str:
    if abs(value) < 1e-12:
        value = 0.0
    rendered = format(value, ".12f").rstrip("0").rstrip(".")
    return rendered if rendered not in {"", "-0"} else "0"


def normalize_path(
    path: Any,
    where: str,
    size: tuple[int, int] | None = None,
    canvas_px: tuple[int, int] | None = None,
) -> str:
    commands = parse_path(path, where)
    if canvas_px is not None:
        width, height = canvas_px
        unit_to_emu = 1.0
    elif size is not None:
        width, height = size
        # path_px uses CSS-style pixels at 96 DPI; OOXML slide size uses EMU.
        unit_to_emu = 9525.0
    else:
        width, height = 1, 1
        unit_to_emu = 1.0
    rendered: list[str] = []
    for command, values in commands:
        rendered.append(command)
        if command == "E":
            continue
        converted: list[float] = []
        for index, value in enumerate(values):
            denominator = width if index % 2 == 0 else height
            converted.append(value * unit_to_emu / denominator)
        rendered.extend(format_number(value) for value in converted)
    return " ".join(rendered)


def validate_canvas_px(value: Any, slide_size: tuple[int, int]) -> tuple[int, int]:
    if not isinstance(value, (list, tuple)) or len(value) != 2:
        raise SpecError("spec.canvas_px must be a [width, height] pair")
    width, height = value
    if any(isinstance(item, bool) or not isinstance(item, int) or item <= 0 for item in (width, height)):
        raise SpecError("spec.canvas_px values must be positive integers")
    slide_width, slide_height = slide_size
    # Independent x/y scaling distorts the path when the design canvas and the
    # slide do not share an aspect ratio.
    if abs(width * slide_height - height * slide_width) > max(1, width * slide_height) * 1e-3:
        raise SpecError(
            "spec.canvas_px aspect ratio does not match the slide size: "
            f"{width}x{height} vs {slide_width}x{slide_height} EMU"
        )
    return width, height


def natural_sort_key(name: str) -> tuple[tuple[int, Any], ...]:
    parts = re.split(r"(\d+)", name)
    return tuple(
        (0, int(part)) if part.isdigit() else (1, part)
        for part in parts
        if part
    )


def expand_animation_sugar(
    raw: Any,
    shapes: list[dict[str, Any]],
    where: str,
) -> list[tuple[dict[str, Any], str | None]]:
    """Expand targets/stagger and from_offset_px before schema normalization."""
    item = dict(require_mapping(raw, where))
    if "stagger" in item and "targets" not in item:
        raise SpecError(f"{where}.stagger requires targets")
    if "targets" in item:
        expanded = _expand_targets(item, shapes, where)
        sugar = "targets"
    else:
        expanded = [item]
        sugar = None
    result: list[tuple[dict[str, Any], str | None]] = []
    for entry in expanded:
        label = sugar
        if "build" in entry:
            label = "build" if sugar is None else f"{sugar}+build"
        if "from_offset_px" in entry:
            label = "from_offset_px" if sugar is None else f"{sugar}+from_offset_px"
            if "build" in entry:
                label = f"{sugar}+build+from_offset_px" if sugar else "build+from_offset_px"
        for built in _expand_build(entry, shapes, where):
            result.append((built, label))
    return result


def _expand_build(
    item: dict[str, Any],
    shapes: list[dict[str, Any]],
    where: str,
) -> list[dict[str, Any]]:
    if "build" not in item:
        return [item]
    if item["build"] != "paragraph":
        raise SpecError(f"{where}.build must be 'paragraph'")
    if "iterate" in item:
        raise SpecError(f"{where}.build and iterate are mutually exclusive")
    if "from_offset_px" in item:
        raise SpecError(f"{where}.build and from_offset_px are mutually exclusive")
    if "repeat" in item:
        raise SpecError(f"{where}.build and repeat are mutually exclusive")
    if "trigger_shape" in item:
        raise SpecError(f"{where}.build and trigger_shape are mutually exclusive")
    trigger = item.get("trigger", "after-previous")
    shape_id = resolve_shape(item["target"], shapes, f"{where}.target")
    shape = next((record for record in shapes if record["id"] == shape_id), None)
    if shape is None or not shape.get("has_text") or shape.get("paragraphs", 0) < 1:
        raise SpecError(f"{where}.build requires a text-bearing target shape")
    result: list[dict[str, Any]] = []
    for index in range(shape["paragraphs"]):
        entry = dict(item)
        entry["paragraph"] = index
        if index:
            entry["trigger"] = "after-previous" if trigger == "after-previous" else "on-click"
            entry["delay"] = 0
        result.append(entry)
    return result


def _expand_targets(
    item: dict[str, Any],
    shapes: list[dict[str, Any]],
    where: str,
) -> list[dict[str, Any]]:
    if "target" in item:
        raise SpecError(f"{where}.target and targets are mutually exclusive")
    if "trigger_shape" in item:
        raise SpecError(f"{where}.trigger_shape cannot be combined with targets")
    raw_targets = item["targets"]
    if not isinstance(raw_targets, list) or not raw_targets:
        raise SpecError(f"{where}.targets must be a non-empty array")
    resolved: list[int] = []
    for index, entry in enumerate(raw_targets):
        entry_where = f"{where}.targets[{index}]"
        if isinstance(entry, bool):
            raise SpecError(f"{entry_where} must be a shape name, glob or positive ID")
        if isinstance(entry, int):
            resolved.append(resolve_shape(entry, shapes, entry_where))
            continue
        if not isinstance(entry, str) or not entry:
            raise SpecError(f"{entry_where} must be a shape name, glob or positive ID")
        if any(character in entry for character in "*?["):
            matches = [shape for shape in shapes if fnmatchcase(shape["name"], entry)]
            if not matches:
                raise SpecError(f"{entry_where} glob matched no shape: {entry!r}")
            if len({shape["name"] for shape in matches}) != len(matches):
                raise SpecError(f"{entry_where} glob matched duplicate shape names")
            matches.sort(key=lambda shape: natural_sort_key(shape["name"]))
            resolved.extend(shape["id"] for shape in matches)
        else:
            resolved.append(resolve_shape(entry, shapes, entry_where))
    if len(resolved) != len(set(resolved)):
        raise SpecError(f"{where}.targets resolves the same shape more than once")
    stagger = item.get("stagger", 0)
    if isinstance(stagger, bool) or not isinstance(stagger, (int, float)) or not math.isfinite(stagger) or stagger < 0:
        raise SpecError(f"{where}.stagger must be a non-negative number of seconds")
    base_delay = item.get("delay", 0)
    seconds_to_ms(base_delay, f"{where}.delay")
    result: list[dict[str, Any]] = []
    for index, shape_id in enumerate(resolved):
        entry = {key: value for key, value in item.items() if key not in {"targets", "stagger"}}
        entry["target"] = shape_id
        if index:
            entry["trigger"] = "with-previous"
            entry["delay"] = base_delay + index * stagger
        result.append(entry)
    return result


def _from_offset_values(
    item: dict[str, Any],
    where: str,
) -> tuple[float, float, float]:
    if "path" in item or "path_px" in item:
        raise SpecError(f"{where}.from_offset_px cannot be combined with path or path_px")
    effect = item.get("effect")
    if effect == CUSTOM_MOTION:
        raise SpecError(f"{where}.from_offset_px requires an entrance effect")
    canonical = describe_animation_effect(effect)["effect"]
    if NATIVE_ANIMATIONS[canonical]["category"] != "entrance":
        raise SpecError(f"{where}.from_offset_px requires an entrance effect")
    for forbidden in (
        "trigger_shape", "repeat_count", "repeat_duration", "repeat",
        "auto_reverse", "rewind", "restart", "bounce_end",
    ):
        if forbidden in item:
            raise SpecError(f"{where}.from_offset_px cannot be combined with {forbidden}")
    offset = item["from_offset_px"]
    if not isinstance(offset, (list, tuple)) or len(offset) != 2:
        raise SpecError(f"{where}.from_offset_px must be a [dx, dy] pair")
    if any(
        isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value)
        for value in offset
    ):
        raise SpecError(f"{where}.from_offset_px values must be finite numbers")
    if offset[0] == 0 and offset[1] == 0:
        raise SpecError(f"{where}.from_offset_px cannot be [0, 0]")
    decelerate = item.get("offset_decelerate", 0.85)
    if (
        isinstance(decelerate, bool)
        or not isinstance(decelerate, (int, float))
        or not math.isfinite(decelerate)
        or decelerate < 0
        or decelerate > 1
    ):
        raise SpecError(f"{where}.offset_decelerate must be between 0 and 1")
    return float(offset[0]), float(offset[1]), float(decelerate)


def default_duration(effect: str) -> float:
    canonical = describe_animation_effect(effect)["effect"]
    entry = NATIVE_ANIMATIONS.get(canonical)
    if entry is None or not entry.get("defaultDurationMs"):
        return 0.5
    # Non-scalable presets keep their authored behavior durations; this value
    # only satisfies the common target model and is not written as their dur.
    return entry["defaultDurationMs"] / 1000


def normalize_animation(
    raw: Any,
    shapes: list[dict[str, Any]],
    slide_size: tuple[int, int],
    where: str,
    canvas_px: tuple[int, int] | None = None,
) -> dict[str, Any]:
    item = require_mapping(raw, where)
    reject_unknown(item, ANIMATION_FIELDS, where)
    if "targets" in item or "stagger" in item:
        raise SpecError(f"{where} targets/stagger must be expanded before normalization")
    if "target" not in item or "effect" not in item:
        raise SpecError(f"{where} requires target and effect")
    shape_id = resolve_shape(item["target"], shapes, f"{where}.target")
    shape = next((record for record in shapes if record["id"] == shape_id), None)
    requested_effect = item["effect"]
    custom = requested_effect == CUSTOM_MOTION
    if custom:
        effect = "path_right"
        if ("path" in item) == ("path_px" in item):
            raise SpecError(f"{where} custom_motion requires exactly one of path or path_px")
    else:
        effect = requested_effect
        if "path" in item or "path_px" in item:
            raise SpecError(f"{where}.path/path_px require effect custom_motion")
    canonical = describe_animation_effect(effect)["effect"]
    category = NATIVE_ANIMATIONS[canonical]["category"]
    core: dict[str, Any] = {
        "shape_id": shape_id,
        "effect": effect,
        "duration": item.get("duration", default_duration(effect)),
        "trigger": item.get("trigger", "after-previous"),
        "delay_ms": seconds_to_ms(item.get("delay", 0), f"{where}.delay"),
    }
    if "effect_options" in item:
        core["effect_options"] = item["effect_options"]
    for field in CORE_TIMING_FIELDS:
        if field in item:
            core[field] = item[field]
    if "repeat" in item:
        if "repeat_count" in item or "repeat_duration" in item:
            raise SpecError(f"{where}.repeat is mutually exclusive with repeat_count and repeat_duration")
        repeat_mode = item["repeat"]
        if repeat_mode not in {"until-next-click", "until-end-of-slide"}:
            raise SpecError(
                f"{where}.repeat must be 'until-next-click' or 'until-end-of-slide'"
            )
        core["repeat_mode"] = repeat_mode
    if "offset_decelerate" in item and "from_offset_px" not in item:
        raise SpecError(f"{where}.offset_decelerate requires from_offset_px")
    if "from_offset_px" in item:
        dx, dy, decelerate = _from_offset_values(item, where)
        canvas = canvas_px or (
            round(slide_size[0] / 9525),
            round(slide_size[1] / 9525),
        )
        core["offset_x"] = dx / canvas[0]
        core["offset_y"] = dy / canvas[1]
        core["offset_decelerate"] = decelerate
    if "trigger_shape" in item:
        trigger_id = resolve_shape(item["trigger_shape"], shapes, f"{where}.trigger_shape")
        if trigger_id == shape_id:
            raise SpecError(f"{where} target and trigger_shape must be different shapes")
        core["trigger_shape_id"] = trigger_id
        if core["trigger"] != "on-click":
            raise SpecError(f"{where} with trigger_shape must use trigger on-click")
    if "iterate" in item:
        iteration = require_mapping(item["iterate"], f"{where}.iterate")
        reject_unknown(iteration, {"by", "delay"}, f"{where}.iterate")
        by = iteration.get("by")
        if by not in {"letter", "word"}:
            raise SpecError(f"{where}.iterate.by must be 'letter' or 'word'")
        if category not in {"entrance", "emphasis"}:
            raise SpecError(f"{where}.iterate requires an entrance or emphasis effect")
        if shape is None or not shape.get("has_text") or shape.get("paragraphs", 0) < 1:
            raise SpecError(f"{where}.iterate requires a text-bearing target shape")
        core["iterate_type"] = by
        core["iterate_delay_ms"] = (
            seconds_to_ms(iteration["delay"], f"{where}.iterate.delay")
            if "delay" in iteration
            else None
        )
    if "build" in item:
        build = item["build"]
        if build != "paragraph":
            raise SpecError(f"{where}.build must be 'paragraph'")
        if "iterate" in item:
            raise SpecError(f"{where}.build and iterate are mutually exclusive")
        if category not in {"entrance", "emphasis"}:
            raise SpecError(f"{where}.build requires an entrance or emphasis effect")
        if shape is None or not shape.get("has_text") or shape.get("paragraphs", 0) < 1:
            raise SpecError(f"{where}.build requires a text-bearing target shape")
        if "paragraph" not in item:
            raise SpecError(f"{where}.build must be expanded into paragraph rows")
        core["build"] = "paragraph"
    if "paragraph" in item:
        if "build" not in item:
            raise SpecError(f"{where}.paragraph requires build")
        paragraph = item["paragraph"]
        if isinstance(paragraph, bool) or not isinstance(paragraph, int) or paragraph < 0:
            raise SpecError(f"{where}.paragraph must be a non-negative integer")
        if shape is None or paragraph >= shape.get("paragraphs", 0):
            raise SpecError(f"{where}.paragraph is out of range")
        core["paragraph_index"] = paragraph
    custom_path = None
    if custom:
        if "path_px" in item:
            custom_path = normalize_path(
                item["path_px"], f"{where}.path_px", slide_size, canvas_px
            )
        else:
            custom_path = normalize_path(item["path"], f"{where}.path")
    consumed = {
        "target", "targets", "stagger", "effect", "duration", "delay", "trigger",
        "trigger_shape", "effect_options", "repeat", "path", "path_px",
        "from_offset_px", "offset_decelerate", "iterate", "build", "paragraph",
    } | set(CORE_TIMING_FIELDS)
    unhandled = set(item) - consumed
    if unhandled:
        raise SpecError(f"{where} has unhandled field(s): {', '.join(sorted(unhandled))}")
    return {
        "core": core,
        "requested_effect": requested_effect,
        "custom_path": custom_path,
        "target": item["target"],
    }


def transition_is_morph(transition: Mapping[str, Any] | None) -> bool:
    """True when a normalized transition spec targets Morph."""
    if not transition or "effect" not in transition:
        return False
    effect, _ = normalize_transition_effect_request(
        transition["effect"], transition.get("effect_options"), allow_none=True
    )
    return effect == "morph"


def normalize_spec(spec: Any, parts: Mapping[str, bytes]) -> dict[str, Any]:
    root = require_mapping(spec, "spec")
    reject_unknown(root, TOP_FIELDS, "spec")
    if root.get("version") != 1:
        raise SpecError("spec.version must be 1")
    raw_slides = root.get("slides")
    if not isinstance(raw_slides, list) or not raw_slides:
        raise SpecError("spec.slides must be a non-empty array")
    if "kiosk" in root and not isinstance(root["kiosk"], bool):
        raise SpecError("spec.kiosk must be a boolean")
    ordered_parts = slide_parts(parts)
    size = slide_size_emu(parts)
    canvas_px = validate_canvas_px(root["canvas_px"], size) if "canvas_px" in root else None
    seen: set[int] = set()
    slides: list[dict[str, Any]] = []
    for raw_slide in raw_slides:
        item = require_mapping(raw_slide, "slide spec")
        number = positive_int(item.get("slide"), "slide spec.slide")
        where = f"slide {number}"
        reject_unknown(item, SLIDE_FIELDS, where)
        if number in seen:
            raise SpecError(f"duplicate slide spec: {number}")
        seen.add(number)
        if number > len(ordered_parts):
            raise SpecError(f"slide {number} is out of range; presentation has {len(ordered_parts)} slides")
        replace_existing = item.get("replace_existing", False)
        if not isinstance(replace_existing, bool):
            raise SpecError(f"{where}.replace_existing must be a boolean")
        part = ordered_parts[number - 1]
        shapes = shape_records(parts[part])
        raw_animations = item.get("animations", [])
        if not isinstance(raw_animations, list):
            raise SpecError(f"{where}.animations must be an array")
        animations: list[dict[str, Any]] = []
        expansions: list[dict[str, Any]] = []
        for index, raw in enumerate(raw_animations):
            generated = expand_animation_sugar(raw, shapes, f"{where}.animations[{index}]")
            for expanded, sugar in generated:
                normalized = normalize_animation(
                    expanded,
                    shapes,
                    size,
                    f"{where}.animations[{index}]",
                    canvas_px,
                )
                normalized["source_index"] = index
                normalized["sugar"] = sugar
                animations.append(normalized)
            if len(generated) != 1 or generated[0][1] is not None:
                expansions.append({
                    "source_index": index,
                    "sugar": generated[0][1] if generated else None,
                    "rows": len(generated),
                })
        transition = item.get("transition")
        if transition is not None:
            transition = require_mapping(transition, f"{where}.transition")
            reject_unknown(transition, TRANSITION_FIELDS, f"{where}.transition")
            if not transition:
                raise SpecError(f"{where}.transition cannot be empty")
            if "effect" not in transition and (
                "duration" in transition or "effect_options" in transition
            ):
                raise SpecError(
                    f"{where}.transition duration/effect_options require an explicit effect"
                )
            if "advance_after" in transition:
                seconds_to_ms(
                    transition["advance_after"],
                    f"{where}.transition.advance_after",
                    allow_zero=False,
                )
            if "advance_on_click" in transition and not isinstance(transition["advance_on_click"], bool):
                raise SpecError(f"{where}.transition.advance_on_click must be a boolean")
            if number == 1 and transition_is_morph(transition):
                raise SpecError(f"{where}.transition morph needs a preceding source slide")
        slides.append({
            "slide": number,
            "part": part,
            "replace_existing": replace_existing,
            "animations": animations,
            "requested_count": len(raw_animations),
            "expansions": expansions,
            "transition": transition,
        })
    return {
        "version": 1,
        "kiosk": root.get("kiosk", False),
        "canvas_px": canvas_px,
        "slides": slides,
    }


def animation_expected_order(animations: list[dict[str, Any]]) -> list[dict[str, Any]]:
    regular = [a for a in animations if "trigger_shape_id" not in a["core"]]
    interactive = [a for a in animations if "trigger_shape_id" in a["core"]]
    return regular + interactive


def timing_branches(root: etree._Element) -> list[etree._Element]:
    branches: list[etree._Element] = []
    for carrier in timing_carriers(root):
        if carrier.tag == qn(PML_NS, "timing"):
            branches.append(carrier)
        else:
            branches.extend(carrier.findall(f".//{qn(PML_NS, 'timing')}"))
    return branches


def row_shape_id(row: etree._Element) -> int | None:
    target = row.find(f".//{qn(PML_NS, 'spTgt')}")
    if target is None:
        return None
    try:
        return int(target.get("spid", ""))
    except ValueError:
        return None


def custom_path_operations(
    root: etree._Element,
    animations: list[dict[str, Any]],
    where: str,
    *,
    patch: bool,
    canonicalize: bool = False,
) -> list[dict[str, Any]]:
    """Patch or exactly verify custom paths in every MCE timing branch."""
    expected = animation_expected_order(animations)
    results: list[dict[str, Any]] = []
    for branch_index, timing in enumerate(timing_branches(root)):
        rows = [row for row in timing.iter(qn(PML_NS, "cTn")) if row.get("presetClass") in PRESET_CLASSES]
        if len(rows) != len(expected):
            raise SpecError(f"{where}: timing branch {branch_index} row count mismatch")
        for index, animation in enumerate(expected):
            custom_path = animation["custom_path"]
            if custom_path is None:
                continue
            row = rows[index]
            if row_shape_id(row) != animation["core"]["shape_id"]:
                raise SpecError(f"{where}: custom path target mismatch at row {index + 1}")
            motions = list(row.iter(qn(PML_NS, "animMotion")))
            if len(motions) != 1:
                raise SpecError(f"{where}: custom motion row must contain exactly one animMotion")
            motion = motions[0]
            path_edit_mode = "relative" if animation["core"].get("effect_options", {}).get("relative", True) else "fixed"
            if patch:
                motion.set("path", "M 0 0 L 0.25 0 E" if canonicalize else custom_path)
                motion.set("origin", "layout")
                motion.set("pathEditMode", path_edit_mode)
            else:
                if motion.get("path") != custom_path:
                    raise SpecError(f"{where}: custom path differs at row {index + 1}, branch {branch_index}")
                if motion.get("origin") != "layout" or motion.get("pathEditMode") != path_edit_mode:
                    raise SpecError(f"{where}: custom path coordinate mode differs at row {index + 1}")
            if branch_index == 0:
                results.append({
                    "shape_id": animation["core"]["shape_id"],
                    "path": custom_path,
                    "pathEditMode": path_edit_mode,
                })
    return results


def validate_animation_semantics(
    slide_xml: bytes,
    animations: list[dict[str, Any]],
    where: str,
) -> tuple[Any, list[dict[str, Any]]]:
    """Verify real custom paths, then validate all other native behavior exactly."""
    actual_root = parse_xml(slide_xml)
    custom_paths = custom_path_operations(actual_root, animations, where, patch=False)
    validation_root = copy.deepcopy(actual_root)
    custom_path_operations(
        validation_root, animations, where, patch=True, canonicalize=True
    )
    summary = validate_generated_animation_xml(
        serialize_slide(validation_root),
        [animation["core"] for animation in animations],
    )
    return summary, custom_paths


def timing_media_nodes(carrier: etree._Element) -> list[etree._Element]:
    """Copy p:audio / p:video subtrees carried by one timing carrier."""
    nodes: list[etree._Element] = []
    for tag in ("audio", "video"):
        nodes.extend(copy.deepcopy(node) for node in carrier.iter(qn(PML_NS, tag)))
    return nodes


def _renumber_timing_ids(nodes: list[etree._Element], next_id: int) -> None:
    for node in nodes:
        for ctn in node.iter(qn(PML_NS, "cTn")):
            ctn.set("id", str(next_id))
            next_id += 1


def _tm_root_child_list(fragment: etree._Element, where: str) -> etree._Element:
    for ctn in fragment.iter(qn(PML_NS, "cTn")):
        if ctn.get("nodeType") == "tmRoot":
            child_list = ctn.find(qn(PML_NS, "childTnLst"))
            if child_list is None:
                child_list = etree.SubElement(ctn, qn(PML_NS, "childTnLst"))
            return child_list
    raise SpecError(f"{where}: generated timing tree has no tmRoot")


def _attach_media(fragment: etree._Element, media: list[etree._Element], where: str) -> None:
    if not media:
        return
    child_list = _tm_root_child_list(fragment, where)
    ids = [
        int(ctn.get("id"))
        for ctn in fragment.iter(qn(PML_NS, "cTn"))
        if (ctn.get("id") or "").isdigit()
    ]
    _renumber_timing_ids(media, max(ids, default=1) + 1)
    for node in media:
        child_list.append(node)


def _timing_shell(media: list[etree._Element], where: str) -> etree._Element:
    wrapper = parse_xml(
        f'<root xmlns:p="{PML_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}" '
        f'xmlns:mc="{MC_NS}" xmlns:p14="{P14_NS}"><p:timing><p:tnLst><p:par>'
        f'<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        f'<p:childTnLst/></p:cTn></p:par></p:tnLst></p:timing></root>'.encode("utf-8")
    )
    fragment = copy.deepcopy(wrapper[0])
    _attach_media(fragment, media, where)
    return fragment


def _insert_timing_element(root: etree._Element, fragment: etree._Element) -> None:
    insert_at = len(root)
    for index, child in enumerate(root):
        if child.tag == qn(PML_NS, "extLst"):
            insert_at = index
            break
    root.insert(insert_at, fragment)


def insert_timing(
    slide_xml: bytes,
    animations: list[dict[str, Any]],
    replace_existing: bool,
    where: str,
) -> tuple[bytes, list[dict[str, Any]], int]:
    root = parse_xml(slide_xml)
    carriers = timing_carriers(root)
    if carriers and not replace_existing:
        raise SpecError(
            f"{where} already has an animation timeline; set replace_existing=true to replace it"
        )
    if len(carriers) > 1:
        raise SpecError(f"{where} has multiple logical animation timelines")
    media: list[etree._Element] = []
    if carriers:
        carrier = carriers[0]
        if carrier.tag == qn(MC_NS, "AlternateContent") and carrier.find(f".//{qn(PML_NS, 'transition')}") is not None:
            raise SpecError(f"{where} combines timing and transition in one MCE carrier; refusing unsafe replacement")
        media = timing_media_nodes(carrier)
        root.remove(carrier)
    if not animations:
        if not media:
            return serialize_slide(root), [], 0
        _insert_timing_element(root, _timing_shell(media, where))
        return serialize_slide(root), [], len(media)
    timing_xml = create_sequence_timing_xml([animation["core"] for animation in animations])
    wrapper = parse_xml(
        f'<root xmlns:p="{PML_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}" '
        f'xmlns:mc="{MC_NS}" xmlns:p14="{P14_NS}">{timing_xml}</root>'.encode("utf-8")
    )
    fragment = copy.deepcopy(wrapper[0])
    _attach_media(fragment, media, where)
    _insert_timing_element(root, fragment)
    custom_path_operations(root, animations, where, patch=True)
    output = serialize_slide(root)
    _, custom_paths = validate_animation_semantics(output, animations, where)
    return output, custom_paths, len(media)


def apply_transition(slide_xml: bytes, spec: Mapping[str, Any], where: str) -> tuple[bytes, bool]:
    if spec is None:
        return slide_xml, False
    if "effect" in spec:
        effect = spec["effect"]
        enter = EnterUpdate(
            policy="none" if effect == "none" else "replace",
            effect=effect,
            duration=spec.get("duration", 0.5),
            effect_options=spec.get("effect_options"),
        )
    else:
        enter = EnterUpdate(policy="preserve")
    has_after = "advance_after" in spec
    click = spec.get("advance_on_click")
    if has_after:
        mode = "both" if click else "after"
        advance = AdvanceUpdate(mode=mode, after=spec["advance_after"])
    elif click is True:
        advance = AdvanceUpdate(mode="click")
    elif click is False:
        raise SpecError(f"{where}.transition.advance_on_click=false requires advance_after")
    else:
        advance = AdvanceUpdate(mode="preserve")
    output, uses_timings = apply_slide_motion_xml(
        slide_xml.decode("utf-8"), enter=enter, advance=advance
    )
    return output.encode("utf-8"), uses_timings


def validate_opc(parts: Mapping[str, bytes]) -> None:
    errors: list[str] = []
    for name, data in parts.items():
        if name.endswith(".xml") or name.endswith(".rels"):
            try:
                parse_xml(data)
            except SpecError as exc:
                errors.append(f"{name}: {exc}")
    for rels_name, data in parts.items():
        if not rels_name.endswith(".rels"):
            continue
        try:
            root = parse_xml(data)
        except SpecError:
            continue
        if rels_name == "_rels/.rels":
            source = ""
        else:
            match = re.fullmatch(r"(.+)/_rels/([^/]+)\.rels", rels_name)
            if not match:
                continue
            source = f"{match.group(1)}/{match.group(2)}"
        for rel in root.findall(qn(PKG_REL_NS, "Relationship")):
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target")
            if not target:
                errors.append(f"{rels_name}: relationship {rel.get('Id')} has no target")
                continue
            resolved = resolve_part(source, target)
            if resolved not in parts:
                errors.append(f"{rels_name}: relationship target is missing: {resolved}")
    if errors:
        raise SpecError("; ".join(dict.fromkeys(errors)))


def write_atomic(path: Path, parts: Mapping[str, bytes], infos: list[zipfile.ZipInfo]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    info_by_name = {info.filename: info for info in infos}
    ordered_names = [info.filename for info in infos]
    ordered_names.extend(sorted(set(parts) - set(ordered_names)))
    fd, temp_name = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=path.parent)
    os.close(fd)
    temp = Path(temp_name)
    try:
        with zipfile.ZipFile(temp, "w", allowZip64=True) as package:
            for name in ordered_names:
                info = info_by_name.get(name)
                if info is None:
                    info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                    info.compress_type = zipfile.ZIP_DEFLATED
                    info.external_attr = 0o600 << 16
                package.writestr(copy.copy(info), parts[name])
        os.replace(temp, path)
    finally:
        if temp.exists():
            temp.unlink()


def top_level_forced_shapes(slide_xml: bytes | str) -> dict[str, str]:
    root = parse_xml(slide_xml)
    tree = root.find(f"{qn(PML_NS, 'cSld')}/{qn(PML_NS, 'spTree')}")
    result: dict[str, list[str]] = {}
    if tree is None:
        return {}
    for shape in tree:
        props = shape.find(f".//{qn(PML_NS, 'cNvPr')}")
        if props is None:
            continue
        name = props.get("name") or ""
        if name.startswith("!!"):
            result.setdefault(name, []).append(etree.QName(shape).localname)
    duplicates = [name for name, types in result.items() if len(types) != 1]
    if duplicates:
        raise SpecError(f"duplicate forced-Morph name(s): {', '.join(sorted(duplicates))}")
    return {name: types[0] for name, types in result.items()}


def validate_morph_edges(
    parts: Mapping[str, bytes],
    *,
    require_pairs: bool,
    only_slides: set[int] | None = None,
) -> list[dict[str, Any]]:
    ordered = slide_parts(parts)
    if only_slides is None:
        destinations = list(range(1, len(ordered)))
    else:
        destinations = [
            number - 1
            for number in sorted(only_slides)
            if 2 <= number <= len(ordered)
        ]
    pairs: list[dict[str, Any]] = []
    for destination_index in destinations:
        source_part = ordered[destination_index - 1]
        destination_part = ordered[destination_index]
        transition = read_slide_transition_xml(parts[destination_part])
        if transition.canonical_effect != "morph":
            continue
        if transition.effect_options.get("morph_by") != "object":
            raise SpecError(f"slide {destination_index + 1} Morph must use option byObject")
        source = top_level_forced_shapes(parts[source_part])
        destination = top_level_forced_shapes(parts[destination_part])
        common = sorted(set(source) & set(destination))
        if require_pairs and not common:
            raise SpecError(
                f"Morph edge {destination_index}->{destination_index + 1} has no shared !!key shape names"
            )
        for name in common:
            if source[name] != destination[name]:
                raise SpecError(f"Morph object {name} changes shape type across slides")
            pairs.append({
                "source_slide": destination_index,
                "destination_slide": destination_index + 1,
                "name": name,
                "shape_type": source[name],
            })
    return pairs


def inspect_package(path: Path) -> dict[str, Any]:
    parts, _ = load_parts(path)
    ordered = slide_parts(parts)
    width, height = slide_size_emu(parts)
    slides = []
    for number, part in enumerate(ordered, 1):
        xml = parts[part]
        try:
            animation = asdict(read_slide_animation_sequence(xml))
        except ValueError as exc:
            animation = {"error": str(exc)}
        slides.append({
            "slide": number,
            "part": part,
            "shapes": shape_records(xml),
            "transition": asdict(read_slide_transition_xml(xml)),
            "animation": animation,
        })
    return {"path": str(path), "slide_size_emu": [width, height], "slides": slides}


def animation_motion_paths(slide_xml: bytes | str) -> list[str | None]:
    """animMotion path values in animation-pane row order."""
    root = parse_xml(slide_xml)
    paths: list[str | None] = []
    for ctn in root.iter(qn(PML_NS, "cTn")):
        if ctn.get("presetClass") not in PRESET_CLASSES:
            continue
        motion = next(ctn.iter(qn(PML_NS, "animMotion")), None)
        paths.append(motion.get("path") if motion is not None else None)
    return paths


@lru_cache(maxsize=None)
def stock_motion_path(effect: str | None) -> str | None:
    """The native animMotion path authored for a stock preset, if any."""
    if not effect:
        return None
    entry = NATIVE_ANIMATIONS.get(effect)
    if not entry or entry.get("presetClass") != "path":
        return None
    match = re.search(r'animMotion[^>]*\spath="([^"]*)"', entry["rowXml"])
    return match.group(1) if match else None


def timeline_package(path: Path, slide_number: int | None = None) -> dict[str, Any]:
    parts, _ = load_parts(path)
    ordered = slide_parts(parts)
    if slide_number is not None:
        positive_int(slide_number, "timeline slide")
        if slide_number > len(ordered):
            raise SpecError(
                f"timeline slide {slide_number} is out of range; presentation has {len(ordered)} slides"
            )
    slides: list[dict[str, Any]] = []
    for number, part in enumerate(ordered, 1):
        if slide_number is not None and number != slide_number:
            continue
        xml = parts[part]
        names = {shape["id"]: shape["name"] for shape in shape_records(xml)}
        try:
            sequence = read_slide_animation_sequence(xml)
        except ValueError as exc:
            slides.append({"slide": number, "part": part, "error": str(exc)})
            continue
        motion_paths = animation_motion_paths(xml)
        groups: list[dict[str, Any]] = []
        current: dict[str, Any] | None = None
        click_index = 0
        for row_index, row in enumerate(sequence.rows):
            if row.trigger_shape_id is not None:
                key = f"shape:{row.trigger_shape_id}"
                kind = "shape"
                trigger_shape = names.get(row.trigger_shape_id, f"shape-{row.trigger_shape_id}")
            else:
                if row.trigger == "on-click" or current is None or current["kind"] != "click":
                    click_index += 1
                key = f"click:{click_index}"
                kind = "click"
                trigger_shape = None
            if current is None or current["key"] != key:
                current = {
                    "key": key,
                    "kind": kind,
                    "trigger_shape": trigger_shape,
                    "rows": [],
                }
                groups.append(current)
            start_s = row.offset_ms / 1000.0
            end_s = (
                None
                if row.playback_duration_ms is None
                else start_s + row.playback_duration_ms / 1000.0
            )
            repeat_mode = getattr(row, "repeat_mode", None)
            if repeat_mode is None:
                if row.repeat_count is not None:
                    repeat_mode = f"count:{row.repeat_count:g}"
                elif row.repeat_duration_ms is not None:
                    repeat_mode = f"duration:{row.repeat_duration_ms / 1000.0:g}s"
            motion_path = (
                motion_paths[row_index] if row_index < len(motion_paths) else None
            )
            current["rows"].append({
                "target": names.get(row.shape_id, f"shape-{row.shape_id}"),
                "shape_id": row.shape_id,
                "effect": row.effect,
                "trigger": row.trigger,
                "start_s": round(start_s, 3),
                "end_s": None if end_s is None else round(end_s, 3),
                "duration_s": (
                    None if row.duration_ms is None else round(row.duration_ms / 1000.0, 3)
                ),
                "repeat": repeat_mode,
                "auto_reverse": row.auto_reverse,
                "motion_path": motion_path,
                "custom": bool(
                    motion_path
                    and stock_motion_path(row.effect) is not None
                    and motion_path != stock_motion_path(row.effect)
                ),
            })
        slides.append({"slide": number, "part": part, "groups": groups})
    return {"path": str(path), "slides": slides}


def validate_transition_semantics(
    slide_xml: bytes,
    spec: Mapping[str, Any] | None,
    where: str,
) -> dict[str, Any]:
    actual = read_slide_transition_xml(slide_xml)
    if spec is None:
        return asdict(actual)
    errors: list[str] = []
    if "effect" in spec:
        requested = spec["effect"]
        if requested == "none":
            if actual.canonical_effect is not None:
                errors.append(f"expected no visual transition, found {actual.canonical_effect!r}")
        else:
            effect, explicit_options = normalize_transition_effect_request(
                requested, spec.get("effect_options"), allow_none=False
            )
            if actual.canonical_effect != effect:
                errors.append(f"transition is {actual.canonical_effect!r}; expected {effect!r}")
            expected_duration = seconds_to_ms(
                spec.get("duration", 0.5), f"{where}.duration", allow_zero=False
            )
            if actual.duration_ms != expected_duration:
                errors.append(f"transition duration is {actual.duration_ms}ms; expected {expected_duration}ms")
            expected_options = {
                name: option["default"]
                for name, option in NATIVE_TRANSITIONS[effect]["effectOptions"].items()
            }
            expected_options.update(explicit_options)
            if dict(actual.effect_options) != expected_options:
                errors.append(
                    f"transition options are {dict(actual.effect_options)!r}; expected {expected_options!r}"
                )
    if "advance_after" in spec:
        expected_after = seconds_to_ms(
            spec["advance_after"], f"{where}.advance_after", allow_zero=False
        )
        expected_click = spec.get("advance_on_click", False)
        if actual.advance_after_ms != expected_after:
            errors.append(f"advance_after is {actual.advance_after_ms}ms; expected {expected_after}ms")
        if actual.advance_on_click != expected_click:
            errors.append(f"advance_on_click is {actual.advance_on_click!r}; expected {expected_click!r}")
    elif spec.get("advance_on_click") is True:
        if actual.advance_after_ms is not None or actual.advance_on_click is not True:
            errors.append("expected click-only slide advance")
    if errors:
        raise SpecError(f"{where}: " + "; ".join(errors))
    return asdict(actual)


def read_kiosk_mode(parts: Mapping[str, bytes]) -> bool:
    props = parts.get("ppt/presProps.xml")
    if props is None:
        return False
    root = parse_xml(props)
    show_properties = root.find(qn(PML_NS, "showPr"))
    return (
        show_properties is not None
        and show_properties.find(qn(PML_NS, "kiosk")) is not None
    )


def validate_package(path: Path, spec: Mapping[str, Any] | None = None) -> dict[str, Any]:
    parts, _ = load_parts(path)
    validate_opc(parts)
    validate_pptx_animation_package(path)
    validate_pptx_transition_package(path)
    checked_slides = []
    checked_transitions = []
    morph_slides: set[int] | None = None
    if spec is None:
        unreadable = []
        for number, part in enumerate(slide_parts(parts), 1):
            try:
                read_slide_animation_sequence(parts[part])
            except ValueError as exc:
                unreadable.append(f"slide {number}: {exc}")
        if unreadable:
            raise SpecError("; ".join(unreadable))
    if spec is not None:
        normalized = normalize_spec(spec, parts)
        morph_slides = {
            slide["slide"]
            for slide in normalized["slides"]
            if transition_is_morph(slide["transition"])
        }
        for slide in normalized["slides"]:
            animations = slide["animations"]
            xml = parts[slide["part"]]
            if slide["transition"] is not None:
                transition = validate_transition_semantics(
                    xml, slide["transition"], f"slide {slide['slide']}"
                )
                checked_transitions.append({
                    "slide": slide["slide"],
                    "effect": transition["canonical_effect"],
                    "duration_ms": transition["duration_ms"],
                    "advance_after_ms": transition["advance_after_ms"],
                })
            if animations:
                summary, custom = validate_animation_semantics(
                    xml, animations, f"slide {slide['slide']}"
                )
                checked_slides.append({
                    "slide": slide["slide"], "rows": len(summary.rows), "custom_paths": custom
                })
        if normalized["kiosk"] and not read_kiosk_mode(parts):
            raise SpecError("spec.kiosk is true but the package is not a kiosk show")
        require_timings = any(
            slide["transition"] and "advance_after" in slide["transition"]
            for slide in normalized["slides"]
        )
        if require_timings:
            validate_pptx_transition_package(path, require_use_timings=True)
    morph_pairs = validate_morph_edges(
        parts, require_pairs=bool(morph_slides), only_slides=morph_slides
    )
    return {
        "path": str(path),
        "valid": True,
        "slides": len(slide_parts(parts)),
        "checked_slides": checked_slides,
        "checked_transitions": checked_transitions,
        "morph_pairs": morph_pairs,
        "kiosk": read_kiosk_mode(parts),
        "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
    }


def apply_spec(input_path: Path, output_path: Path, spec: Mapping[str, Any]) -> dict[str, Any]:
    parts, infos = load_parts(input_path)
    normalized = normalize_spec(spec, parts)
    uses_timings = False
    results = []
    for slide in normalized["slides"]:
        number, part = slide["slide"], slide["part"]
        xml = parts[part]
        xml, slide_uses_timings = apply_transition(xml, slide["transition"], f"slide {number}")
        uses_timings = uses_timings or slide_uses_timings
        custom_paths: list[dict[str, Any]] = []
        row_count = 0
        preserved_media = 0
        if slide["animations"] or slide["replace_existing"]:
            xml, custom_paths, preserved_media = insert_timing(
                xml, slide["animations"], slide["replace_existing"], f"slide {number}"
            )
            row_count = len(slide["animations"])
        parts[part] = xml
        results.append({
            "slide": number,
            "part": part,
            "requested_rows": slide["requested_count"],
            "animation_rows": row_count,
            "expanded_rows": row_count,
            "expansions": slide["expansions"],
            "custom_paths": custom_paths,
            "preserved_media": preserved_media,
            "transition": validate_transition_semantics(
                xml, slide["transition"], f"slide {number}"
            ),
        })
    if uses_timings:
        set_package_use_timings(parts, enabled=True)
    if normalized["kiosk"]:
        set_package_kiosk_show(parts)
    validate_opc(parts)
    morph_slides = {
        slide["slide"]
        for slide in normalized["slides"]
        if transition_is_morph(slide["transition"])
    }
    morph_pairs = validate_morph_edges(
        parts, require_pairs=bool(morph_slides), only_slides=morph_slides
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fd, staging_name = tempfile.mkstemp(
        prefix=output_path.name + ".", suffix=".tmp", dir=output_path.parent
    )
    os.close(fd)
    staging = Path(staging_name)
    try:
        write_atomic(staging, parts, infos)
        validate_pptx_animation_package(staging)
        validate_pptx_transition_package(staging, require_use_timings=uses_timings)
        digest = hashlib.sha256(staging.read_bytes()).hexdigest()
        os.replace(staging, output_path)
    except Exception:
        staging.unlink(missing_ok=True)
        raise
    return {
        "input": str(input_path),
        "output": str(output_path),
        "slides_changed": results,
        "slides_unchanged": len(slide_parts(parts)) - len(results),
        "morph_pairs": morph_pairs,
        "sha256": digest,
        "validation": {
            "opc_xml_and_relationships": "passed",
            "native_animation_semantic_readback": "passed",
            "transition_readback": "passed",
            "powerpoint_playback": "not_run_requires_microsoft_powerpoint",
        },
    }


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise SpecError(f"unable to read JSON {path}: {exc}") from exc


def emit(value: Any) -> None:
    print(json.dumps(value, ensure_ascii=False, indent=2))


def animation_description(effect: str) -> dict[str, Any]:
    if effect == CUSTOM_MOTION:
        result = {
            "effect": CUSTOM_MOTION,
            "native_carrier": "path_right / p:animMotion",
            "path": "normalized slide coordinates; M/L/C/E only",
            "path_px": (
                "96-DPI pixels mapped through p:sldSz or the top-level canvas_px; "
                "M/L/C/E only"
            ),
            "effect_options": {"relative": {"type": "boolean", "default": True}},
        }
    else:
        result = describe_animation_effect(effect)
        result.pop("sound", None)
    result["timing"] = {
        "duration": "positive seconds; optional, defaults to the native authored duration",
        "delay": "non-negative seconds",
        "trigger": ["on-click", "with-previous", "after-previous"],
        "trigger_shape": "other unique shape name or positive ID; requires trigger on-click",
        "repeat": ["until-next-click", "until-end-of-slide"],
        "repeat_count": "positive number; mutually exclusive with repeat_duration",
        "repeat_duration": "positive seconds; mutually exclusive with repeat_count",
        "auto_reverse": "boolean",
        "rewind": "boolean",
        "accelerate": "number from 0 to 1",
        "decelerate": "number from 0 to 1; accelerate + decelerate <= 1",
        "bounce_end": "number from 0 to 1; interpolated behaviors only; clears decelerate",
        "restart": ["always", "when-not-active", "never"],
        "iterate": {
            "by": ["letter", "word"],
            "delay": "optional non-negative seconds; defaults to 10% of duration",
        },
        "build": "paragraph; expands to one p:pRg row per paragraph",
        "targets": "array of unique names, positive IDs or fnmatch globs; expands one row per match",
        "stagger": "non-negative seconds between expanded target rows",
        "from_offset_px": "[dx, dy] entrance offset in design pixels; entrance effects only",
        "offset_decelerate": "number from 0 to 1; defaults to 0.85",
    }
    result["canvas_px"] = (
        "top-level [width, height] design canvas; path_px and from_offset_px resolve through it"
    )
    result["sound"] = "not exposed by the high-level interface"
    return result


def transition_description(effect: str) -> dict[str, Any]:
    result = describe_transition_effect(effect)
    result.pop("sound", None)
    result["timing"] = {
        "duration": "positive seconds",
        "advance_after": "non-negative seconds; also sets show useTimings=1",
        "advance_on_click": "boolean; false requires advance_after",
    }
    return result


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    apply_parser = sub.add_parser("apply", help="apply a JSON animation plan")
    apply_parser.add_argument("input", type=Path)
    apply_parser.add_argument("spec", type=Path)
    apply_parser.add_argument("-o", "--output", type=Path, required=True)
    apply_parser.add_argument("--report", type=Path)
    inspect_parser = sub.add_parser("inspect", help="list slide shapes, timings and transitions")
    inspect_parser.add_argument("input", type=Path)
    timeline_parser = sub.add_parser("timeline", help="print animation start/end times per click group")
    timeline_parser.add_argument("input", type=Path)
    timeline_parser.add_argument("--slide", type=int, dest="slide_number")
    validate_parser = sub.add_parser("validate", help="validate a PPTX and optionally a plan")
    validate_parser.add_argument("input", type=Path)
    validate_parser.add_argument("--spec", type=Path)
    list_parser = sub.add_parser("list-effects", help="list native object animation presets")
    list_parser.add_argument("--category", choices=["entrance", "emphasis", "path", "exit"])
    describe_parser = sub.add_parser("describe", help="describe one object animation")
    describe_parser.add_argument("effect")
    sub.add_parser("list-transitions", help="list native page transitions")
    describe_transition_parser = sub.add_parser("describe-transition", help="describe one page transition")
    describe_transition_parser.add_argument("effect")
    return parser


def main() -> None:
    args = build_parser().parse_args()
    try:
        if args.command == "apply":
            report = apply_spec(args.input, args.output, read_json(args.spec))
            if args.report:
                args.report.parent.mkdir(parents=True, exist_ok=True)
                args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            emit(report)
        elif args.command == "inspect":
            emit(inspect_package(args.input))
        elif args.command == "timeline":
            emit(timeline_package(args.input, args.slide_number))
        elif args.command == "validate":
            spec = read_json(args.spec) if args.spec else None
            emit(validate_package(args.input, spec))
        elif args.command == "list-effects":
            effects = [
                key for key in get_available_animations()
                if key in NATIVE_ANIMATIONS and (
                    args.category is None or NATIVE_ANIMATIONS[key]["category"] == args.category
                )
            ]
            emit({"count": len(effects), "effects": effects, "custom_alias": CUSTOM_MOTION})
        elif args.command == "describe":
            emit(animation_description(args.effect))
        elif args.command == "list-transitions":
            emit({"count": len(NATIVE_TRANSITIONS), "effects": get_available_transitions()})
        elif args.command == "describe-transition":
            emit(transition_description(args.effect))
    except (SpecError, ValueError, OSError, KeyError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        raise SystemExit(2) from exc


if __name__ == "__main__":
    main()
