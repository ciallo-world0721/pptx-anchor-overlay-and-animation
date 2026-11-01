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
from functools import lru_cache
from pathlib import Path
from typing import Any, Iterable, Mapping

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
TOP_FIELDS = {"version", "slides", "kiosk"}
SLIDE_FIELDS = {"slide", "replace_existing", "animations", "transition"}
ANIMATION_FIELDS = {
    "target", "effect", "duration", "delay", "trigger", "trigger_shape",
    "effect_options", "repeat_count", "repeat_duration", "auto_reverse",
    "rewind", "accelerate", "decelerate", "bounce_end", "restart",
    "after_effect", "path", "path_px",
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
        records.append({
            "id": shape_id,
            "name": prop.get("name") or "",
            "type": etree.QName(shape).localname,
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


def transition_has_morph(slide_xml: bytes | str) -> bool:
    summary = read_slide_transition_xml(slide_xml)
    return summary.canonical_effect == "morph"


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


def normalize_path(path: Any, where: str, size: tuple[int, int] | None = None) -> str:
    commands = parse_path(path, where)
    width, height = size if size is not None else (1, 1)
    # path_px uses CSS-style pixels at 96 DPI; OOXML slide size uses EMU.
    unit_to_emu = 9525 if size is not None else 1
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


@lru_cache(maxsize=1)
def animation_manifest() -> list[dict[str, Any]]:
    manifest = VENDOR_DIR / "pptx_animation_presets.json"
    return json.loads(manifest.read_text(encoding="utf-8"))["effects"]


def default_duration(effect: str) -> float:
    canonical = describe_animation_effect(effect)["effect"]
    entry = next((item for item in animation_manifest() if item["key"] == canonical), None)
    if entry is None:
        return 0.5
    # Non-scalable presets keep their authored behavior durations; this value
    # only satisfies the common target model and is not written as their dur.
    return (entry["default_duration_ms"] or 500) / 1000


def normalize_animation(
    raw: Any,
    shapes: list[dict[str, Any]],
    slide_size: tuple[int, int],
    where: str,
) -> dict[str, Any]:
    item = require_mapping(raw, where)
    reject_unknown(item, ANIMATION_FIELDS, where)
    if "target" not in item or "effect" not in item:
        raise SpecError(f"{where} requires target and effect")
    shape_id = resolve_shape(item["target"], shapes, f"{where}.target")
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
    if "trigger_shape" in item:
        trigger_id = resolve_shape(item["trigger_shape"], shapes, f"{where}.trigger_shape")
        if trigger_id == shape_id:
            raise SpecError(f"{where} target and trigger_shape must be different shapes")
        core["trigger_shape_id"] = trigger_id
        if core["trigger"] != "on-click":
            raise SpecError(f"{where} with trigger_shape must use trigger on-click")
    custom_path = None
    if custom:
        if "path_px" in item:
            custom_path = normalize_path(item["path_px"], f"{where}.path_px", slide_size)
        else:
            custom_path = normalize_path(item["path"], f"{where}.path")
    return {
        "core": core,
        "requested_effect": requested_effect,
        "custom_path": custom_path,
        "target": item["target"],
    }


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
        animations = [
            normalize_animation(raw, shapes, size, f"{where}.animations[{index}]")
            for index, raw in enumerate(raw_animations)
        ]
        transition = item.get("transition")
        if transition is not None:
            transition = require_mapping(transition, f"{where}.transition")
            reject_unknown(transition, TRANSITION_FIELDS, f"{where}.transition")
            if not transition:
                raise SpecError(f"{where}.transition cannot be empty")
            if "advance_after" in transition:
                seconds_to_ms(transition["advance_after"], f"{where}.transition.advance_after")
            if "advance_on_click" in transition and not isinstance(transition["advance_on_click"], bool):
                raise SpecError(f"{where}.transition.advance_on_click must be a boolean")
        slides.append({
            "slide": number,
            "part": part,
            "replace_existing": replace_existing,
            "animations": animations,
            "transition": transition,
        })
    return {"version": 1, "kiosk": root.get("kiosk", False), "slides": slides}


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


def insert_timing(slide_xml: bytes, animations: list[dict[str, Any]], replace_existing: bool, where: str) -> tuple[bytes, list[dict[str, Any]]]:
    root = parse_xml(slide_xml)
    carriers = timing_carriers(root)
    if carriers and not replace_existing:
        raise SpecError(
            f"{where} already has an animation timeline; set replace_existing=true to replace it"
        )
    if len(carriers) > 1:
        raise SpecError(f"{where} has multiple logical animation timelines")
    if carriers:
        carrier = carriers[0]
        if carrier.tag == qn(MC_NS, "AlternateContent") and carrier.find(f".//{qn(PML_NS, 'transition')}") is not None:
            raise SpecError(f"{where} combines timing and transition in one MCE carrier; refusing unsafe replacement")
        root.remove(carrier)
    native_targets = [animation["core"] for animation in animations]
    timing_xml = create_sequence_timing_xml(native_targets)
    wrapper = parse_xml(
        f'<root xmlns:p="{PML_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}" '
        f'xmlns:mc="{MC_NS}" xmlns:p14="{P14_NS}">{timing_xml}</root>'.encode("utf-8")
    )
    fragment = copy.deepcopy(wrapper[0])
    insert_at = len(root)
    for index, child in enumerate(root):
        if child.tag == qn(PML_NS, "extLst"):
            insert_at = index
            break
    root.insert(insert_at, fragment)
    custom_path_operations(root, animations, where, patch=True)
    output = serialize_slide(root)
    _, custom_paths = validate_animation_semantics(output, animations, where)
    return output, custom_paths


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


def validate_morph_edges(parts: Mapping[str, bytes], *, require_pairs: bool) -> list[dict[str, Any]]:
    ordered = slide_parts(parts)
    pairs: list[dict[str, Any]] = []
    for destination_index in range(1, len(ordered)):
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
        expected_after = seconds_to_ms(spec["advance_after"], f"{where}.advance_after")
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


def validate_package(path: Path, spec: Mapping[str, Any] | None = None) -> dict[str, Any]:
    parts, _ = load_parts(path)
    validate_opc(parts)
    validate_pptx_animation_package(path)
    validate_pptx_transition_package(path)
    morph_pairs = validate_morph_edges(parts, require_pairs=False)
    checked_slides = []
    checked_transitions = []
    if spec is not None:
        normalized = normalize_spec(spec, parts)
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
        if any(
            slide["transition"] and "advance_after" in slide["transition"]
            for slide in normalized["slides"]
        ):
            validate_pptx_transition_package(path, require_use_timings=True)
        if any(
            slide["transition"] and slide["transition"].get("effect") == "morph"
            for slide in normalized["slides"]
        ):
            morph_pairs = validate_morph_edges(parts, require_pairs=True)
    return {
        "path": str(path),
        "valid": True,
        "slides": len(slide_parts(parts)),
        "checked_slides": checked_slides,
        "checked_transitions": checked_transitions,
        "morph_pairs": morph_pairs,
        "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
    }


def apply_spec(input_path: Path, output_path: Path, spec: Mapping[str, Any]) -> dict[str, Any]:
    parts, infos = load_parts(input_path)
    original = dict(parts)
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
        if slide["animations"]:
            xml, custom_paths = insert_timing(
                xml, slide["animations"], slide["replace_existing"], f"slide {number}"
            )
            row_count = len(slide["animations"])
        parts[part] = xml
        results.append({
            "slide": number,
            "part": part,
            "animation_rows": row_count,
            "custom_paths": custom_paths,
            "transition": validate_transition_semantics(
                xml, slide["transition"], f"slide {number}"
            ),
        })
    if uses_timings:
        set_package_use_timings(parts, enabled=True)
    if normalized["kiosk"]:
        set_package_kiosk_show(parts)
    validate_opc(parts)
    validate_morph_edges(
        parts,
        require_pairs=any(
            slide["transition"] and slide["transition"].get("effect") == "morph"
            for slide in normalized["slides"]
        ),
    )
    write_atomic(output_path, parts, infos)
    try:
        validate_pptx_animation_package(output_path)
        validate_pptx_transition_package(output_path, require_use_timings=uses_timings)
        for slide in normalized["slides"]:
            if slide["animations"]:
                validate_animation_semantics(
                    parts[slide["part"]], slide["animations"], f"slide {slide['slide']}"
                )
        morph_pairs = validate_morph_edges(parts, require_pairs=False)
    except Exception:
        output_path.unlink(missing_ok=True)
        raise
    return {
        "input": str(input_path),
        "output": str(output_path),
        "slides_changed": results,
        "slides_unchanged": len(slide_parts(parts)) - len(results),
        "unchanged_parts_verified": sum(
            1 for name, data in original.items() if name in parts and parts[name] == data
        ),
        "morph_pairs": morph_pairs,
        "sha256": hashlib.sha256(output_path.read_bytes()).hexdigest(),
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
            "path_px": "96-DPI pixels converted with p:sldSz; M/L/C/E only",
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
        "repeat_count": "positive number; mutually exclusive with repeat_duration",
        "repeat_duration": "positive seconds; mutually exclusive with repeat_count",
        "auto_reverse": "boolean",
        "rewind": "boolean",
        "accelerate": "number from 0 to 1",
        "decelerate": "number from 0 to 1; accelerate + decelerate <= 1",
        "bounce_end": "number from 0 to 1; interpolated behaviors only; clears decelerate",
        "restart": ["always", "when-not-active", "never"],
    }
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
