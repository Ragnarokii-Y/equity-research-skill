#!/usr/bin/env python3
r"""Inspect local PPTX structure; never extracts media or modifies the input.

Usage: python -B inspect_prototype.py --input INPUT.pptx --output G:\Codex\TASK\prototype.json
Add --include-text only for an authorized local working paper; default omits source prose.
"""
import argparse
import collections
import datetime
import json
import os
from pathlib import Path
import posixpath
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
EMU_PER_PT = 12700
OUTPUT_ROOT = Path(r"G:\Codex")
ROLES = ["cover", "executive_summary", "company_divider", "history", "ownership",
         "revenue_profit", "margins_expenses", "segment_financials", "business_overview",
         "business_A", "business_B", "business_C_or_new_products", "industry_divider",
         "driver_1", "driver_2", "driver_3", "chain_A", "competition_A", "chain_B",
         "competition_B", "chain_C", "competition_C", "structural_trend",
         "valuation_divider", "earnings_valuation", "ending"]


def pt(value):
    return round(int(value) / EMU_PER_PT, 4)


def relationships(archive, part):
    location = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    if location not in archive.namelist():
        return {}
    result = {}
    for rel in ET.fromstring(archive.read(location)):
        target = rel.get("Target", "")
        mode = rel.get("TargetMode", "Internal")
        result[rel.get("Id")] = {
            "type": rel.get("Type", "").rsplit("/", 1)[-1],
            "mode": mode,
            "target": target if mode == "External" else posixpath.normpath(
                posixpath.join(posixpath.dirname(part), target.lstrip("/"))) if not target.startswith("/") else target.lstrip("/"),
        }
    return result


def related_part(archive, part, kind):
    return next((r["target"] for r in relationships(archive, part).values()
                 if r["type"] == kind and r["mode"] != "External"), None)


def placeholder(shape):
    node = shape.find("./p:nvSpPr/p:nvPr/p:ph", NS)
    return dict(node.attrib) if node is not None else None


def find_inherited(shape, ancestor):
    ph = placeholder(shape)
    if ph is None or ancestor is None:
        return None
    candidates = ancestor.findall("./p:cSld/p:spTree/p:sp", NS)
    for candidate in candidates:
        other = placeholder(candidate)
        if other is not None and other.get("idx", "0") == ph.get("idx", "0"):
            return candidate
    for candidate in candidates:
        other = placeholder(candidate)
        if other is not None and other.get("type", "obj") == ph.get("type", "obj"):
            return candidate
    return None


def transform(shape):
    if shape is None:
        return None
    for xpath in ["./p:spPr/a:xfrm", "./p:xfrm", "./p:grpSpPr/a:xfrm"]:
        node = shape.find(xpath, NS)
        if node is None:
            continue
        off, extent = node.find("a:off", NS), node.find("a:ext", NS)
        if off is not None and extent is not None:
            result = {"x_pt": pt(off.get("x")), "y_pt": pt(off.get("y")),
                      "width_pt": pt(extent.get("cx")), "height_pt": pt(extent.get("cy"))}
            for key in ["rot", "flipH", "flipV"]:
                if key in node.attrib:
                    result[key] = node.get(key)
            ch_off, ch_ext = node.find("a:chOff", NS), node.find("a:chExt", NS)
            if ch_off is not None and ch_ext is not None:
                result["group_child_coordinates_emu"] = dict(ch_off.attrib) | dict(ch_ext.attrib)
            return result
    return None


def font_record(node):
    result = dict(node.attrib)
    if "sz" in result:
        result["size_pt"] = int(result.pop("sz")) / 100
    for kind in ["latin", "ea", "cs"]:
        child = node.find(f"a:{kind}", NS)
        if child is not None:
            result[kind] = child.get("typeface")
    for kind in ["srgbClr", "schemeClr", "sysClr"]:
        child = node.find(f"a:solidFill/a:{kind}", NS)
        if child is not None:
            result["color"] = {"kind": kind, **child.attrib}
    return result


def style_records(shape):
    if shape is None:
        return []
    return [font_record(x) for x in shape.findall("./p:txBody/a:lstStyle/*/a:defRPr", NS)]


def inspect_shape(shape, layout, master, include_text, group_depth=0):
    kind = shape.tag.rsplit("}", 1)[-1]
    if kind in ["nvGrpSpPr", "grpSpPr", "extLst"]:
        return None
    inherited_layout = find_inherited(shape, layout)
    inherited_master = find_inherited(shape, master)
    direct_position = transform(shape)
    inherited_position = transform(inherited_layout) or transform(inherited_master)
    text = "".join(x.text or "" for x in shape.findall(".//a:t", NS))
    result = {"kind": kind, "group_depth": group_depth,
              "geometry": direct_position or inherited_position,
              "geometry_source": "direct" if direct_position else "placeholder_inheritance" if inherited_position else "unknown",
              "placeholder": placeholder(shape), "text_char_count": len(text),
              "paragraph_count": len(shape.findall(".//a:p", NS)),
              "has_chart": shape.find(".//a:graphicData", NS) is not None and any(
                  node.get("uri", "").endswith("/chart") for node in shape.findall(".//a:graphicData", NS)),
              "has_table": shape.find(".//a:tbl", NS) is not None,
              "explicit_font_properties": [font_record(x) for x in shape.findall(".//a:rPr", NS) + shape.findall(".//a:defRPr", NS)],
              "inherited_layout_font_defaults": style_records(inherited_layout),
              "inherited_master_font_defaults": style_records(inherited_master)}
    body = shape.find("./p:txBody/a:bodyPr", NS)
    if body is not None:
        result["text_frame_properties"] = dict(body.attrib)
        result["autofit_mode"] = [x.tag.rsplit("}", 1)[-1] for x in body]
    if include_text:
        result["text"] = text
    if kind == "grpSp":
        result["geometry_note"] = "Nested child geometries use group-local coordinates; render to assess final placement."
        result["children"] = [record for child in shape for record in [
            inspect_shape(child, layout, master, include_text, group_depth + 1)] if record is not None]
    return result


def inspect(archive, include_text):
    pres = ET.fromstring(archive.read("ppt/presentation.xml"))
    size = pres.find("p:sldSz", NS)
    rels = relationships(archive, "ppt/presentation.xml")
    ids = pres.findall("p:sldIdLst/p:sldId", NS)
    slides = [rels[x.get("{" + NS["r"] + "}id")]["target"] for x in ids]
    result = {"slide_count": len(slides), "size_emu": dict(size.attrib),
              "size_pt": {"width": pt(size.get("cx")), "height": pt(size.get("cy"))},
              "source_text_included": include_text, "media_extracted": False,
              "notes": ["XML direct font values do not fully resolve all PowerPoint style inheritance.",
                        "Geometry is not a visual overflow check; inspect PowerPoint renders."],
              "themes": [], "pages": []}
    for name in sorted(n for n in archive.namelist() if n.startswith("ppt/theme/theme") and n.endswith(".xml")):
        root = ET.fromstring(archive.read(name))
        scheme = root.find(".//a:clrScheme", NS)
        fonts = root.find(".//a:fontScheme", NS)
        result["themes"].append({"part": name, "color_scheme": [
            {"role": x.tag.rsplit("}", 1)[-1], "values": [dict(v.attrib) for v in x]} for x in scheme] if scheme is not None else [],
            "font_families": [dict(x.attrib) for x in fonts.iter() if "typeface" in x.attrib] if fonts is not None else []})
    font_sizes, families, colors = collections.Counter(), collections.Counter(), collections.Counter()
    for index, name in enumerate(slides, 1):
        root = ET.fromstring(archive.read(name))
        layout_name = related_part(archive, name, "slideLayout")
        layout = ET.fromstring(archive.read(layout_name)) if layout_name else None
        master_name = related_part(archive, layout_name, "slideMaster") if layout_name else None
        master = ET.fromstring(archive.read(master_name)) if master_name else None
        tree = root.find("p:cSld/p:spTree", NS)
        records = [r for s in tree for r in [inspect_shape(s, layout, master, include_text)] if r is not None]
        page_rels = relationships(archive, name)
        result["pages"].append({"page": index, "role_if_26_pages": ROLES[index - 1] if len(slides) == 26 else None,
                                "slide_part": name, "layout_part": layout_name, "master_part": master_name,
                                "shapes": records, "relationship_type_counts": dict(collections.Counter(r["type"] for r in page_rels.values())),
                                "external_relationship_count": sum(r["mode"] == "External" for r in page_rels.values())})
        for prop in root.findall(".//a:rPr", NS) + root.findall(".//a:defRPr", NS):
            if prop.get("sz"):
                font_sizes[int(prop.get("sz")) / 100] += 1
        for node in root.findall(".//a:latin", NS) + root.findall(".//a:ea", NS):
            families[node.get("typeface")] += 1
        for node in root.findall(".//a:srgbClr", NS):
            colors[node.get("val")] += 1
    result["explicit_style_summary"] = {"font_size_counts": dict(font_sizes), "font_family_counts": dict(families), "rgb_color_counts": dict(colors)}
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--include-text", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()
    source, output = args.input.resolve(), args.output.resolve()
    if source == output:
        parser.error("Input and output must be different paths.")
    if not source.is_file():
        parser.error("Input must be an existing local PPTX file.")
    root = OUTPUT_ROOT.resolve()
    if os.path.normcase(os.path.commonpath([str(root), str(output)])) != os.path.normcase(str(root)):
        parser.error("All output must be stored within G:\\Codex.")
    with zipfile.ZipFile(source) as archive:
        result = inspect(archive, args.include_text)
    result["inspected_at"] = datetime.datetime.now(datetime.timezone.utc).isoformat()
    result["local_input"] = str(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w" if args.overwrite else "x", encoding="utf-8") as stream:
        json.dump(result, stream, ensure_ascii=False, indent=2)
    print(json.dumps({"output": str(output), "pages": result["slide_count"], "source_text_included": args.include_text}, ensure_ascii=False))


if __name__ == "__main__":
    main()

