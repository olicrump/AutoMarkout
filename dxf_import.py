
"""
Finds all block references (INSERT entities) and their attributes (ATTRIB)
in a given DXF file, including inserts nested inside block definitions.

Anonymous block aliases
-----------------------
When a block definition is edited in-place in AutoCAD/BricsCAD, the software
sometimes re-saves it under an anonymous name (e.g. *U12, *U14) while keeping
the original named definition as well. These anonymous blocks are structurally
identical to their named counterpart (same ATTDEFs, same geometry) but appear
as separate block definitions in the DXF. This module detects these aliases by
comparing ATTDEF tag sets and resolves them back to the canonical named block,
so all insertions are reported under a single consistent name.

Usage:
    from find_block_refs import get_block_references

    refs = get_block_references("drawing.dxf")

    # Optionally filter by the name of the inserted block:
    refs = get_block_references("drawing.dxf", block_name="My Multi Block")

Each entry in the returned list is a dict:
    {
        "block_name":        str,   # canonical (named) block name, alias-resolved
        "block_name_raw":    str,   # actual block name as stored in the DXF
        "parent":            str,   # block definition or layout containing this insert
        "handle":            str,
        "layer":             str,
        "insertion_point":   {"x": float, "y": float, "z": float},
        "rotation_deg":      float,
        "scale":             {"x": float, "y": float, "z": float},
        "attributes":        {tag: {"value": str, "layer": str, "handle": str}, ...},
    }
"""

import sys
from pathlib import Path
import ezdxf
from ezdxf.entities import Insert, AttDef


def _build_alias_map(doc) -> dict[str, str]:
    """
    Return a mapping of anonymous block name -> canonical named block name
    for any anonymous block (*U/*E etc.) whose ATTDEF tag set exactly matches
    a named block definition. Named blocks are preferred over anonymous ones.
    """

    def attdef_tags(blk) -> frozenset[str]:
        return frozenset(
            e.dxf.tag for e in blk if isinstance(e, AttDef)
        )

    # Index named blocks (non-anonymous, non-layout) by their attdef signature
    named_by_signature: dict[frozenset, str] = {}
    for blk in doc.blocks:
        if blk.is_any_layout or blk.name.startswith("*"):
            continue
        sig = attdef_tags(blk)
        if sig:  # only consider blocks that actually have ATTDEFs
            named_by_signature[sig] = blk.name

    # Map anonymous blocks to their named equivalent where signatures match
    alias_map: dict[str, str] = {}
    for blk in doc.blocks:
        if not blk.name.startswith("*") or blk.is_any_layout:
            continue
        sig = attdef_tags(blk)
        if sig and sig in named_by_signature:
            alias_map[blk.name] = named_by_signature[sig]

    return alias_map


def get_block_references(
        dxf_path: str | Path,
        block_name: str | None = None,
) -> list[dict]:
    """
    Return all INSERT-based block references and their attributes from a DXF
    file, including those nested inside block definitions.

    Anonymous blocks that are structurally identical to a named block (same
    ATTDEF tags) are resolved to the canonical named block name. The original
    raw name is preserved in the "block_name_raw" field.

    Args:
        dxf_path:   Path to the .dxf file.
        block_name: If given, only return references to this block
                    (matched against the resolved canonical name,
                    case-insensitive).

    Returns:
        List of dicts describing each INSERT and its ATTRIBs.

    Raises:
        FileNotFoundError:       If the file does not exist.
        ezdxf.DXFStructureError: If the file cannot be parsed.
    """
    dxf_path = Path(dxf_path)
    if not dxf_path.exists():
        raise FileNotFoundError(f"DXF file not found: {dxf_path}")

    doc = ezdxf.readfile(str(dxf_path))
    alias_map = _build_alias_map(doc)

    results = []

    for block_def in doc.blocks:
        for entity in block_def:
            if not isinstance(entity, Insert):
                continue

            raw_name = entity.dxf.name
            canonical_name = alias_map.get(raw_name, raw_name)

            if block_name and canonical_name.lower() != block_name.lower():
                continue

            ip = entity.dxf.insert
            scale = (
                entity.dxf.get("xscale", 1.0),
                entity.dxf.get("yscale", 1.0),
                entity.dxf.get("zscale", 1.0),
            )

            attributes = {
                attrib.dxf.tag.strip(): {
                    "value": attrib.dxf.text.strip(),
                    "layer": attrib.dxf.get("layer", "0"),
                    "handle": attrib.dxf.handle,
                }
                for attrib in entity.attribs
            }

            results.append({
                "block_name": canonical_name,
                "block_name_raw": raw_name,
                "parent": block_def.name,
                "handle": entity.dxf.handle,
                "layer": entity.dxf.get("layer", "0"),
                "insertion_point": {"x": round(ip.x, 1), "y": round(ip.y, 1), "z": round(ip.z, 1)},
                "rotation_deg": round(entity.dxf.get("rotation", 0.0), 4),
                "scale": {"x": scale[0], "y": scale[1], "z": scale[2]},
                "attributes": attributes,
            })

    return results


def count_block_references(dxf_path: str | Path) -> dict[str, int]:
    """
    Return a dictionary of block names and the number of times each is
    inserted in the given DXF file.

    Anonymous blocks that are aliases of a named block (same ATTDEF tags)
    are resolved to the canonical name, consistent with get_block_references().

    Args:
        dxf_path: Path to the .dxf file.

    Returns:
        Dict mapping canonical block name -> insertion count.

    Raises:
        FileNotFoundError:       If the file does not exist.
        ezdxf.DXFStructureError: If the file cannot be parsed.
    """
    from collections import Counter

    refs = get_block_references(dxf_path)
    counts: Counter[str] = Counter(r["block_name"] for r in refs)
    return dict(counts)

if __name__ == '__main__':

    filename1 = "/Users/olicrump/My Drive/Audio Work/20260401 Rick Astley/CAD/Old/0417 Manchester COOP (RA 26).dxf"
    filename2 = "/Users/olicrump/My Drive/Audio Work/20260401 Rick Astley/CAD/Old/RA BG V1.4.dxf"

    # importer = DXF_Importer(filename2)
    # # importer.find_anon()
    # importer.findnames()
    listd = get_block_references(filename2,block_name="My Multi Block 2025 PRACTICE")
    print(listd)