"""Read-only snapshots of slide background fills.

This models the background fill, not graphics drawn on a slide or master.
Unsupported fills are reported explicitly rather than changing the document.
"""

from __future__ import annotations

import colorsys
from dataclasses import dataclass

from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn


@dataclass(frozen=True)
class BackgroundInfo:
    """Resolved background data, independent of the source XML.

    kind is solid, gradient, picture, none, or unknown. Gradient angles are
    counter-clockwise degrees. Picture crop values are fractional insets in
    left, top, right, bottom order. Only stretched picture fills are supported.
    reason explains unsupported data. No getter creates or changes XML.
    """

    kind: str
    color: RGBColor | None = None
    stops: tuple = ()
    angle: float = 0.0
    radial: bool = False
    center: tuple = (0.5, 0.5)
    image: bytes | None = None
    crop: tuple = (0.0, 0.0, 0.0, 0.0)
    reason: str | None = None


def background_info(owner):
    """Resolve inheritance and theme background styles without modifying XML."""
    chain = [owner]
    if hasattr(owner, "slide_layout"):
        chain.append(owner.slide_layout)
    if hasattr(chain[-1], "slide_master"):
        chain.append(chain[-1].slide_master)
    master = chain[-1]
    try:
        if any(rel.reltype == RT.THEME_OVERRIDE for item in chain for rel in item.part.rels.values()):
            raise ValueError("theme overrides are not supported")
        theme = master.theme
        color_map = {}
        for item in reversed(chain):
            mapping = item._element.find(qn("p:clrMap"))
            override = item._element.find(qn("p:clrMapOvr"))
            if override is not None:
                mapping = override.find(qn("a:overrideClrMapping"))
            if mapping is not None:
                color_map.update(mapping.attrib)
        for item in chain:
            bg = item._element.cSld.find(qn("p:bg"))
            if bg is None:
                continue
            explicit = bg.find(qn("p:bgPr"))
            if explicit is not None:
                return _fill_info(explicit, item.part, theme, color_map)
            reference = bg.find(qn("p:bgRef"))
            if reference is None:
                raise ValueError("background has no fill or style reference")
            index = int(reference.get("idx", "0"))
            if index == 0:
                return BackgroundInfo("none")
            path = "a:themeElements/a:fmtScheme/" + (
                "a:bgFillStyleLst" if index >= 1001 else "a:fillStyleLst"
            )
            styles = theme._element
            for name in path.split("/"):
                styles = styles.find(qn(name))
                if styles is None:
                    raise ValueError("theme background style list is missing")
            offset = index - 1001 if index >= 1001 else index - 1
            if offset < 0 or offset >= len(styles):
                raise ValueError("theme background style index is out of range")
            placeholder = _color(reference, theme, color_map)
            return _fill_info(
                styles[offset], master.part.part_related_by(RT.THEME), theme, color_map, placeholder
            )
        return BackgroundInfo("solid", color=RGBColor(255, 255, 255))
    except (ValueError, KeyError, AttributeError, TypeError) as exc:
        return BackgroundInfo("unknown", reason=str(exc))


def _color(parent, theme, mapping, placeholder=None):
    choices = {"srgbClr", "schemeClr", "sysClr"}
    node = next((child for child in parent if child.tag.rsplit("}", 1)[-1] in choices), None)
    if node is None:
        raise ValueError("unsupported or missing background color")
    kind = node.tag.rsplit("}", 1)[-1]
    if kind == "srgbClr":
        rgb = RGBColor.from_string(node.get("val"))
    elif kind == "sysClr":
        rgb = RGBColor.from_string(node.get("lastClr"))
    else:
        slot = node.get("val")
        if slot == "phClr":
            if placeholder is None:
                raise ValueError("unresolved placeholder color")
            rgb = placeholder
        else:
            slot = mapping.get(slot, slot)
            rgb = theme.color_scheme[MSO_THEME_COLOR.from_xml(slot)]
    channels = tuple(c / 255 for c in rgb)
    for transform in node:
        name = transform.tag.rsplit("}", 1)[-1]
        value = int(transform.get("val", "0")) / 100000
        if name == "alpha" and value == 1:
            continue
        if name in {"lumMod", "lumOff", "satMod"}:
            h, light, saturation = colorsys.rgb_to_hls(*channels)
            if name == "lumMod":
                light *= value
            elif name == "lumOff":
                light += value
            else:
                saturation *= value
            channels = colorsys.hls_to_rgb(h, max(0, min(1, light)), max(0, min(1, saturation)))
        elif name == "tint":
            channels = tuple(1 - (1 - c) * value for c in channels)
        elif name == "shade":
            channels = tuple(c * value for c in channels)
        else:
            raise ValueError("unsupported background color transform: " + name)
    return RGBColor(*(round(max(0, min(1, c)) * 255) for c in channels))


def _fill_info(parent, part, theme, mapping, placeholder=None):
    if any(parent.find(qn(name)) is not None for name in ("a:effectLst", "a:effectDag")):
        effects = [parent.find(qn(name)) for name in ("a:effectLst", "a:effectDag")]
        if any(effect is not None and len(effect) for effect in effects):
            raise ValueError("background effects are not supported")
    fills = {"solidFill", "gradFill", "blipFill", "noFill"}
    tag = parent.tag.rsplit("}", 1)[-1]
    node = (
        parent
        if tag in fills
        else next((child for child in parent if child.tag.rsplit("}", 1)[-1] in fills), None)
    )
    if node is None:
        raise ValueError("unsupported background fill")
    kind = node.tag.rsplit("}", 1)[-1]
    if kind == "noFill":
        return BackgroundInfo("none")
    if kind == "solidFill":
        return BackgroundInfo("solid", color=_color(node, theme, mapping, placeholder))
    if kind == "blipFill":
        if node.find(qn("a:tile")) is not None:
            raise ValueError("tiled background pictures are not supported")
        stretch = node.find(qn("a:stretch"))
        if stretch is None:
            raise ValueError("background picture has no stretch geometry")
        rect = stretch.find(qn("a:fillRect"))
        if rect is not None and any(int(rect.get(k, "0")) for k in ("l", "t", "r", "b")):
            raise ValueError("inset picture fills are not supported")
        blip = node.find(qn("a:blip"))
        if blip is None or len(blip):
            raise ValueError("missing background picture or unsupported picture effects")
        image = part.related_part(blip.get(qn("r:embed"))).blob
        rect = node.find(qn("a:srcRect"))
        crop = (
            tuple(int(rect.get(k, "0")) / 100000 for k in ("l", "t", "r", "b"))
            if rect is not None
            else (0.0, 0.0, 0.0, 0.0)
        )
        return BackgroundInfo("picture", image=image, crop=crop)
    stop_list = node.find(qn("a:gsLst"))
    if stop_list is None or len(stop_list) < 2:
        raise ValueError("background gradient needs at least two stops")
    stops = tuple(
        (int(s.get("pos")) / 100000, _color(s, theme, mapping, placeholder)) for s in stop_list
    )
    linear = node.find(qn("a:lin"))
    if linear is not None:
        if linear.get("scaled", "0") not in {"0", "false"}:
            raise ValueError("scaled gradient geometry is not supported")
        return BackgroundInfo(
            "gradient", stops=stops, angle=(-int(linear.get("ang", "0")) / 60000) % 360
        )
    path = node.find(qn("a:path"))
    if path is None or path.get("path") != "circle":
        raise ValueError("unsupported gradient path")
    rect = path.find(qn("a:fillToRect"))
    values = (
        tuple(int(rect.get(k, "0")) / 100000 for k in ("l", "t", "r", "b"))
        if rect is not None
        else (0.0, 0.0, 0.0, 0.0)
    )
    if abs(values[0] + values[2] - 1) > 0.00001 or abs(values[1] + values[3] - 1) > 0.00001:
        raise ValueError("radial gradient focal region must be a point")
    return BackgroundInfo("gradient", stops=stops, radial=True, center=values[:2])
