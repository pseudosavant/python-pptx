"""Public theme color and font scheme objects."""

from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement


class Theme:
    """A theme shared by one or more masters. Changes affect every reference."""

    def __init__(self, element):
        self._element = element

    @property
    def name(self) -> str:
        return self._element.get("name", "")

    @name.setter
    def name(self, value: str):
        self._element.set("name", value)

    @property
    def color_scheme(self):
        element = self._element.find("./" + qn("a:themeElements") + "/" + qn("a:clrScheme"))
        if element is None:
            raise ValueError("theme does not contain a color scheme")
        return ThemeColorScheme(element)

    @property
    def font_scheme(self):
        element = self._element.find("./" + qn("a:themeElements") + "/" + qn("a:fontScheme"))
        if element is None:
            raise ValueError("theme does not contain a font scheme")
        return ThemeFontScheme(element)


class ThemeColorScheme:
    """RGB colors indexed by the twelve base MSO_THEME_COLOR values.

    System colors return their last saved RGB value. Unsupported color models
    raise ValueError. Assignment replaces only the selected color definition.
    """

    _slots = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
              "accent4", "accent5", "accent6", "hlink", "folHlink")

    def __init__(self, element):
        self._element = element

    @property
    def name(self) -> str:
        return self._element.get("name", "")

    @name.setter
    def name(self, value: str):
        self._element.set("name", value)

    def _slot(self, key):
        name = MSO_THEME_COLOR.to_xml(key)
        if name not in self._slots:
            raise KeyError(key)
        slot = self._element.find(qn("a:" + name))
        if slot is None:
            raise ValueError("theme color slot is missing: " + name)
        return slot

    def __getitem__(self, key) -> RGBColor:
        slot = self._slot(key)
        if len(slot) == 0:
            raise ValueError("theme color slot has no color")
        color = slot[0]
        if color.tag == qn("a:srgbClr"):
            value = color.get("val")
        elif color.tag == qn("a:sysClr"):
            value = color.get("lastClr")
        else:
            raise ValueError("theme color is not RGB or a system color")
        if value is None:
            raise ValueError("theme color has no saved RGB value")
        return RGBColor.from_string(value)

    def __setitem__(self, key, value: RGBColor):
        if not isinstance(value, RGBColor):
            raise TypeError("theme colors require RGBColor values")
        slot = self._slot(key)
        color = OxmlElement("a:srgbClr")
        color.set("val", str(value))
        slot[:] = [color]


class ThemeFontScheme:
    """Latin typefaces of the major (heading) and minor (body) font collections.

    Editing Latin fonts preserves East Asian, complex script, and supplemental fonts.
    """

    def __init__(self, element):
        self._element = element

    def _latin(self, kind):
        latin = self._element.find("./" + qn("a:" + kind + "Font") + "/" + qn("a:latin"))
        if latin is None:
            raise ValueError("theme font collection has no Latin font")
        return latin

    @property
    def major_latin(self) -> str:
        return self._latin("major").get("typeface", "")

    @major_latin.setter
    def major_latin(self, value: str):
        self._latin("major").set("typeface", value)

    @property
    def minor_latin(self) -> str:
        return self._latin("minor").get("typeface", "")

    @minor_latin.setter
    def minor_latin(self, value: str):
        self._latin("minor").set("typeface", value)
