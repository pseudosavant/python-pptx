"""Theme package part."""

from pptx.opc.package import XmlPart
from pptx.theme import Theme
from pptx.util import lazyproperty


class ThemePart(XmlPart):
    """DrawingML theme with editable colors and fonts."""

    @lazyproperty
    def theme(self):
        return Theme(self._element)
