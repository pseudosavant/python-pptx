Themes
======

SlideMaster.theme returns a shared editable Theme. Theme.name is editable.
Theme.color_scheme exposes RGBColor values indexed by the twelve base
MSO_THEME_COLOR enum values. System colors are read using their saved RGB
fallback. Assignment changes only the selected slot. Its name is also editable.

Theme.font_scheme.major_latin and minor_latin set the heading and body Latin
typefaces. Other script definitions and theme extensions are preserved.

Font.theme_font accepts "major", "minor", or None. A selection updates the Latin,
East Asian, and complex script references. None clears these explicit references.

SlideMaster.text_style_font(kind, level=0) reads explicit character defaults.
The kind is "title", "body", or "other" and the level is 0 through 8. The result
is Font or None. Reading missing defaults does not mutate the template.
