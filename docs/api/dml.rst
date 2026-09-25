.. _dml_api:

DrawingML objects
=================

Low-level drawing elements like fill and color that appear repeatedly in
various aspects of shapes.


|ChartFormat| objects
---------------------

.. autoclass:: pptx.dml.chtfmt.ChartFormat
   :members:


|FillFormat| objects
--------------------

.. autoclass:: pptx.dml.fill.FillFormat
   :members:
   :exclude-members: from_fill_parent
   :undoc-members:


|LineFormat| objects
--------------------

.. autoclass:: pptx.dml.line.LineFormat
   :members:
   :undoc-members:


|ColorFormat| objects
---------------------

.. autoclass:: pptx.dml.color.ColorFormat
   :members: brightness, rgb, theme_color, type
   :undoc-members:


|RGBColor| objects
------------------

.. autoclass:: pptx.dml.color.RGBColor
   :members: from_string
   :undoc-members:


|ShadowFormat| objects
----------------------

.. autoclass:: pptx.dml.effect.ShadowFormat
   :members:
   :undoc-members:

Explicit gradients
------------------

FillFormat.set_gradient(stops, angle=0.0, radial=False, center=(0.5, 0.5))
replaces a fill with explicit colors and positions. Stops are pairs of position
and RGBColor or MSO_THEME_COLOR. Positions must be ascending fractions in the
range 0..1 with at least two stops. Repeated positions are allowed for abrupt
color changes. Invalid values do not change the existing fill. Linear angles
are counter-clockwise degrees. Radial fills use a circular path and a fractional
center measured from the upper-left corner.
