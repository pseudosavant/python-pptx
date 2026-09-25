.. _text_api:

Text-related objects
====================


.. currentmodule:: pptx.text.text


|TextFrame| objects
--------------------

.. autoclass:: TextFrame()
   :members:
   :member-order: bysource
   :undoc-members:


|Font| objects
--------------

The |Font| object is encountered as a property of |_Run|, |_Paragraph|, and in
future other presentation text objects.

.. autoclass:: Font()
   :members:
   :member-order: bysource
   :undoc-members:


|_Paragraph| objects
--------------------

.. autoclass:: _Paragraph()
   :members:
   :member-order: bysource
   :undoc-members:


|_Run| objects
--------------

.. autoclass:: _Run()
   :members:
   :member-order: bysource
   :undoc-members:

Baseline and strikethrough
-------------------------

Font.baseline is an offset as a fraction of font height. Use 0.3 to raise text
and -0.25 to lower it. Font size is set independently. Zero resets the baseline
and None restores inheritance. Font.strike accepts MSO_TEXT_STRIKE_TYPE.NONE,
SINGLE, DOUBLE, or None to inherit.
