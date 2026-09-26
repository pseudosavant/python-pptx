.. _action_api:

Click Action-related Objects
============================

The following classes represent click and hover mouse actions, typically
a hyperlink. Other actions such as navigating to another slide in the
presentation or running a macro are also possible.


|ActionSetting| objects
-----------------------

.. autoclass:: pptx.action.ActionSetting()
   :members:
   :inherited-members:
   :undoc-members:


|Hyperlink| objects
-----------------------

.. autoclass:: pptx.action.Hyperlink()
   :members:
   :inherited-members:
   :undoc-members:

ScreenTips
----------

Text-run hyperlinks and shape click or hover hyperlinks expose ``screen_tip``.
Set the address first, then assign the ScreenTip text. Assign ``None`` or an
empty string to clear just the tip. Assigning a new address clears the old tip.
ScreenTip access does not require direct XML manipulation.
