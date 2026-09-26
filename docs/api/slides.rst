.. _slides_api:

Slides
======

|Slides| objects
-----------------

The |Slides| object is accessed using the
:attr:`~pptx.presentation.Presentation.slides` property of |Presentation|. It
is not intended to be constructed directly.

.. autoclass:: pptx.slide.Slides()
   :members:
   :member-order: bysource
   :undoc-members:


|Slide| objects
---------------

An individual |Slide| object is accessed by index from |Slides| or as the
return value of :meth:`add_slide`.

.. autoclass:: pptx.slide.Slide()
   :members:
   :exclude-members: part
   :inherited-members:
   :undoc-members:


|SlideLayouts| objects
----------------------

The |SlideLayouts| object is accessed using the
:attr:`~pptx.slide.SlideMaster.slide_layouts` property of |SlideMaster|, typically::

    >>> from pptx import Presentation
    >>> prs = Presentation()
    >>> slide_layouts = prs.slide_master.slide_layouts

As a convenience, since most presentations have only a single slide master, the
|SlideLayouts| collection for the first master may be accessed directly from the
|Presentation| object::

    >>> slide_layouts = prs.slide_layouts

This class is not intended to be constructed directly.

.. autoclass:: pptx.slide.SlideLayouts()
   :members:
   :exclude-members: element, parent
   :inherited-members:
   :undoc-members:


|SlideLayout| objects
---------------------

.. autoclass:: pptx.slide.SlideLayout
   :members:
   :exclude-members: iter_cloneable_placeholders


|SlideMasters| objects
----------------------

The |SlideMasters| object is accessed using the
:attr:`~pptx.presentation.slide_masters` property of |Presentation|, typically::

    >>> from pptx import Presentation
    >>> prs = Presentation()
    >>> slide_masters = prs.slide_masters

As a convenience, since most presentations have only a single slide master, the
first master may be accessed directly from the |Presentation| object without indexing
the collection::

    >>> slide_master = prs.slide_master

This class is not intended to be constructed directly.

.. autoclass:: pptx.slide.SlideMasters()
   :members:
   :exclude-members: element, parent
   :inherited-members:
   :undoc-members:

|SlideMaster| objects
---------------------

.. autoclass:: pptx.slide.SlideMaster
   :members:
   :exclude-members: related_slide_layout, sldLayoutIdLst


|SlidePlaceholders| objects
---------------------------

.. autoclass:: pptx.shapes.shapetree.SlidePlaceholders
   :members:


|NotesSlide| objects
--------------------

.. autoclass:: pptx.slide.NotesSlide
   :members:
   :exclude-members: clone_master_placeholders
   :inherited-members:

Slide collection and master graphics
-----------------------------------

Slides.clear() removes all slides and retains masters and layouts for reuse as
a template. Shared parts referenced from retained parts are preserved.
Slide.show_master_shapes and SlideLayout.show_master_shapes accept True, False,
or None. None removes the explicit setting and uses the file format default.

Picture backgrounds
-------------------

Slides, layouts, and masters expose set_background_picture(image_file). The
argument is a path or file-like object. The image is embedded and stretched
to fill the background. Existing background fill proxies reflect the change.

Read-only background inspection
------------------------------

``slide.background_info`` returns an immutable snapshot of the effective fill.
The same property is available on layouts and masters. It follows background
inheritance and theme style references without changing the presentation.
Unlike ``background.fill``, this getter never interrupts inheritance.

The snapshot's ``kind`` is ``solid``, ``gradient``, ``picture``, ``none``, or
``unknown``. It includes resolved RGB colors, gradient stops and geometry, or
embedded picture bytes and crop fractions as applicable. ``reason`` describes
an unsupported fill. Pattern fills, translucent colors, scaled linear
gradients, and unsupported picture effects are reported as unknown.

This inspects background fills only. It does not composite slide or master
shapes, resolve theme overrides, or reproduce PowerPoint's text rendering.

::

    info = slide.background_info
    if info.kind == "solid":
        print(info.color)
    elif info.kind == "unknown":
        print(info.reason)
