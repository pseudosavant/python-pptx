ps-python-pptx
==============

This is the pseudosavant fork of python-pptx, published as ps-python-pptx.
It retains the pptx import package and the upstream MIT license and attribution.

Install with pip install ps-python-pptx. Use a fresh environment when switching
from python-pptx because both distributions provide the same import package.

Version 1.1.0 adds public APIs for alternative text and titles, shape stacking,
table styles, baseline shifts, strikethrough, paragraph bullets and indents,
theme colors and fonts, slide clearing, and explicit gradient/picture backgrounds.
See FORK.md for feature branches, upstream proposals, and release conventions.

Upstream project documentation follows.

*python-pptx* is a Python library for creating, reading, and updating PowerPoint (.pptx)
files.

A typical use would be generating a PowerPoint presentation from dynamic content such as
a database query, analytics output, or a JSON payload, perhaps in response to an HTTP
request and downloading the generated PPTX file in response. It runs on any Python
capable platform, including macOS and Linux, and does not require the PowerPoint
application to be installed or licensed.

It can also be used to analyze PowerPoint files from a corpus, perhaps to extract search
indexing text and images.

In can also be used to simply automate the production of a slide or two that would be
tedious to get right by hand, which is how this all got started.

More information is available in the `python-pptx documentation`_.

Browse `examples with screenshots`_ to get a quick idea what you can do with
python-pptx.

.. _`python-pptx documentation`:
   https://python-pptx.readthedocs.org/en/latest/

.. _`examples with screenshots`:
   https://python-pptx.readthedocs.org/en/latest/user/quickstart.html
