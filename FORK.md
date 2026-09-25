# ps-python-pptx

This fork retains the `pptx` import namespace. It is based on upstream master
at `278b47b1` and is distributed separately as `ps-python-pptx`.

## Upstream contribution structure

Feature branches are based on upstream master or an existing upstream proposal.
Fork naming, versioning, release automation, and integration formatting are
separate commits. No upstream pull requests are automatically submitted.

| Branch | Scope | Related upstream work |
| --- | --- | --- |
| `codex/portable-test-fixtures` | Current dependencies, Windows fixtures, isolated mocks | Test maintenance |
| `codex/shape-alternative-text` | Descriptions, titles, and removal semantics | Builds on #1097 by Huan Di |
| `codex/shape-stacking-order` | Send to back and bring to front | Independent |
| `codex/table-style-reference` | Validated table style GUIDs | Independent |
| `codex/font-character-formatting` | Baseline and strikethrough | Related to #1046, independent implementation |
| `codex/paragraph-list-formatting` | Bullet styles, starts, indents, paragraph marks | Builds on #1060 by Thomas Vakili |
| `codex/theme-authoring` | Native theme editing and font references | Independent |
| `codex/slide-management` | Clear slides and master graphics visibility | Related to #1029, independent implementation |
| `codex/background-authoring` | Explicit gradients and embedded backgrounds | Independent |

The original commits from #1097 and #1060 retain their authorship. Coordinate
with those existing proposals when preparing upstream contributions. Integration
merges combine overlapping APIs. Include relevant formatting changes when
preparing a focused upstream patch. Do not include fork packaging commits.

## Public API documentation

API details are in `docs/api/shapes.rst`, `text.rst`, `table.rst`, `slides.rst`,
`dml.rst`, and `theme.rst`. Round-trip tests cover saving and reopening files.
Theme edits preserve unmodified theme data. Shared themes affect all masters
that reference them. Master text style access reads explicit defaults without
resolving the complete inheritance hierarchy.

## Releases

Create a GitHub release tagged `v` followed by the package version. The
`publish-pypi.yml` workflow runs CI, builds and checks distributions, then uses
PyPI Trusted Publishing with the `pypi` GitHub environment. No PyPI token is used.

Pending publisher settings are project `ps-python-pptx`, owner `pseudosavant`,
repository `python-pptx`, workflow `publish-pypi.yml`, and environment `pypi`.
The first fork release is 1.1.0. The upstream base remains identifiable in Git.
