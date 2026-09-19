# Titillium TrueType builds — notice

The `Titillium-*.ttf` files in this folder are **modified** builds of the Titillium
typeface: the original CFF/OpenType outlines converted to TrueType outlines by
`build_titillium_ttf.py` (cubic-to-quadratic conversion, max error 1/1000 em; every
name record, advance width and OpenType layout table is carried over unchanged). They
exist because Microsoft Word on Windows silently substitutes Calibri for CFF/OpenType
fonts when exporting PDF; see the repository README, "Windows: make Word export the
real Titillium".

## Copyright and licence

Titillium is copyright (c) Accademia di Belle Arti di Urbino and the students of the
MA course of Visual Design (Urbino, Italy), released under the **SIL Open Font
License, Version 1.1**. The full licence text is in `OFL.txt` next to this file. The
source OTFs carry no embedded copyright or licence name record; this notice and
`OFL.txt` travel with every bundle that distributes these builds.

Under the OFL these builds may be used, studied, modified and redistributed, with
this notice and the licence included, and may not be sold by themselves. The OFL
also restricts using a **Reserved Font Name** for modified versions when one is
declared by the copyright holder; whether the original Titillium release declares
"Titillium" as a Reserved Font Name is not recorded in the source files and has not
been confirmed. Until it is, treat these builds as **ITQ-internal** (they keep the
family name "Titillium" so that documents naming the brand font resolve to it) and
do not publish them outside ITQ.

Build date: 2026-09-18. Source: the Titillium OTFs shipped in ITQ's `itq-design`
skill (`fonts/`), themselves from the public Titillium release.
