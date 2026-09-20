# Titillium — notice

The `Titillium-*.otf` files in this folder are the **official, unmodified** Titillium
OpenType fonts, the ITQ brand typeface, byte-identical to the release ITQ's marketing
department distributes internally. Seven faces: Regular, Regular Italic, Semibold,
Semibold Italic (family "Titillium"), Bold, Bold Italic (family "Titillium Bd") and
Light (family "Titillium Lt").

## Copyright and licence

Titillium is copyright (c) 2009-2011 Accademia di Belle Arti di Urbino and the students
of the MA course of Visual Design (Urbino, Italy), released under the **SIL Open Font
License, Version 1.1**. The full licence text is in `OFL.txt` next to this file; the
font files themselves carry no embedded copyright or licence record, so this notice and
`OFL.txt` travel with every copy. The OFL permits use, redistribution and bundling with
software, provided the licence and notice are included and the fonts are not sold by
themselves.

## Windows and Word: local TrueType conversion

Microsoft Word on Windows silently substitutes Calibri for CFF/OpenType fonts when
laying out and exporting PDF (verified 2026-09-18/19 on an ITQ workstation, per-user and
all-users installs alike). It handles TrueType fine. `install_titillium_ttf.ps1`
therefore runs `build_titillium_ttf.py` **on the user's own machine** to convert these
OTFs to TrueType outlines (same glyphs, metrics and names) into `fonts/ttf/`, and
registers those for the current user. That conversion is a private copy made for use on
that machine; **no modified font files are distributed** with this repository, the
`itq-word-render` skill bundle or the `itq-word-plugin` download.
