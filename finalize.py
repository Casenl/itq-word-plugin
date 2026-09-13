#!/usr/bin/env python3
"""Bake the Word TOC and export a PDF, on whatever platform this happens to be.

The renderer (md_to_itq_docx.py) has no layout engine, so the table of contents
ships as an un-evaluated Word field. Only Word itself can resolve it. This script
drives Word through whichever automation bridge the current platform offers:

    Windows : COM, via PowerShell            (verified working)
    macOS   : AppleScript, via osascript     (VERIFY ON FIRST RUN -- see below)
    other   : nothing; exits 2 with guidance

Exit codes
    0  TOC baked and PDF written
    2  no Word available -- .docx is still valid and self-heals its TOC the
       first time a human opens it in Word (updateFields is set in settings.xml)
    1  Word was available but the automation failed

Design rule: never export a PDF whose TOC was not actually updated. A stale TOC
in a shipped PDF is invisible to the reader and to every programmatic check, so
a failed update must fail loudly rather than degrade quietly.

The macOS AppleScript below was written against Word's documented object model
but has NOT been executed on macOS hardware. If it fails, the exact script and
the osascript error are printed so the command names can be corrected against
Word's own AppleScript dictionary (Script Editor > File > Open Dictionary >
Microsoft Word). Nothing else in the pipeline depends on it.
"""

import argparse
import os
import platform
import shutil
import subprocess
import sys
import tempfile


# exit 3 inside the PowerShell child means "Word is not installed here"; the
# parent maps it to this script's exit 2 (no Word) rather than 1 (failure).
_PS_NO_WORD = 3

# A sync client (Syncthing, OneDrive) scanning the freshly written .docx can
# hold a transient lock, which Word reports as "cannot complete the save due
# to a file permission error". One retry after a short pause clears that case;
# a persistent error (read-only file, real ACL problem) still fails loudly.
SAVE_ATTEMPTS = 2
SAVE_RETRY_SECONDS = 3

# Kept as a standalone function (not inlined in the template) so the retry
# behaviour can be exercised in pwsh against a fake document object, without
# Word. Not a .format template: no brace doubling here.
PS_SAVE_FUNCTION = r'''
function Save-WithRetry($doc, $attempts, $pauseSeconds) {
    for ($attempt = 1; $attempt -le $attempts; $attempt++) {
        try {
            $doc.Save()
            return
        } catch {
            if ($attempt -ge $attempts) { throw }
            [Console]::Error.WriteLine('finalize: save attempt ' + $attempt + ' failed (' + $_.Exception.Message.Trim() + '); retrying in ' + $pauseSeconds + 's')
            Start-Sleep -Seconds $pauseSeconds
        }
    }
}
'''

# Cleanup discipline: the document is closed WITHOUT saving and Word is quit
# in `finally`, both wrapped so one failure cannot skip the other. Before this,
# a failed Save left a dirty document open, Quit() prompted "save changes?" in
# an invisible window, and a headless WINWORD /Automation process stayed alive
# holding the .docx locked. DisplayAlerts=0 (wdAlertsNone) keeps every other
# dialog from blocking too.
#
# Read-only guard: Save() on a document Word opened read-only (file attribute,
# or another process holding it) does NOT raise -- it parks a Save As dialog
# behind the invisible window and blocks forever, alerts off or not (verified
# 2026-09-13). Refusing up front turns that hang into a clear exit 1.
POWERSHELL_TEMPLATE = r'''
$ErrorActionPreference = 'Stop'
$docx = {docx}
$pdf  = {pdf}
try {{
    $word = New-Object -ComObject Word.Application
}} catch {{
    [Console]::Error.WriteLine('WORD_UNAVAILABLE: ' + $_.Exception.Message)
    exit {no_word}
}}
$word.Visible = $false
$word.DisplayAlerts = 0   # wdAlertsNone
$doc = $null
$quitError = $null
try {{
    $doc = $word.Documents.Open($docx)
    if ($doc.ReadOnly) {{
        throw ('document opened read-only (file attribute or another process holds it): ' + $docx)
    }}
    foreach ($toc in $doc.TablesOfContents) {{ $toc.Update() }}
    $doc.Fields.Update() | Out-Null
    Save-WithRetry $doc {save_attempts} {retry_seconds}
    if ($pdf) {{ $doc.ExportAsFixedFormat($pdf, 17) }}   # 17 = wdExportFormatPDF
}} finally {{
    if ($null -ne $doc) {{ try {{ $doc.Close(0) }} catch {{}} }}   # 0 = wdDoNotSaveChanges
    try {{ $word.Quit(0) }} catch {{ $quitError = $_.Exception.Message }}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}}
# A Quit() failure after a successful save must not masquerade as success: an
# invisible WINWORD may still be alive and holding the .docx locked. Every
# cleanup step above still ran; now report it.
if ($quitError) {{
    [Console]::Error.WriteLine('WORD_QUIT_FAILED: ' + $quitError + ' (a WINWORD /Automation process may still hold the document; check Task Manager)')
    exit {quit_failed}
}}
'''
# exit 4 inside the PowerShell child: the work succeeded but Word did not quit.
_PS_QUIT_FAILED = 4


# Reports tocUpdated=true/false on stdout. The PDF is exported ONLY when an
# update actually succeeded, and the document/app are cleaned up on every path
# so a failed run cannot leave the .docx locked open behind an orphaned Word.
APPLESCRIPT_TEMPLATE = '''
set docxPath to "{docx}"
set pdfPath to "{pdf}"
set tocUpdated to false
tell application "Microsoft Word"
    set wasRunning to running
    set theDoc to open file name docxPath with read only false
    try
        try
            repeat with t in (get tables of contents of theDoc)
                update t
            end repeat
            set tocUpdated to true
        end try
        try
            update fields of theDoc
            set tocUpdated to true
        end try
        if tocUpdated then
            save theDoc
            if pdfPath is not "" then
                save as theDoc file name pdfPath file format format PDF
            end if
        end if
        close theDoc saving no
        if not wasRunning then quit
    on error errMsg number errNum
        try
            close theDoc saving no
        end try
        try
            if not wasRunning then quit
        end try
        error errMsg number errNum
    end try
end tell
return "tocUpdated=" & (tocUpdated as string)
'''


def _ps_quote(value):
    """Single-quote a string for PowerShell (doubling embedded quotes)."""
    if not value:
        return "''"
    return "'" + value.replace("'", "''") + "'"


def _as_quote(value):
    """Escape a string for embedding in an AppleScript literal."""
    return value.replace('\\', '\\\\').replace('"', '\\"')


def _powershell_exe():
    """Prefer PowerShell 7+ (always UTF-8) over Windows PowerShell 5.1, which
    decodes BOM-less script files using the ANSI code page and so mangles
    non-ASCII characters in paths."""
    return shutil.which('pwsh') or shutil.which('powershell')


def build_powershell_script(docx, pdf):
    return PS_SAVE_FUNCTION + POWERSHELL_TEMPLATE.format(
        docx=_ps_quote(docx), pdf=_ps_quote(pdf or ''), no_word=_PS_NO_WORD,
        quit_failed=_PS_QUIT_FAILED,
        save_attempts=SAVE_ATTEMPTS, retry_seconds=SAVE_RETRY_SECONDS)


def finalize_windows(docx, pdf):
    exe = _powershell_exe()
    if not exe:
        print('finalize: no pwsh/powershell on PATH', file=sys.stderr)
        return 1
    script = build_powershell_script(docx, pdf)
    handle, path = tempfile.mkstemp(suffix='.ps1')
    try:
        # utf-8-sig: the BOM makes Windows PowerShell 5.1 decode the file as
        # UTF-8 too, so a path containing e.g. 'e-acute' survives either host.
        with os.fdopen(handle, 'w', encoding='utf-8-sig') as fh:
            fh.write(script)
        proc = subprocess.run(
            [exe, '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', path],
            capture_output=True, text=True,
        )
    finally:
        os.unlink(path)
    if proc.returncode == _PS_NO_WORD:
        print('finalize: Microsoft Word is not installed on this host', file=sys.stderr)
        return 2
    if proc.returncode != 0:
        print('finalize: Word COM automation failed', file=sys.stderr)
        print(proc.stderr.strip(), file=sys.stderr)
        return 1
    return 0


def finalize_macos(docx, pdf):
    if not os.path.isdir('/Applications/Microsoft Word.app'):
        print('finalize: Microsoft Word is not installed', file=sys.stderr)
        return 2
    script = APPLESCRIPT_TEMPLATE.format(docx=_as_quote(docx), pdf=_as_quote(pdf or ''))
    proc = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)

    def _report(headline):
        print('finalize: ' + headline, file=sys.stderr)
        print('This macOS path is unverified on real hardware. Correct the command', file=sys.stderr)
        print("names against Word's AppleScript dictionary and re-run.", file=sys.stderr)
        if proc.stderr.strip():
            print('\n--- osascript error ---\n' + proc.stderr.strip(), file=sys.stderr)
        print('\n--- script that was run ---' + script, file=sys.stderr)

    if proc.returncode != 0:
        _report('Word AppleScript automation failed.')
        return 1
    if 'tocUpdated=true' not in proc.stdout:
        # Word accepted neither update command. Exporting now would ship a PDF
        # with a stale TOC, which no downstream check can detect.
        _report('Word accepted no TOC update command, so NO PDF was written.')
        return 1
    return 0


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--docx', required=True, help='Path to the rendered .docx')
    ap.add_argument('--pdf', default=None,
                    help='Path for the exported PDF. Omit to bake the TOC only.')
    args = ap.parse_args()

    docx = os.path.abspath(args.docx)
    pdf = os.path.abspath(args.pdf) if args.pdf else None
    if not os.path.isfile(docx):
        print('finalize: no such file: ' + docx, file=sys.stderr)
        return 1

    system = platform.system()
    if system == 'Windows':
        rc = finalize_windows(docx, pdf)
    elif system == 'Darwin':
        rc = finalize_macos(docx, pdf)
    else:
        rc = 2
        print('finalize: no Word automation on ' + system + '.', file=sys.stderr)

    if rc == 2:
        print('', file=sys.stderr)
        print('The .docx is still valid and correctly branded. Its TOC carries', file=sys.stderr)
        print('updateFields, so Word rebuilds the page numbers the first time a', file=sys.stderr)
        print('human opens it. Hand off the .docx; do not ship a PDF from here.', file=sys.stderr)
        return 2

    if rc == 0:
        print('finalize: TOC baked' + (' and PDF written to ' + pdf if pdf else ''))
    return rc


if __name__ == '__main__':
    sys.exit(main())
