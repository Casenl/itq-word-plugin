<#
.SYNOPSIS
  Make Word on Windows render and embed the real ITQ brand font "Titillium".

.DESCRIPTION
  Word's PDF export on Windows silently substitutes Calibri for CFF/OpenType (.otf)
  fonts - verified 2026-09-18 on Michel's workstation with an unrelated system-wide
  OTF as control. It handles TrueType (.ttf) fine. This script replaces the per-user
  registration of the Titillium .otf faces with the TrueType builds produced by
  build_titillium_ttf.py (same outlines, same family/style names), so a document that
  names "Titillium" exports a PDF that actually embeds Titillium. No admin needed.
  The .otf files stay on disk; -Rollback restores the previous registration exactly.

  Fonts are looked up in fonts\ttf next to this script, or one level up (the
  portable skill keeps the script under scripts\ and the fonts under fonts\ttf).

  .\install_titillium_ttf.ps1              # install / replace (safe to re-run)
  .\install_titillium_ttf.ps1 -DryRun      # show what would happen, change nothing
  .\install_titillium_ttf.ps1 -Rollback    # restore the .otf registration

  -UserFonts / -RegKey / -FontDir / -NoGdi exist for the test-suite (temp dir,
  throwaway HKCU key, no GDI session calls); leave them alone in normal use.
#>
param(
  [switch]$Rollback,
  [switch]$DryRun,
  [string]$FontDir,
  [string]$UserFonts = (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'),
  [string]$RegKey = 'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts',
  [switch]$NoGdi
)
$ErrorActionPreference = 'Stop'
$Here   = Split-Path -Parent $MyInvocation.MyCommand.Path
$Backup = Join-Path $UserFonts 'titillium-otf-registration.backup.json'

Add-Type @"
using System; using System.Runtime.InteropServices;
public class TitFont {
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode)] public static extern int AddFontResource(string f);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode)] public static extern int RemoveFontResource(string f);
  [DllImport("user32.dll")] public static extern int SendMessage(IntPtr h, int m, IntPtr w, IntPtr l);
  public static void Broadcast() { SendMessage((IntPtr)0xffff, 0x001D, IntPtr.Zero, IntPtr.Zero); }
}
"@
function Gdi-Add($p)    { if (-not $NoGdi) { [TitFont]::AddFontResource($p)    | Out-Null } }
function Gdi-Remove($p) { if (-not $NoGdi) { [TitFont]::RemoveFontResource($p) | Out-Null } }
function Gdi-Broadcast  { if (-not $NoGdi) { [TitFont]::Broadcast() } }
function Same-File($a, $b) {
  (Test-Path $a) -and (Test-Path $b) -and ((Get-FileHash $a).Hash -eq (Get-FileHash $b).Hash)
}
function Reg-Get($name) { (Get-ItemProperty -Path $RegKey -Name $name -EA SilentlyContinue).$name }

# registry value name -> ttf file. Names match what the OTF installer registered.
$Faces = [ordered]@{
  'Titillium (TrueType)'                 = 'Titillium-Regular.ttf'
  'Titillium Regular Italic (TrueType)'  = 'Titillium-RegularItalic.ttf'
  'Titillium Semibold (TrueType)'        = 'Titillium-Semibold.ttf'
  'Titillium Semibold Italic (TrueType)' = 'Titillium-SemiboldItalic.ttf'
  'Titillium Bold (TrueType)'            = 'Titillium-Bold.ttf'
  'Titillium Bold Italic (TrueType)'     = 'Titillium-BoldItalic.ttf'
  'Titillium Light (TrueType)'           = 'Titillium-Light.ttf'
}

if (-not $Rollback -and -not $FontDir) {
  # Rollback needs only the backup + installed files, never the download.
  foreach ($cand in @((Join-Path $Here 'fonts\ttf'), (Join-Path (Split-Path -Parent $Here) 'fonts\ttf'))) {
    if (Test-Path $cand) { $FontDir = (Resolve-Path $cand).Path; break }
  }
  if (-not $FontDir) { throw "no fonts\ttf next to or above $Here - run: python build_titillium_ttf.py" }
}

if ($DryRun -and $Rollback) {
  if (-not (Test-Path $Backup)) { throw "no backup at $Backup - nothing to roll back" }
  $prev = Get-Content $Backup -Raw | ConvertFrom-Json
  "backup      : $Backup"
  foreach ($name in $Faces.Keys) {
    $ttf = Join-Path $UserFonts $Faces[$name]
    $old = $prev.$name
    $action = if ($old -and ($old -ieq $ttf)) { "keep (pre-existing)" } elseif ($old) { "restore -> $old" } else { "remove registration" }
    "  {0,-38} {1}{2}" -f $name, $action, $(if ((Test-Path $ttf) -and -not ($old -ieq $ttf)) { "; delete $ttf" } else { "" })
  }
  "Dry run - nothing changed."; exit 0
}

if ($DryRun) {
  "font source : $FontDir"
  "user fonts  : $UserFonts $(if (Test-Path $UserFonts) {''} else {'(will be created)'})"
  "registry    : $RegKey $(if (Test-Path $RegKey) {''} else {'(will be created)'})"
  foreach ($name in $Faces.Keys) {
    $src = Join-Path $FontDir $Faces[$name]
    $state = if (Test-Path $src) { 'ok' } else { 'MISSING' }
    "  {0,-38} {1}  [{2}]  currently: {3}" -f $name, $src, $state, (Reg-Get $name)
  }
  "Dry run - nothing changed."; exit 0
}

if ($Rollback) {
  if (-not (Test-Path $Backup)) { throw "no backup at $Backup - nothing to roll back" }
  $prev = Get-Content $Backup -Raw | ConvertFrom-Json
  foreach ($name in $Faces.Keys) {
    $ttf = Join-Path $UserFonts $Faces[$name]
    $old = $prev.$name
    if ($old -and ($old -ieq $ttf)) {
      # The face was already registered at this very path before we ran: not ours to
      # delete. If we overwrote a different file there, put the original bytes back.
      $bak = "$ttf.pre-itq.bak"
      if (Test-Path $bak) { Gdi-Remove $ttf; Copy-Item $bak $ttf -Force; Remove-Item $bak -Force; Gdi-Add $ttf }
      Set-ItemProperty -Path $RegKey -Name $name -Value $old; "kept      $name -> $old (pre-existing$(if (Test-Path $ttf) {', original bytes restored'}))"
      continue
    }
    if (Test-Path $ttf) { Gdi-Remove $ttf; Remove-Item $ttf -Force }
    if ($old) { Set-ItemProperty -Path $RegKey -Name $name -Value $old; Gdi-Add $old; "restored  $name -> $old" }
    else      { Remove-ItemProperty -Path $RegKey -Name $name -EA SilentlyContinue; "removed   $name (was not registered before)" }
  }
  Gdi-Broadcast; Remove-Item $Backup -Force
  "Rolled back. Restart Word to pick up the change."; exit 0
}

foreach ($name in $Faces.Keys) {
  if (-not (Test-Path (Join-Path $FontDir $Faces[$name]))) { throw "missing $($Faces[$name]) in $FontDir - run: python build_titillium_ttf.py" }
}
# Fresh profiles have neither the per-user Fonts folder nor the HKCU Fonts key yet.
if (-not (Test-Path $UserFonts)) { New-Item -ItemType Directory -Path $UserFonts -Force | Out-Null; "created $UserFonts" }
if (-not (Test-Path $RegKey))    { New-Item -Path $RegKey -Force | Out-Null; "created $RegKey" }

if (-not (Test-Path $Backup)) {
  $snap = [ordered]@{}
  foreach ($name in $Faces.Keys) { $snap[$name] = Reg-Get $name }
  $snap | ConvertTo-Json | Set-Content $Backup -Encoding UTF8
  "backed up previous registration -> $Backup"
}
foreach ($name in $Faces.Keys) {
  $src = Join-Path $FontDir $Faces[$name]
  $dst = Join-Path $UserFonts $Faces[$name]
  $old = Reg-Get $name
  if ($old -and $old -ne $dst -and (Test-Path $old)) { Gdi-Remove $old }
  if (Same-File $src $dst) {
    # AddFontResource is reference-counted and rollback removes exactly one, so
    # only load when the registry did not already point here (= not loaded by us).
    if ($old -ieq $dst) { "unchanged  $name ($dst already current, already loaded)" }
    else { Gdi-Add $dst; "unchanged  $name ($dst already current, loaded now)" }
  } else {
    if (Test-Path $dst) {
      Gdi-Remove $dst   # a loaded font file cannot be overwritten
      # A different file was already registered at this very path: keep its bytes
      # so -Rollback can put them back (see the "kept" branch there).
      if (($old -ieq $dst) -and -not (Test-Path "$dst.pre-itq.bak")) { Copy-Item $dst "$dst.pre-itq.bak" }
    }
    try { Copy-Item $src $dst -Force }
    catch { throw "cannot replace $dst (in use). Close Word/Outlook/PowerPoint and re-run. ($_)" }
    Gdi-Add $dst
    "installed  $name -> $dst"
  }
  New-ItemProperty -Path $RegKey -Name $name -Value $dst -PropertyType String -Force | Out-Null
}
Gdi-Broadcast
"Done. Restart Word (or any open Office app) to pick up the change. Rollback: -Rollback"
