<#
.SYNOPSIS
  Make Word on Windows render and embed the ITQ brand font "Titillium".

.DESCRIPTION
  Word's PDF export on Windows silently substitutes Calibri for CFF/OpenType (.otf)
  fonts - verified 2026-09-18/19 on an ITQ workstation, per-user and all-users installs
  alike, with an unrelated system-wide OTF as control. It handles TrueType (.ttf) fine.

  This script converts the OFFICIAL Titillium OTFs shipped in fonts\otf to TrueType ON
  THIS MACHINE (build_titillium_ttf.py, same outlines, metrics and names; needs Python
  with fonttools) into fonts\ttf, then registers those faces for the current user in
  place of any .otf registration of the same names. No admin needed. Nothing modified is
  ever distributed: the .ttf files exist only where they were built. The .otf files stay
  on disk; -Rollback restores the previous registration exactly.

  Layouts: fonts\otf (and fonts\ttf) next to this script, or one level up (the portable
  skill keeps scripts\ apart). A fonts\ttf that already holds all seven faces is reused.

  .\install_titillium_ttf.ps1              # convert if needed, install / replace (safe to re-run)
  .\install_titillium_ttf.ps1 -Rebuild     # force a fresh conversion first
  .\install_titillium_ttf.ps1 -DryRun      # show what would happen, change nothing
  .\install_titillium_ttf.ps1 -Rollback    # restore the previous registration

  -Python names the interpreter (default: python on PATH). -UserFonts / -RegKey /
  -FontDir / -NoGdi exist for the test-suite; leave them alone in normal use.
#>
param(
  [switch]$Rollback,
  [switch]$DryRun,
  [switch]$Rebuild,
  [string]$FontDir,
  [string]$Python = 'python',
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
  [DllImport("user32.dll", SetLastError=true)] public static extern IntPtr SendMessageTimeout(IntPtr h, uint m, IntPtr w, IntPtr l, uint flags, uint timeout, out IntPtr result);
  // WM_FONTCHANGE to every top-level window. A plain SendMessage(HWND_BROADCAST) blocks forever
  // on one unresponsive window (hit 2026-09-19: an orphan headless WINWORD); SMTO_ABORTIFHUNG
  // (0x2) + 2 s cap keeps the installer from hanging on someone else's stuck app.
  public static void Broadcast() { IntPtr r; SendMessageTimeout((IntPtr)0xffff, 0x001D, IntPtr.Zero, IntPtr.Zero, 0x2, 2000, out r); }
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
function Complete-TtfDir($dir) {
  if (-not $dir -or -not (Test-Path $dir)) { return $false }
  foreach ($f in $Faces.Values) { if (-not (Test-Path (Join-Path $dir $f))) { return $false } }
  return $true
}

# ---------------------------------------------------------------- source discovery (install only)
$OtfDir = $null; $Builder = $null; $NeedBuild = $false
if (-not $Rollback -and -not $FontDir) {
  $roots = @($Here, (Split-Path -Parent $Here))
  foreach ($r in $roots) { $c = Join-Path $r 'fonts\otf'; if (-not $OtfDir -and (Test-Path (Join-Path $c 'Titillium-Regular.otf'))) { $OtfDir = (Resolve-Path $c).Path } }
  foreach ($r in $roots) { $c = Join-Path $r 'fonts\ttf'; if (-not $FontDir -and -not $Rebuild -and (Complete-TtfDir $c)) { $FontDir = (Resolve-Path $c).Path } }
  if (-not $FontDir) {
    if (-not $OtfDir) { throw "no fonts\otf (official Titillium OTFs) next to or above $Here - the bundle is incomplete, re-download it" }
    $Builder = Join-Path $Here 'build_titillium_ttf.py'
    if (-not (Test-Path $Builder)) { throw "missing $Builder next to this script - the bundle is incomplete, re-download it" }
    $FontDir = Join-Path (Split-Path -Parent $OtfDir) 'ttf'
    $NeedBuild = $true
  }
}

function Assert-PythonWithFonttools {
  $ok = $false
  try { & $Python -c "import fontTools" 2>$null; $ok = ($LASTEXITCODE -eq 0) } catch { $ok = $false }
  if (-not $ok) {
    throw ("Converting the official Titillium OTFs to TrueType needs Python with the fonttools package " +
           "(interpreter tried: '$Python'). Install Python 3, then run:  $Python -m pip install fonttools  " +
           "and re-run this script. Nothing was changed.")
  }
}

# ---------------------------------------------------------------- previews
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
  "official otf: $(if ($OtfDir) { $OtfDir } else { '(none found; using -FontDir)' })"
  "truetype dir: $FontDir $(if ($NeedBuild) { '(will be BUILT here from the official OTFs on this machine)' } else { '(prebuilt, reused)' })"
  if ($NeedBuild) { "converter   : $Builder via '$Python' (needs fonttools)" }
  "user fonts  : $UserFonts $(if (Test-Path $UserFonts) {''} else {'(will be created)'})"
  "registry    : $RegKey $(if (Test-Path $RegKey) {''} else {'(will be created)'})"
  foreach ($name in $Faces.Keys) {
    $src = Join-Path $FontDir $Faces[$name]
    $state = if ($NeedBuild) { 'to build' } elseif (Test-Path $src) { 'ok' } else { 'MISSING' }
    "  {0,-38} {1}  [{2}]  currently: {3}" -f $name, $src, $state, (Reg-Get $name)
  }
  "Dry run - nothing changed."; exit 0
}

# ---------------------------------------------------------------- rollback
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

# ---------------------------------------------------------------- local conversion
if ($NeedBuild) {
  Assert-PythonWithFonttools
  "converting the official Titillium OTFs in $OtfDir to TrueType on this machine -> $FontDir"
  # Anything the converter prints on stderr would become a terminating NativeCommandError
  # under 'Stop'; judge it by its exit code instead.
  $ErrorActionPreference = 'Continue'
  $out = & $Python $Builder --src $OtfDir --out $FontDir 2>&1
  $code = $LASTEXITCODE
  $ErrorActionPreference = 'Stop'
  $out | ForEach-Object { "  $_" }
  if ($code -ne 0) { throw "build_titillium_ttf.py failed (exit $code); nothing was registered" }
  if (-not (Complete-TtfDir $FontDir)) { throw "conversion did not produce all seven faces in $FontDir" }
  "built 7 TrueType faces (private copies for this machine; not for redistribution)"
}
foreach ($name in $Faces.Keys) {
  if (-not (Test-Path (Join-Path $FontDir $Faces[$name]))) { throw "missing $($Faces[$name]) in $FontDir - run with -Rebuild" }
}

# ---------------------------------------------------------------- install
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
