# Render test for consent.html against the reference consent PDF (Windows: Edge + pdftotext + Windows.Data.Pdf).
#   powershell -File tests\consent-test.ps1 -ReferencePdf "C:\path\Consent_for_esketamine_treatment_updated.pdf"
# Prints the form twice with headless Edge (blank, and with FICTIONAL values), then checks:
#   1. Both printouts are 4 Letter pages, like the reference.
#   2. Page by page, the text matches the reference word for word (filled values removed first).
#   3. Page by page, every printed line starts within -Tol px (at 150 dpi) of the same line in the reference.
# Output (PDFs, page PNGs, red/blue overlays) goes to tests\out\, which is not committed. Exit code 1 on any failure.
param(
  [Parameter(Mandatory = $true)][string]$ReferencePdf,
  [int]$Tol = 3,
  [string]$Edge = "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
  [string]$PdfToText = ''
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$out = Join-Path $PSScriptRoot 'out'; New-Item -ItemType Directory -Force $out | Out-Null
if (-not $PdfToText) {
  $c = Get-Command pdftotext -ErrorAction SilentlyContinue
  $PdfToText = if ($c) { $c.Source } else { @("$env:LOCALAPPDATA\Programs\Git\ucrt64\bin\pdftotext.exe", "$env:ProgramFiles\Git\ucrt64\bin\pdftotext.exe", "$env:ProgramFiles\Git\mingw64\bin\pdftotext.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1 }
}
if (-not $PdfToText -or -not (Test-Path $PdfToText)) { throw 'pdftotext not found. Install Git for Windows or pass -PdfToText.' }
if (-not (Test-Path $Edge)) { throw "Edge not found at $Edge" }
$ReferencePdf = (Resolve-Path $ReferencePdf).Path

# FICTIONAL test values only.
$FILL = @{ provider = 'Dr. Avery Example'; patient = 'Testpatient Sample'; reason = 'Major Depressive Disorder (MDD)' }

# ---------- helpers ----------
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @"
using System; using System.Collections.Generic; using System.Drawing; using System.Drawing.Imaging; using System.Runtime.InteropServices;
public class CBand { public int Y0, Y1, X0, X1; }
public static class CImg {
  static byte[] Read(string p, out int w, out int h, out int s) { using (var bm = new Bitmap(p)) { w = bm.Width; h = bm.Height;
    var d = bm.LockBits(new Rectangle(0,0,w,h), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb); s = d.Stride;
    var b = new byte[s*h]; Marshal.Copy(d.Scan0, b, 0, b.Length); bm.UnlockBits(d); return b; } }
  public static List<CBand> Bands(string p, int t) { int w,h,s; var B = Read(p, out w, out h, out s); var res = new List<CBand>(); CBand cur = null;
    for (int y=0;y<h;y++){ int mn=-1,mx=-1; for(int x=0;x<w;x++){ int i=y*s+x*3; if((B[i]+B[i+1]+B[i+2])/3<t){ if(mn<0)mn=x; mx=x; } }
      if (mn>=0){ if(cur==null) cur=new CBand{Y0=y,Y1=y,X0=mn,X1=mx}; else { cur.Y1=y; if(mn<cur.X0)cur.X0=mn; if(mx>cur.X1)cur.X1=mx; } }
      else if (cur!=null){ res.Add(cur); cur=null; } }
    return res; }
  public static void Overlay(string a, string b, string o, int t) { int w,h,s,w2,h2,s2; var A = Read(a, out w, out h, out s); var B = Read(b, out w2, out h2, out s2);
    using (var bm = new Bitmap(w, h, PixelFormat.Format24bppRgb)) { var d = bm.LockBits(new Rectangle(0,0,w,h), ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb); var O = new byte[d.Stride*h];
      for (int y=0;y<h;y++) for (int x=0;x<w;x++) { int i=y*s+x*3, j=y*s2+x*3, k=y*d.Stride+x*3;
        bool da=(A[i]+A[i+1]+A[i+2])/3<t, db=y<h2&&x<w2&&(B[j]+B[j+1]+B[j+2])/3<t; byte r=255,g=255,bl=255;
        if(da&&db){r=g=bl=0;} else if(da){r=230;g=40;bl=40;} else if(db){r=20;g=120;bl=230;} O[k]=bl;O[k+1]=g;O[k+2]=r; }
      Marshal.Copy(O,0,d.Scan0,O.Length); bm.UnlockBits(d); bm.Save(o, ImageFormat.Png); } }
}
"@

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$ext = [System.WindowsRuntimeSystemExtensions].GetMethods()
$asOp = ($ext | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
$asAct = ($ext | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
function Await($op, [Type]$t) { $task = $asOp.MakeGenericMethod($t).Invoke($null, @($op)); $task.Wait(); $task.Result }
function PdfToPng([string]$pdf, [string]$prefix) {
  $f = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdf)) ([Windows.Storage.StorageFile])
  $doc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($f)) ([Windows.Data.Pdf.PdfDocument])
  $dir = Await ([Windows.Storage.StorageFolder]::GetFolderFromPathAsync($out)) ([Windows.Storage.StorageFolder])
  $pages = @()
  for ($i = 0; $i -lt $doc.PageCount; $i++) {
    $pg = $doc.GetPage($i); $name = "$prefix-$($i + 1).png"
    $o = Await ($dir.CreateFileAsync($name, [Windows.Storage.CreationCollisionOption]::ReplaceExisting)) ([Windows.Storage.StorageFile])
    $s = Await ($o.OpenAsync([Windows.Storage.FileAccessMode]::ReadWrite)) ([Windows.Storage.Streams.IRandomAccessStream])
    $opt = New-Object Windows.Data.Pdf.PdfPageRenderOptions; $opt.DestinationWidth = 1275
    $asAct.Invoke($null, @($pg.RenderToStreamAsync($s, $opt))).Wait(); $s.Dispose()
    $pages += [pscustomobject]@{ Png = (Join-Path $out $name); W = $pg.Size.Width; H = $pg.Size.Height }; $pg.Dispose()
  }
  $pages
}
function PageTexts([string]$pdf) {
  $txt = & $PdfToText -enc UTF-8 $pdf - | Out-String
  @($txt -split "`f" | Where-Object { $_.Trim() })
}
# words only: bullets, sub-bullet marks and underscores dropped, all whitespace removed
function Norm([string]$s) { ($s -replace '[\u2022\u00B7\uF0B7\uFFFD_]', '' -replace '(?m)^\s*o\s+', '' -replace '\s+', '') }
function RenderForm([hashtable]$vals, [string]$name) {
  $json = '{' + (($vals.GetEnumerator() | ForEach-Object { '"{0}":"{1}"' -f $_.Key, $_.Value }) -join ',') + '}'
  $html = (Get-Content (Join-Path $root 'consent.html') -Raw -Encoding UTF8) -replace '<script src="consent.js"></script>', "<script>window.__consentTest=$json</script><script src=`"consent.js`"></script>"
  $tmp = Join-Path $root "consent.$name.test.html"
  [IO.File]::WriteAllText($tmp, $html, (New-Object Text.UTF8Encoding $false))
  $pdf = Join-Path $out "$name.pdf"; if (Test-Path $pdf) { Remove-Item $pdf }
  $url = 'file:///' + ($tmp -replace '\\', '/')
  Start-Process $Edge -ArgumentList '--headless=new', '--disable-gpu', '--no-pdf-header-footer', '--virtual-time-budget=5000', "--print-to-pdf=$pdf", $url -Wait -WindowStyle Hidden
  Remove-Item $tmp
  if (-not (Test-Path $pdf)) { throw "Edge did not produce $pdf" }
  $pdf
}

$fail = @()
function Check([bool]$ok, [string]$what) { if ($ok) { Write-Host "  PASS  $what" -ForegroundColor Green } else { Write-Host "  FAIL  $what" -ForegroundColor Red; $script:fail += $what } }

Write-Host "Reference: $ReferencePdf"
$ref = PdfToPng $ReferencePdf 'ref'; $refText = PageTexts $ReferencePdf
$blank = PdfToPng (RenderForm @{ provider = ''; patient = ''; reason = '' } 'blank') 'blank'
$filledPdf = RenderForm $FILL 'filled'; $filled = PdfToPng $filledPdf 'filled'; $filledText = PageTexts $filledPdf

Write-Host "`n1. Pages"
Check ($ref.Count -eq 4) "reference has 4 pages ($($ref.Count))"
Check ($blank.Count -eq $ref.Count) "blank printout has $($blank.Count) pages"
Check ($filled.Count -eq $ref.Count) "filled printout has $($filled.Count) pages"
Check (@($filled | Where-Object { [Math]::Abs($_.W - 816) -gt 1 -or [Math]::Abs($_.H - 1056) -gt 1 }).Count -eq 0) 'every page is Letter (8.5 x 11 in)'

Write-Host "`n2. Text, page by page"
for ($i = 0; $i -lt $refText.Count; $i++) {
  $want = Norm $refText[$i]; $got = Norm $filledText[$i]
  if ($i -eq 0) {
    foreach ($k in 'provider', 'patient', 'reason') { $v = Norm $FILL[$k]; Check ($got.Contains($v)) "page 1 shows the $k value"; $got = $got.Replace($v, '') }
  }
  $ok = $want -eq $got
  if (-not $ok) { $n = 0; while ($n -lt [Math]::Min($want.Length, $got.Length) -and $want[$n] -eq $got[$n]) { $n++ }
    Write-Host ("        differs at: ref '...{0}' vs printout '...{1}'" -f $want.Substring([Math]::Max(0, $n - 30), [Math]::Min(60, $want.Length - [Math]::Max(0, $n - 30))), $got.Substring([Math]::Max(0, $n - 30), [Math]::Min(60, $got.Length - [Math]::Max(0, $n - 30)))) }
  Check $ok "page $($i + 1) wording matches the PDF"
}

Write-Host "`n3. Layout (blank printout vs reference, tolerance $Tol px at 150 dpi)"
for ($i = 0; $i -lt [Math]::Min($ref.Count, $blank.Count); $i++) {
  [CImg]::Overlay($ref[$i].Png, $blank[$i].Png, (Join-Path $out "overlay-$($i + 1).png"), 170)
  # the logo's light-grey tagline sits on the threshold; compare it by its top-left only
  $a = [CImg]::Bands($ref[$i].Png, 160); $b = [CImg]::Bands($blank[$i].Png, 160); $worst = 0; $bad = @()
  if ($a.Count -ne $b.Count) { Check $false "page $($i + 1): $($b.Count) printed lines vs $($a.Count) in the PDF"; continue }
  for ($j = 0; $j -lt $a.Count; $j++) {
    $p = $a[$j]; $q = $b[$j]; $logo = $p.Y1 -lt 140 -and $p.X0 -lt 300
    $m = [Math]::Max([Math]::Abs($q.Y0 - $p.Y0), [Math]::Abs($q.X0 - $p.X0)); if (-not $logo) { $m = [Math]::Max($m, [Math]::Abs($q.X1 - $p.X1)) }
    if ($m -gt $worst) { $worst = $m }; if ($m -gt $Tol) { $bad += "y$($p.Y0)" }
  }
  Check ($bad.Count -eq 0) ("page $($i + 1): $($a.Count) lines, worst offset $worst px" + $(if ($bad) { " (off: $($bad -join ', '))" } else { '' }))
}

Write-Host "`nOutput: $out (overlay-N.png: red = PDF only, blue = printout only, black = both)"
if ($fail.Count) { Write-Host "`n$($fail.Count) check(s) failed." -ForegroundColor Red; exit 1 } else { Write-Host "`nAll checks passed." -ForegroundColor Green; exit 0 }
