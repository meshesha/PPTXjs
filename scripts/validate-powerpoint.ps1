# Opens exported presentations in Microsoft PowerPoint (COM automation) to
# verify they are valid: each file is opened read-only, its slide count is
# printed, PNG previews of the first slides and a PDF are written next to it.
# Windows + PowerPoint only.
#
# usage: powershell -ExecutionPolicy Bypass -File scripts/validate-powerpoint.ps1 test/output/exports
param([string]$Dir = "test/output/exports")

$files = Get-ChildItem -Path $Dir -Include *.pptx, *.odp -Recurse | Where-Object { $_.Name -notlike "*.roundtrip.pptx" }
$wasRunning = $null -ne (Get-Process POWERPNT -ErrorAction SilentlyContinue)
$app = New-Object -ComObject PowerPoint.Application
$failed = 0
foreach ($f in $files) {
  try {
    # Open(FileName, ReadOnly, Untitled, WithWindow)
    $p = $app.Presentations.Open($f.FullName, -1, 0, 0)
    $count = $p.Slides.Count
    $base = Join-Path $f.DirectoryName ($f.Name -replace '\.', '_')
    $p.SaveAs("$base.pdf", 32)   # ppSaveAsPDF
    if ($f.Extension -eq ".odp") {
      # slide image export needs a native presentation: convert, then reopen
      $copy = "$base.roundtrip.pptx"
      $p.SaveCopyAs($copy, 24)   # ppSaveAsOpenXMLPresentation
      $p.Close()
      $p = $app.Presentations.Open($copy, -1, 0, 0)
    }
    for ($i = 1; $i -le [Math]::Min(3, $count); $i++) {
      $p.Slides.Item($i).Export("$base-slide$i.png", "PNG", 960, 540)
    }
    Write-Output ("OK    {0,-28} slides={1}" -f $f.Name, $count)
    $p.Close()
  } catch {
    $failed++
    Write-Output ("FAIL  {0,-28} {1}" -f $f.Name, $_.Exception.Message)
  }
}
if (-not $wasRunning) { $app.Quit() }
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
exit $failed
