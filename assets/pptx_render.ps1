# Render every slide of a .pptx twice with PowerPoint: once as-is (reference)
# and once with every text shape hidden (the "plate" — background + decoration
# only). The plate is what makes a template reusable: new copy can be dropped on
# top without the old copy showing through.
#
# Driven by assets/pptx_template.py, which owns the JSON contract. This script
# only renders and prints one JSON line.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File pptx_render.ps1 \
#       -Src <file.pptx> -OutDir <dir> [-Width 1600] [-Height 900] [-MaxSlides 0]
param(
  [Parameter(Mandatory = $true)][string]$Src,
  [Parameter(Mandatory = $true)][string]$OutDir,
  [int]$Width = 1600,
  [int]$Height = 900,
  [int]$MaxSlides = 0
)

$ErrorActionPreference = 'Stop'
$msoFalse = 0
$msoTrue = -1

function Write-Json($obj) {
  $obj | ConvertTo-Json -Compress -Depth 6 | Write-Output
}

# Hide (or restore) every text frame under a shape collection, descending into
# groups. Charts/tables keep their frame: hiding a chart would delete its
# graphic, and a template's charts are decoration we want to keep.
function Set-TextVisibility($shapes, $visible) {
  $touched = @()
  foreach ($shape in $shapes) {
    try {
      $type = [int]$shape.Type
      # 6 = msoGroup
      if ($type -eq 6) {
        $touched += Set-TextVisibility $shape.GroupItems $visible
        continue
      }
      # 3 = chart, 19 = table — keep the whole shape.
      if ($type -eq 3 -or $type -eq 19) { continue }
      if ($shape.HasTextFrame -eq $msoTrue -and $shape.TextFrame.HasText -eq $msoTrue) {
        if ($visible) { $shape.Visible = $msoTrue } else { $shape.Visible = $msoFalse }
        $touched += $shape.Name
      }
    } catch {
      # A shape we cannot inspect is left alone; rendering must not fail over it.
    }
  }
  return $touched
}

$app = $null
$pres = $null
try {
  New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
  $refDir = Join-Path $OutDir 'reference'
  $plateDir = Join-Path $OutDir 'plates'
  New-Item -ItemType Directory -Force -Path $refDir | Out-Null
  New-Item -ItemType Directory -Force -Path $plateDir | Out-Null

  $app = New-Object -ComObject PowerPoint.Application
  # ReadOnly + WithWindow=false so a headless render never steals the user's screen.
  $pres = $app.Presentations.Open($Src, $msoTrue, $msoFalse, $msoFalse)
  $count = $pres.Slides.Count
  $limit = $count
  if ($MaxSlides -gt 0 -and $MaxSlides -lt $count) { $limit = $MaxSlides }

  $pages = @()
  for ($i = 1; $i -le $limit; $i++) {
    $slide = $pres.Slides.Item($i)
    $ref = Join-Path $refDir ('s' + ($i - 1) + '.png')
    $slide.Export($ref, 'PNG', $Width, $Height)

    $hidden = @(Set-TextVisibility $slide.Shapes $false)
    $plate = Join-Path $plateDir ('s' + ($i - 1) + '.png')
    try {
      $slide.Export($plate, 'PNG', $Width, $Height)
    } finally {
      # Put the text back so the next slide's reference render is complete.
      [void](Set-TextVisibility $slide.Shapes $true)
    }
    $pages += @{ index = ($i - 1); hiddenTextShapes = $hidden.Count }
  }

  Write-Json ([ordered]@{
    ok        = $true
    engine    = 'powerpoint'
    slides    = $limit
    total     = $count
    width     = $Width
    height    = $Height
    outDir    = $OutDir
    pages     = $pages
  })
  exit 0
} catch {
  Write-Json ([ordered]@{
    ok      = $false
    engine  = 'powerpoint'
    error   = $_.Exception.Message
  })
  exit 1
} finally {
  if ($pres -ne $null) { try { $pres.Close() } catch {} }
  if ($app -ne $null) { try { $app.Quit() } catch {} }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
