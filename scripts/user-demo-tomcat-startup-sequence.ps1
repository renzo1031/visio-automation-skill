param(
  [string] $OutputDirectory = '',
  [switch] $Invisible
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'skills\visio-automation\scripts\visio_helpers.ps1')

if (-not $OutputDirectory) {
  $OutputDirectory = Join-Path $repoRoot 'out\tomcat-startup-sequence'
}
$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

$outputPath = Join-Path $OutputDirectory 'tomcat-startup-sequence.vsdx'
$previewPath = Join-Path $OutputDirectory 'tomcat-startup-sequence.png'
if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }
if (Test-Path -LiteralPath $previewPath) { Remove-Item -LiteralPath $previewPath -Force }

function Set-ShapeNoFillNoLine {
  param([Parameter(Mandatory = $true)] $Shape)
  $null = $Shape.CellsU('FillPattern').FormulaU = '0'
  $null = $Shape.CellsU('LinePattern').FormulaU = '0'
}

function Set-TextStyle {
  param(
    [Parameter(Mandatory = $true)] $Shape,
    [double] $Size = 9,
    [int] $Bold = 0,
    [string] $Color = 'RGB(0,0,0)'
  )
  Set-VisioTextStyle -Shape $Shape -Size $Size -Bold $Bold -Color $Color
}

function Add-Box {
  param(
    [Parameter(Mandatory = $true)] $Page,
    [Parameter(Mandatory = $true)] [string] $Text,
    [double] $X,
    [double] $Y,
    [double] $Width,
    [double] $Height,
    [string] $Fill,
    [string] $Line = 'RGB(0,0,0)',
    [double] $TextSize = 8.5,
    [int] $Bold = 1
  )
  $shape = $Page.DrawRectangle($X - ($Width / 2), $Y - ($Height / 2), $X + ($Width / 2), $Y + ($Height / 2))
  $shape.Text = $Text
  Set-VisioShapeFill -Shape $shape -Fill $Fill -Line $Line -LineWeight '1 pt'
  Set-TextStyle -Shape $shape -Size $TextSize -Bold $Bold
  return $shape
}

function Add-LifeLine {
  param(
    [Parameter(Mandatory = $true)] $Page,
    [double] $X,
    [double] $TopY,
    [double] $BottomY
  )
  $line = $Page.DrawLine($X, $TopY, $X, $BottomY)
  $null = $line.CellsU('LineColor').FormulaU = 'RGB(0,0,0)'
  $null = $line.CellsU('LineWeight').FormulaU = '0.75 pt'
  $null = $line.CellsU('LinePattern').FormulaU = '2'
  return $line
}

function Add-Activation {
  param(
    [Parameter(Mandatory = $true)] $Page,
    [double] $X,
    [double] $TopY,
    [double] $BottomY
  )
  $width = 0.12
  $shape = $Page.DrawRectangle($X - ($width / 2), $BottomY, $X + ($width / 2), $TopY)
  Set-VisioShapeFill -Shape $shape -Fill 'RGB(255,255,255)' -Line 'RGB(0,0,0)' -LineWeight '0.9 pt'
  return $shape
}

function Add-Message {
  param(
    [Parameter(Mandatory = $true)] $Page,
    [string] $Text,
    [double] $FromX,
    [double] $ToX,
    [double] $Y,
    [switch] $Dashed,
    [switch] $NoArrow
  )
  $line = $Page.DrawLine($FromX, $Y, $ToX, $Y)
  $null = $line.CellsU('LineColor').FormulaU = 'RGB(0,0,0)'
  $null = $line.CellsU('LineWeight').FormulaU = '0.9 pt'
  if ($Dashed) {
    $null = $line.CellsU('LinePattern').FormulaU = '2'
  }
  if (-not $NoArrow) {
    $null = $line.CellsU('EndArrow').FormulaU = '4'
  }
  $labelX = ($FromX + $ToX) / 2
  $labelY = $Y + 0.12
  $labelWidth = [Math]::Max(0.45, [Math]::Min(1.15, 0.12 * $Text.Length + 0.18))
  $label = Add-VisioLabel -Page $Page -Text $Text -X $labelX -Y $labelY -Width $labelWidth -Height 0.18
  Set-TextStyle -Shape $label -Size 8.3 -Bold 1
  return [pscustomobject]@{ Line = $line; Label = $label }
}

function Add-SelfMessage {
  param(
    [Parameter(Mandatory = $true)] $Page,
    [string] $Text,
    [double] $X,
    [double] $Y,
    [double] $Width = 0.34,
    [double] $Height = 0.22,
    [switch] $Left
  )
  if ($Left) {
    $x2 = $X - $Width
    $l1 = $Page.DrawLine($X, $Y, $x2, $Y)
    $l2 = $Page.DrawLine($x2, $Y, $x2, $Y - $Height)
    $l3 = $Page.DrawLine($x2, $Y - $Height, $X - 0.02, $Y - $Height)
  } else {
    $x2 = $X + $Width
    $l1 = $Page.DrawLine($X, $Y, $x2, $Y)
    $l2 = $Page.DrawLine($x2, $Y, $x2, $Y - $Height)
    $l3 = $Page.DrawLine($x2, $Y - $Height, $X + 0.02, $Y - $Height)
  }
  foreach ($line in @($l1, $l2, $l3)) {
    $null = $line.CellsU('LineColor').FormulaU = 'RGB(0,0,0)'
    $null = $line.CellsU('LineWeight').FormulaU = '0.9 pt'
  }
  $null = $l3.CellsU('EndArrow').FormulaU = '4'
  $labelX = if ($Left) { $X - ($Width / 2) } else { $X + ($Width / 2) }
  $label = Add-VisioLabel -Page $Page -Text $Text -X $labelX -Y ($Y + 0.12) -Width 0.52 -Height 0.18
  Set-TextStyle -Shape $label -Size 8.3 -Bold 1
  return @($l1, $l2, $l3, $label)
}

function Add-ActorIcon {
  param(
    [Parameter(Mandatory = $true)] $Page,
    [double] $X,
    [double] $Y
  )
  $fill = 'RGB(139,195,74)'
  $lineColor = 'RGB(33,84,24)'
  $head = $Page.DrawOval($X - 0.04, $Y + 0.16, $X + 0.04, $Y + 0.24)
  $body = $Page.DrawRectangle($X - 0.06, $Y - 0.05, $X + 0.06, $Y + 0.14)
  $leftArm = $Page.DrawLine($X - 0.06, $Y + 0.10, $X - 0.13, $Y - 0.02)
  $rightArm = $Page.DrawLine($X + 0.06, $Y + 0.10, $X + 0.13, $Y - 0.02)
  $leftLeg = $Page.DrawLine($X - 0.03, $Y - 0.05, $X - 0.08, $Y - 0.18)
  $rightLeg = $Page.DrawLine($X + 0.03, $Y - 0.05, $X + 0.08, $Y - 0.18)
  foreach ($shape in @($head, $body)) {
    Set-VisioShapeFill -Shape $shape -Fill $fill -Line $lineColor -LineWeight '1 pt'
  }
  foreach ($line in @($leftArm, $rightArm, $leftLeg, $rightLeg)) {
    $null = $line.CellsU('LineColor').FormulaU = $lineColor
    $null = $line.CellsU('LineWeight').FormulaU = '1.2 pt'
  }
}

$visio = if ($Invisible) { New-InvisibleVisioApplication } else { New-VisibleVisioApplication }
$doc = $null

try {
  $doc = $visio.Documents.Add('')
  $page = $visio.ActivePage
  $pageWidth = 17.5
  $pageHeight = 7.5
  $null = $page.PageSheet.CellsU('PageWidth').FormulaU = "$pageWidth in"
  $null = $page.PageSheet.CellsU('PageHeight').FormulaU = "$pageHeight in"

  $participants = @(
    [pscustomobject]@{ id = 'startup'; label = 'startup.bat'; x = 0.65; width = 1.25; fill = 'RGB(139,195,74)' },
    [pscustomobject]@{ id = 'bootstrap'; label = 'BootStrap'; x = 2.25; width = 0.92; fill = 'RGB(194,225,236)' },
    [pscustomobject]@{ id = 'catalina'; label = 'Catalina'; x = 3.75; width = 0.92; fill = 'RGB(194,225,236)' },
    [pscustomobject]@{ id = 'server'; label = 'Server'; x = 5.25; width = 0.92; fill = 'RGB(194,225,236)' },
    [pscustomobject]@{ id = 'service'; label = 'Service'; x = 6.80; width = 0.92; fill = 'RGB(194,225,236)' },
    [pscustomobject]@{ id = 'executor'; label = 'Executor'; x = 8.45; width = 0.92; fill = 'RGB(253,222,196)' },
    [pscustomobject]@{ id = 'engine'; label = 'Engine'; x = 9.95; width = 0.92; fill = 'RGB(254,242,198)' },
    [pscustomobject]@{ id = 'host'; label = 'Host'; x = 11.45; width = 0.92; fill = 'RGB(254,242,198)' },
    [pscustomobject]@{ id = 'context'; label = 'Context'; x = 12.85; width = 0.92; fill = 'RGB(254,242,198)' },
    [pscustomobject]@{ id = 'connector'; label = 'Connector'; x = 14.35; width = 0.92; fill = 'RGB(253,222,196)' },
    [pscustomobject]@{ id = 'protocol'; label = 'ProtocolHandler'; x = 16.45; width = 1.52; fill = 'RGB(253,222,196)' }
  )

  $x = @{}
  foreach ($participant in $participants) {
    $x[$participant.id] = [double] $participant.x
    Add-Box -Page $page -Text $participant.label -X $participant.x -Y 6.75 -Width $participant.width -Height 0.36 -Fill $participant.fill -TextSize 8.5 -Bold 1 | Out-Null
    Add-LifeLine -Page $page -X $participant.x -TopY 6.55 -BottomY 0.65 | Out-Null
  }
  Add-ActorIcon -Page $page -X $x.startup -Y 7.18

  Add-Activation -Page $page -X $x.startup -TopY 6.35 -BottomY 5.70 | Out-Null
  Add-Activation -Page $page -X $x.bootstrap -TopY 6.35 -BottomY 0.78 | Out-Null
  Add-Activation -Page $page -X ($x.bootstrap + 0.08) -TopY 5.20 -BottomY 4.58 | Out-Null
  Add-Activation -Page $page -X ($x.bootstrap + 0.15) -TopY 3.34 -BottomY 2.95 | Out-Null
  Add-Activation -Page $page -X ($x.bootstrap + 0.07) -TopY 2.72 -BottomY 2.18 | Out-Null

  Add-Activation -Page $page -X $x.catalina -TopY 5.28 -BottomY 4.40 | Out-Null
  Add-Activation -Page $page -X ($x.catalina + 0.18) -TopY 4.88 -BottomY 4.55 | Out-Null
  Add-Activation -Page $page -X $x.catalina -TopY 2.83 -BottomY 2.20 | Out-Null
  Add-Activation -Page $page -X $x.server -TopY 4.80 -BottomY 4.00 | Out-Null
  Add-Activation -Page $page -X $x.server -TopY 2.75 -BottomY 2.08 | Out-Null
  Add-Activation -Page $page -X $x.service -TopY 4.50 -BottomY 3.12 | Out-Null
  Add-Activation -Page $page -X $x.service -TopY 2.56 -BottomY 1.47 | Out-Null
  Add-Activation -Page $page -X $x.executor -TopY 4.32 -BottomY 4.02 | Out-Null
  Add-Activation -Page $page -X $x.executor -TopY 2.15 -BottomY 1.83 | Out-Null
  Add-Activation -Page $page -X $x.engine -TopY 4.30 -BottomY 3.92 | Out-Null
  Add-Activation -Page $page -X $x.engine -TopY 2.60 -BottomY 2.28 | Out-Null
  Add-Activation -Page $page -X $x.host -TopY 4.20 -BottomY 3.65 | Out-Null
  Add-Activation -Page $page -X $x.host -TopY 2.42 -BottomY 2.08 | Out-Null
  Add-Activation -Page $page -X $x.context -TopY 4.00 -BottomY 3.66 | Out-Null
  Add-Activation -Page $page -X $x.context -TopY 2.32 -BottomY 1.95 | Out-Null
  Add-Activation -Page $page -X $x.connector -TopY 3.55 -BottomY 2.87 | Out-Null
  Add-Activation -Page $page -X $x.connector -TopY 1.88 -BottomY 1.43 | Out-Null
  Add-Activation -Page $page -X $x.protocol -TopY 3.40 -BottomY 2.95 | Out-Null
  Add-Activation -Page $page -X $x.protocol -TopY 1.78 -BottomY 1.35 | Out-Null

  Add-Message -Page $page -Text 'main()' -FromX ($x.startup + 0.06) -ToX ($x.bootstrap - 0.06) -Y 6.10 | Out-Null
  Add-SelfMessage -Page $page -Text 'init' -X ($x.bootstrap + 0.06) -Y 5.86 -Width 0.28 -Height 0.18 | Out-Null
  Add-SelfMessage -Page $page -Text 'load' -X ($x.bootstrap + 0.06) -Y 5.60 -Width 0.28 -Height 0.18 | Out-Null
  Add-Message -Page $page -Text 'load' -FromX ($x.bootstrap + 0.12) -ToX ($x.catalina - 0.06) -Y 5.30 | Out-Null
  Add-SelfMessage -Page $page -Text '创建Server' -X ($x.catalina + 0.06) -Y 5.07 -Width 0.30 -Height 0.30 | Out-Null

  Add-Message -Page $page -Text 'init' -FromX ($x.catalina + 0.06) -ToX ($x.server - 0.06) -Y 4.42 | Out-Null
  Add-Message -Page $page -Text 'init' -FromX ($x.server + 0.06) -ToX ($x.service - 0.06) -Y 4.36 | Out-Null
  Add-Message -Page $page -Text 'init' -FromX ($x.service + 0.06) -ToX ($x.engine - 0.06) -Y 4.18 | Out-Null
  Add-Message -Page $page -Text 'init' -FromX ($x.engine + 0.06) -ToX ($x.host - 0.06) -Y 4.05 | Out-Null
  Add-Message -Page $page -Text 'init' -FromX ($x.host + 0.06) -ToX ($x.context - 0.06) -Y 3.86 | Out-Null
  Add-Message -Page $page -Text 'init' -FromX ($x.service + 0.06) -ToX ($x.executor - 0.06) -Y 3.72 | Out-Null
  Add-Message -Page $page -Text 'init' -FromX ($x.service + 0.06) -ToX ($x.connector - 0.06) -Y 3.34 | Out-Null
  Add-Message -Page $page -Text 'init' -FromX ($x.connector + 0.06) -ToX ($x.protocol - 0.06) -Y 3.12 | Out-Null

  Add-SelfMessage -Page $page -Text 'start' -X ($x.bootstrap + 0.06) -Y 3.12 -Width 0.28 -Height 0.22 -Left | Out-Null
  Add-Message -Page $page -Text 'start' -FromX ($x.bootstrap + 0.08) -ToX ($x.catalina - 0.06) -Y 2.72 | Out-Null
  Add-Message -Page $page -Text 'start' -FromX ($x.catalina + 0.06) -ToX ($x.server - 0.06) -Y 2.58 | Out-Null
  Add-Message -Page $page -Text 'start' -FromX ($x.server + 0.06) -ToX ($x.service - 0.06) -Y 2.46 | Out-Null
  Add-Message -Page $page -Text 'start' -FromX ($x.service + 0.06) -ToX ($x.engine - 0.06) -Y 2.22 | Out-Null
  Add-Message -Page $page -Text 'start' -FromX ($x.engine + 0.06) -ToX ($x.host - 0.06) -Y 2.12 | Out-Null
  Add-Message -Page $page -Text 'start' -FromX ($x.host + 0.06) -ToX ($x.context - 0.06) -Y 2.02 | Out-Null
  Add-Message -Page $page -Text 'start' -FromX ($x.service + 0.06) -ToX ($x.executor - 0.06) -Y 1.84 | Out-Null
  Add-Message -Page $page -Text 'start' -FromX ($x.service + 0.06) -ToX ($x.connector - 0.06) -Y 1.55 | Out-Null
  Add-Message -Page $page -Text 'start' -FromX ($x.connector + 0.06) -ToX ($x.protocol - 0.06) -Y 1.46 | Out-Null

  $doc.SaveAs($outputPath)
  $page.Export($previewPath)

  Update-VisioSessionCache `
    -WorkspacePath $repoRoot `
    -DiagramType 'uml-sequence' `
    -Stencil 'USEQME_M.VSSX' `
    -LastOutputPath $outputPath `
    -LastPreviewPath $previewPath `
    -LastLayoutStrategy 'ManualSequence' `
    -LastLayoutDirection 'LR' | Out-Null

  [pscustomobject]@{
    OutputPath = $outputPath
    PreviewPath = $previewPath
    Participants = $participants.Count
    Messages = 22
    Visible = -not [bool] $Invisible
  } | ConvertTo-Json -Depth 4
} finally {
  if ($Invisible) {
    if ($null -ne $doc) {
      try { $doc.Close() } catch {}
    }
    if ($null -ne $visio) {
      try { $visio.Quit() } catch {}
    }
  }
}
