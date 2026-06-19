param(
  [string] $OutputDirectory = (Join-Path (Split-Path -Parent $PSScriptRoot) 'out\visio-business-regression'),
  [switch] $NoVisio,
  [switch] $Visible,
  [switch] $ApproveGolden,
  [double] $ImageDiffThreshold = 0.05,
  [int] $ImageDiffMaxPixelDelta = 30
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'skills\visio-automation\scripts\visio_helpers.ps1')

$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

$goldenRoot = Join-Path $repoRoot 'tests\golden-images'
$expectedRoot = Join-Path $goldenRoot 'expected'
$actualRoot = Join-Path $goldenRoot 'actual'
New-Item -ItemType Directory -Path $expectedRoot -Force | Out-Null
New-Item -ItemType Directory -Path $actualRoot -Force | Out-Null

function Get-FileSha256 {
  param([Parameter(Mandatory = $true)] [string] $Path)
  if (-not (Test-Path -LiteralPath $Path)) { return $null }
  try {
    return (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash
  } catch {
    return $null
  }
}

function Compare-VisioPngPixel {
  param(
    [Parameter(Mandatory = $true)] [string] $ExpectedPath,
    [Parameter(Mandatory = $true)] [string] $ActualPath,
    [double] $Threshold = 0.05,
    [int] $MaxPixelDelta = 30
  )

  if (-not (Test-Path -LiteralPath $ExpectedPath) -or -not (Test-Path -LiteralPath $ActualPath)) {
    return [pscustomobject]@{
      canCompare = $false
      reason = 'Missing file(s)'
      match = $false
      differentPixelRatio = 1.0
      maxChannelDelta = 255
    }
  }

  try {
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
  } catch {
    return [pscustomobject]@{
      canCompare = $false
      reason = 'System.Drawing not available'
      match = $false
      differentPixelRatio = 1.0
      maxChannelDelta = 255
    }
  }

  try {
    $expectedImg = [System.Drawing.Bitmap]::new($ExpectedPath)
    $actualImg = [System.Drawing.Bitmap]::new($ActualPath)

    if ($expectedImg.Width -ne $actualImg.Width -or $expectedImg.Height -ne $actualImg.Height) {
      $expectedImg.Dispose()
      $actualImg.Dispose()
      return [pscustomobject]@{
        canCompare = $true
        reason = 'Size mismatch'
        match = $false
        differentPixelRatio = 1.0
        maxChannelDelta = 255
        expectedSize = "$($expectedImg.Width)x$($expectedImg.Height)"
        actualSize = "$($actualImg.Width)x$($actualImg.Height)"
      }
    }

    $totalPixels = $expectedImg.Width * $expectedImg.Height
    $differentPixels = 0
    $worstChannelDelta = 0

    for ($y = 0; $y -lt $expectedImg.Height; $y++) {
      for ($x = 0; $x -lt $expectedImg.Width; $x++) {
        $ePx = $expectedImg.GetPixel($x, $y)
        $aPx = $actualImg.GetPixel($x, $y)

        $dr = [Math]::Abs([int] $ePx.R - [int] $aPx.R)
        $dg = [Math]::Abs([int] $ePx.G - [int] $aPx.G)
        $db = [Math]::Abs([int] $ePx.B - [int] $aPx.B)
        $da = [Math]::Abs([int] $ePx.A - [int] $aPx.A)
        $delta = [Math]::Max([Math]::Max($dr, $dg), [Math]::Max($db, $da))

        if ($delta -gt $worstChannelDelta) {
          $worstChannelDelta = $delta
        }

        if ($delta -gt $MaxPixelDelta) {
          $differentPixels++
        }
      }
    }

    $expectedImg.Dispose()
    $actualImg.Dispose()

    $ratio = if ($totalPixels -gt 0) { [double] $differentPixels / [double] $totalPixels } else { 0.0 }
    $isMatch = $ratio -le $Threshold

    return [pscustomobject]@{
      canCompare = $true
      reason = 'Pixel comparison complete'
      match = $isMatch
      differentPixelRatio = [math]::Round($ratio, 6)
      maxChannelDelta = $worstChannelDelta
      differentPixels = $differentPixels
      totalPixels = $totalPixels
    }
  } catch {
    return [pscustomobject]@{
      canCompare = $false
      reason = "Pixel comparison error: $($_.Exception.Message)"
      match = $false
      differentPixelRatio = 1.0
      maxChannelDelta = 255
    }
  }
}

function Assert-VisioDocumentStructure {
  param(
    [Parameter(Mandatory = $true)] [string] $VsdxPath,
    [Parameter(Mandatory = $true)] [int] $ExpectedNodeCount,
    [Parameter(Mandatory = $true)] [int] $ExpectedEdgeCount,
    [string[]] $ExpectedRouteTypes = @()
  )

  $visio = $null
  $doc = $null
  $structuralErrors = @()

  try {
    $visio = New-InvisibleVisioApplication
    $doc = $visio.Documents.OpenEx($VsdxPath, 64 + 2)
    $page = $visio.ActivePage

    $allShapes = @($page.Shapes)
    $connectors = @($allShapes | Where-Object { $_.OneD -ne 0 })
    $nonConnectors = @($allShapes | Where-Object { $_.OneD -eq 0 })

    if ($nonConnectors.Count -ne $ExpectedNodeCount) {
      $structuralErrors += "Node count mismatch: expected $ExpectedNodeCount, got $($nonConnectors.Count)"
    }

    if ($connectors.Count -ne $ExpectedEdgeCount) {
      $structuralErrors += "Edge count mismatch: expected $ExpectedEdgeCount, got $($connectors.Count)"
    }

    $connectorChecks = foreach ($connector in $connectors) {
      $beginX = $connector.CellsU('BeginX').FormulaU
      $endX = $connector.CellsU('EndX').FormulaU
      $isGlued = ($beginX -match 'Glue|GUARD|Connections|Pin') -and ($endX -match 'Glue|GUARD|Connections|Pin')

      $routeType = 'orthogonal'
      $routeStyle = if ($connector.CellExistsU('ShapeRouteStyle', 0) -ne 0) { [int] $connector.CellsU('ShapeRouteStyle').ResultIU } else { 0 }
      $routeExt = if ($connector.CellExistsU('ConLineRouteExt', 0) -ne 0) { [int] $connector.CellsU('ConLineRouteExt').ResultIU } else { 0 }

      if ($routeStyle -eq 2 -and $routeExt -eq 1) {
        $routeType = 'straight'
      } elseif ($routeExt -eq 2) {
        $routeType = 'curved'
      }

      if (-not $isGlued) {
        $structuralErrors += "Connector '$($connector.NameU)' endpoints are not glued (BeginX=$beginX, EndX=$endX)"
      }

      [pscustomobject]@{
        nameU = $connector.NameU
        oneD = [int] $connector.OneD
        isGlued = $isGlued
        routeType = $routeType
        shapeRouteStyle = $routeStyle
        conLineRouteExt = $routeExt
      }
    }

    if ($ExpectedRouteTypes.Count -gt 0) {
      foreach ($expected in $ExpectedRouteTypes) {
        $found = @($connectorChecks | Where-Object { $_.routeType -eq $expected }).Count
        if ($found -eq 0) {
          $structuralErrors += "Expected route type '$expected' not found among connectors"
        }
      }
    }

    return [pscustomobject]@{
      passed = ($structuralErrors.Count -eq 0)
      errors = $structuralErrors
      totalShapes = $allShapes.Count
      nodeCount = $nonConnectors.Count
      edgeCount = $connectors.Count
      connectors = @($connectorChecks)
    }
  } catch {
    return [pscustomobject]@{
      passed = $false
      errors = @("Structural verification error: $($_.Exception.Message)")
      totalShapes = 0
      nodeCount = 0
      edgeCount = 0
      connectors = @()
    }
  } finally {
    if ($null -ne $doc) {
      try { $doc.Saved = $true; $doc.Close() } catch {}
    }
    if ($null -ne $visio) {
      try { $visio.Quit() } catch {}
    }
  }
}

function New-BusinessReportEntry {
  param(
    [Parameter(Mandatory = $true)] [string] $Name,
    [Parameter(Mandatory = $true)] [string] $Kind,
    [Parameter(Mandatory = $true)] [string] $OutputPath,
    [Parameter(Mandatory = $true)] [string] $PreviewPath,
    [Parameter(Mandatory = $true)] $Result,
    [string] $ExpectedPreviewPath = '',
    $StructureCheck = $null,
    $ImageDiff = $null
  )

  $entry = [ordered]@{
    name = $Name
    kind = $Kind
    outputPath = $OutputPath
    previewPath = $PreviewPath
    outputExists = Test-Path -LiteralPath $OutputPath
    previewExists = Test-Path -LiteralPath $PreviewPath
    outputSha256 = Get-FileSha256 -Path $OutputPath
    previewSha256 = Get-FileSha256 -Path $PreviewPath
    nodes = @($Result.nodes).Count
    edges = @($Result.edges).Count
    routeTypes = @($Result.edges | ForEach-Object { $_.routeType })
    visibleMode = $true
    expectedPreviewPath = $ExpectedPreviewPath
    goldenHashMatch = $null
    structuralCheck = $null
    imageDiff = $null
    overallPassed = $false
    failureCategory = ''
  }

  if ($null -ne $StructureCheck) {
    $entry.structuralCheck = [ordered]@{
      passed = $StructureCheck.passed
      nodeCount = $StructureCheck.nodeCount
      edgeCount = $StructureCheck.edgeCount
      errors = @($StructureCheck.errors)
    }
  }

  if ($null -ne $ImageDiff) {
    $entry.imageDiff = [ordered]@{
      canCompare = $ImageDiff.canCompare
      match = $ImageDiff.match
      differentPixelRatio = $ImageDiff.differentPixelRatio
      maxChannelDelta = $ImageDiff.maxChannelDelta
      reason = $ImageDiff.reason
    }
  }

  if ($ExpectedPreviewPath -and (Test-Path -LiteralPath $ExpectedPreviewPath) -and (Test-Path -LiteralPath $PreviewPath)) {
    $expectedHash = Get-FileSha256 -Path $ExpectedPreviewPath
    $entry.goldenHashMatch = ($expectedHash -eq $entry.previewSha256)
  }

  $structuralPass = if ($null -ne $StructureCheck) { $StructureCheck.passed } else { $true }
  $visualPass = if ($null -ne $ImageDiff) { $ImageDiff.match } else { $true }
  $outputPass = $entry.outputExists -and $entry.previewExists

  $entry.overallPassed = $structuralPass -and $visualPass -and $outputPass

  if (-not $outputPass) {
    $entry.failureCategory = 'output-missing'
  } elseif (-not $structuralPass) {
    $entry.failureCategory = 'structural'
  } elseif (-not $visualPass) {
    $entry.failureCategory = 'visual'
  }

  return [pscustomobject]$entry
}

function Save-BusinessCase {
  param(
    [Parameter(Mandatory = $true)] [string] $Name,
    [Parameter(Mandatory = $true)] [string] $Kind,
    [Parameter(Mandatory = $true)] $Graph,
    [Parameter(Mandatory = $true)] [string] $FileName,
    [string] $LayoutStrategy = '',
    [string] $LayoutDirection = 'LR',
    [string[]] $ExpectedRouteTypes = @()
  )

  $outputPath = Join-Path $OutputDirectory $FileName
  $previewPath = [System.IO.Path]::ChangeExtension($outputPath, '.png')
  $expectedPreviewPath = Join-Path $expectedRoot ([System.IO.Path]::ChangeExtension($FileName, '.png'))
  $actualPreviewPath = Join-Path $actualRoot ([System.IO.Path]::ChangeExtension($FileName, '.png'))

  if ($NoVisio) {
    $dryRunResult = [pscustomobject]@{
      diagramType = $Graph.diagramType
      layout = if ($Graph.layout) { $Graph.layout } else { $null }
      outputPath = $outputPath
      previewPath = $previewPath
      nodes = @($Graph.nodes)
      edges = @($Graph.edges)
    }
    return New-BusinessReportEntry -Name $Name -Kind $Kind -OutputPath $outputPath -PreviewPath $previewPath -Result $dryRunResult -ExpectedPreviewPath $expectedPreviewPath
  }

  $openDocument = Get-OpenVisioDocumentByPath -Path $outputPath
  if ($null -ne $openDocument) {
    try { $openDocument.Saved = $true } catch {}
    try { $openDocument.Close() } catch {}
    Start-Sleep -Milliseconds 300
  }

  $renderArgs = @{
    GraphModel = $Graph
    OutputPath = $outputPath
    PreviewPath = $previewPath
    Force = $true
  }
  if ($Visible) {
    $renderArgs.Visible = $true
  } else {
    $renderArgs.Invisible = $true
  }
  if ($LayoutStrategy) {
    $renderArgs.LayoutStrategy = $LayoutStrategy
    $renderArgs.LayoutDirection = $LayoutDirection
  }

  $result = Render-VisioGraphModel @renderArgs
  if (Test-Path -LiteralPath $previewPath) {
    Copy-Item -LiteralPath $previewPath -Destination $actualPreviewPath -Force
  }

  $expectedNodeCount = @($Graph.nodes).Count
  $expectedEdgeCount = @($Graph.edges).Count
  $structureCheck = Assert-VisioDocumentStructure -VsdxPath $outputPath -ExpectedNodeCount $expectedNodeCount -ExpectedEdgeCount $expectedEdgeCount -ExpectedRouteTypes $ExpectedRouteTypes

  $imageDiff = $null
  if ((Test-Path -LiteralPath $expectedPreviewPath) -and (Test-Path -LiteralPath $previewPath)) {
    $imageDiff = Compare-VisioPngPixel -ExpectedPath $expectedPreviewPath -ActualPath $previewPath -Threshold $ImageDiffThreshold -MaxPixelDelta $ImageDiffMaxPixelDelta
  }

  $entry = New-BusinessReportEntry -Name $Name -Kind $Kind -OutputPath $outputPath -PreviewPath $previewPath -Result $result -ExpectedPreviewPath $expectedPreviewPath -StructureCheck $structureCheck -ImageDiff $imageDiff
  $entry.visibleMode = [bool] $Visible
  return $entry
}

$results = [ordered]@{
  outputDirectory = $OutputDirectory
  goldenRoot = $goldenRoot
  imageDiffThreshold = $ImageDiffThreshold
  imageDiffMaxPixelDelta = $ImageDiffMaxPixelDelta
  cases = @()
  files = @()
  structuralFailures = @()
  visualFailures = @()
}

# Case 1: approval flowchart.
$approvalGraph = [pscustomobject]@{
  diagramType = 'flowchart'
  page = [pscustomobject]@{ width = 9.5; height = 5.5 }
  nodes = @(
    [pscustomobject]@{ id = 'start'; text = '发起申请'; semanticType = 'start'; stencil = 'BASFLO_M.VSSX'; x = 1.0; y = 3.7; width = 1.4; height = 0.58 }
    [pscustomobject]@{ id = 'review'; text = '部门审批'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 3.0; y = 3.7; width = 1.45; height = 0.6 }
    [pscustomobject]@{ id = 'decision'; text = '是否通过?'; semanticType = 'decision'; stencil = 'BASFLO_M.VSSX'; x = 5.0; y = 3.7; width = 1.35; height = 0.82 }
    [pscustomobject]@{ id = 'end'; text = '结束'; semanticType = 'end'; stencil = 'BASFLO_M.VSSX'; x = 7.0; y = 3.7; width = 1.2; height = 0.58 }
    [pscustomobject]@{ id = 'reject'; text = '退回修改'; semanticType = 'document'; stencil = 'BASFLO_M.VSSX'; x = 5.0; y = 2.1; width = 1.4; height = 0.6 }
  )
  edges = @(
    [pscustomobject]@{ from = 'start'; to = 'review'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'review'; to = 'decision'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'decision'; to = 'end'; text = '通过'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'decision'; to = 'reject'; text = '拒绝'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'reject'; to = 'review'; routeType = 'curved'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX'; fromX = 0.5; fromY = 0.0; toX = 0.5; toY = 1.0 }
  )
}
$approvalResult = Save-BusinessCase -Name 'approval-flowchart' -Kind 'flowchart' -Graph $approvalGraph -FileName 'approval-flowchart.vsdx' -ExpectedRouteTypes @('orthogonal', 'curved')
$results.cases += $approvalResult
$results.files += $approvalResult.outputPath
$results.files += $approvalResult.previewPath

# Case 2: BPMN with lanes and gateways.
$bpmnGraph = [pscustomobject]@{
  diagramType = 'bpmn'
  page = [pscustomobject]@{ width = 10.0; height = 6.0 }
  nodes = @(
    [pscustomobject]@{ id = 'start'; text = '开始'; semanticType = 'start'; stencil = 'BPMN_M.VSSX'; master = 'Start Event'; x = 1.0; y = 4.4; width = 1.2; height = 0.58 }
    [pscustomobject]@{ id = 'task1'; text = '提交申请'; semanticType = 'task'; stencil = 'BPMN_M.VSSX'; master = 'Task'; x = 3.0; y = 4.4; width = 1.5; height = 0.62 }
    [pscustomobject]@{ id = 'gateway'; text = '通过?'; semanticType = 'gateway'; stencil = 'BPMN_M.VSSX'; master = 'Gateway'; x = 5.1; y = 4.4; width = 1.2; height = 0.82 }
    [pscustomobject]@{ id = 'task2'; text = '处理结果'; semanticType = 'task'; stencil = 'BPMN_M.VSSX'; master = 'Task'; x = 7.2; y = 4.4; width = 1.4; height = 0.62 }
    [pscustomobject]@{ id = 'end'; text = '结束'; semanticType = 'end'; stencil = 'BPMN_M.VSSX'; master = 'End Event'; x = 5.1; y = 2.4; width = 1.2; height = 0.58 }
  )
  edges = @(
    [pscustomobject]@{ from = 'start'; to = 'task1'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'task1'; to = 'gateway'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'gateway'; to = 'task2'; text = '是'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'gateway'; to = 'end'; text = '否'; routeType = 'curved'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX'; fromX = 0.5; fromY = 0.0; toX = 0.5; toY = 1.0 }
  )
}
$bpmnResult = Save-BusinessCase -Name 'bpmn-approval' -Kind 'bpmn' -Graph $bpmnGraph -FileName 'bpmn-approval.vsdx' -ExpectedRouteTypes @('orthogonal', 'curved')
$results.cases += $bpmnResult
$results.files += $bpmnResult.outputPath
$results.files += $bpmnResult.previewPath

# Case 3: network / architecture diagram.
$networkGraph = [pscustomobject]@{
  diagramType = 'network'
  page = [pscustomobject]@{ width = 10.0; height = 6.5 }
  nodes = @(
    [pscustomobject]@{ id = 'client'; text = '客户端'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 1.0; y = 4.6; width = 1.3; height = 0.6 }
    [pscustomobject]@{ id = 'lb'; text = '负载均衡'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 3.0; y = 4.6; width = 1.4; height = 0.6 }
    [pscustomobject]@{ id = 'api'; text = 'API 服务'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 5.0; y = 4.6; width = 1.4; height = 0.6 }
    [pscustomobject]@{ id = 'cache'; text = '缓存'; semanticType = 'database'; stencil = 'BASFLO_M.VSSX'; x = 5.0; y = 2.6; width = 1.1; height = 0.75 }
    [pscustomobject]@{ id = 'db'; text = '数据库'; semanticType = 'database'; stencil = 'BASFLO_M.VSSX'; x = 7.2; y = 2.6; width = 1.3; height = 0.75 }
  )
  edges = @(
    [pscustomobject]@{ from = 'client'; to = 'lb'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'lb'; to = 'api'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'api'; to = 'cache'; routeType = 'straight'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'api'; to = 'db'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
  )
}
$networkResult = Save-BusinessCase -Name 'network-architecture' -Kind 'network' -Graph $networkGraph -FileName 'network-architecture.vsdx' -LayoutStrategy 'Hierarchical' -LayoutDirection 'LR' -ExpectedRouteTypes @('orthogonal', 'straight')
$results.cases += $networkResult
$results.files += $networkResult.outputPath
$results.files += $networkResult.previewPath

# Case 4: DFD (Data Flow Diagram).
$dfdGraph = [pscustomobject]@{
  diagramType = 'dfd'
  page = [pscustomobject]@{ width = 9.5; height = 5.5 }
  nodes = @(
    [pscustomobject]@{ id = 'actor'; text = '用户'; semanticType = 'process'; stencil = 'DATFLO_M.VSSX'; master = 'External interactor'; x = 1.0; y = 3.5; width = 1.4; height = 0.7 }
    [pscustomobject]@{ id = 'process1'; text = '订单处理'; semanticType = 'process'; stencil = 'DATFLO_M.VSSX'; master = 'Data process'; x = 3.5; y = 3.5; width = 1.6; height = 0.85 }
    [pscustomobject]@{ id = 'process2'; text = '库存查询'; semanticType = 'process'; stencil = 'DATFLO_M.VSSX'; master = 'Data process'; x = 6.0; y = 3.5; width = 1.6; height = 0.85 }
    [pscustomobject]@{ id = 'datastore'; text = '商品库'; semanticType = 'database'; stencil = 'DATFLO_M.VSSX'; master = 'Data store'; x = 6.0; y = 1.8; width = 1.6; height = 0.6 }
  )
  edges = @(
    [pscustomobject]@{ from = 'actor'; to = 'process1'; text = '下单'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'process1'; to = 'process2'; text = '查询库存'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'process2'; to = 'datastore'; text = '读取'; routeType = 'straight'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
  )
}
$dfdResult = Save-BusinessCase -Name 'dfd-order-processing' -Kind 'dfd' -Graph $dfdGraph -FileName 'dfd-order-processing.vsdx' -ExpectedRouteTypes @('orthogonal', 'straight')
$results.cases += $dfdResult
$results.files += $dfdResult.outputPath
$results.files += $dfdResult.previewPath

# Case 5: org chart.
$orgChartGraph = [pscustomobject]@{
  diagramType = 'orgchart'
  page = [pscustomobject]@{ width = 8.5; height = 6.5 }
  nodes = @(
    [pscustomobject]@{ id = 'ceo'; text = '总经理'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 4.2; y = 5.5; width = 1.5; height = 0.6 }
    [pscustomobject]@{ id = 'cto'; text = '技术总监'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 2.0; y = 3.8; width = 1.5; height = 0.6 }
    [pscustomobject]@{ id = 'cfo'; text = '财务总监'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 6.5; y = 3.8; width = 1.5; height = 0.6 }
    [pscustomobject]@{ id = 'dev'; text = '开发团队'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 1.0; y = 2.0; width = 1.5; height = 0.6 }
    [pscustomobject]@{ id = 'ops'; text = '运维团队'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 3.0; y = 2.0; width = 1.5; height = 0.6 }
    [pscustomobject]@{ id = 'acct'; text = '会计部'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 6.5; y = 2.0; width = 1.5; height = 0.6 }
  )
  edges = @(
    [pscustomobject]@{ from = 'ceo'; to = 'cto'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'ceo'; to = 'cfo'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'cto'; to = 'dev'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'cto'; to = 'ops'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'cfo'; to = 'acct'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
  )
}
$orgChartResult = Save-BusinessCase -Name 'org-chart' -Kind 'orgchart' -Graph $orgChartGraph -FileName 'org-chart.vsdx' -ExpectedRouteTypes @('orthogonal')
$results.cases += $orgChartResult
$results.files += $orgChartResult.outputPath
$results.files += $orgChartResult.previewPath

# Case 6: UML component diagram.
$umlGraph = [pscustomobject]@{
  diagramType = 'uml-component'
  page = [pscustomobject]@{ width = 10.0; height = 5.5 }
  nodes = @(
    [pscustomobject]@{ id = 'webapp'; text = 'Web 前端'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 1.2; y = 3.7; width = 1.6; height = 0.7 }
    [pscustomobject]@{ id = 'apigw'; text = 'API 网关'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 3.8; y = 3.7; width = 1.5; height = 0.7 }
    [pscustomobject]@{ id = 'svc'; text = '业务服务'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 6.2; y = 3.7; width = 1.5; height = 0.7 }
    [pscustomobject]@{ id = 'msgq'; text = '消息队列'; semanticType = 'database'; stencil = 'BASFLO_M.VSSX'; x = 6.2; y = 1.8; width = 1.5; height = 0.75 }
    [pscustomobject]@{ id = 'auth'; text = '认证服务'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 3.8; y = 1.8; width = 1.5; height = 0.7 }
  )
  edges = @(
    [pscustomobject]@{ from = 'webapp'; to = 'apigw'; text = 'HTTPS'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'apigw'; to = 'svc'; text = 'gRPC'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'apigw'; to = 'auth'; text = '验证'; routeType = 'straight'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'svc'; to = 'msgq'; text = '发布'; routeType = 'straight'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
  )
}
$umlResult = Save-BusinessCase -Name 'uml-component' -Kind 'uml-component' -Graph $umlGraph -FileName 'uml-component.vsdx' -LayoutStrategy 'Hierarchical' -LayoutDirection 'LR' -ExpectedRouteTypes @('orthogonal', 'straight')
$results.cases += $umlResult
$results.files += $umlResult.outputPath
$results.files += $umlResult.previewPath

# Promote actual PNGs to expected golden images when -ApproveGolden is passed.
if ($ApproveGolden) {
  foreach ($case in @($results.cases)) {
    if ($case.previewExists) {
      $actualPreviewPath = Join-Path $actualRoot ([System.IO.Path]::GetFileName($case.previewPath))
      $expectedPreviewPath = Join-Path $expectedRoot ([System.IO.Path]::GetFileName($case.previewPath))
      if (Test-Path -LiteralPath $actualPreviewPath) {
        Copy-Item -LiteralPath $actualPreviewPath -Destination $expectedPreviewPath -Force
      } elseif (Test-Path -LiteralPath $case.previewPath) {
        Copy-Item -LiteralPath $case.previewPath -Destination $expectedPreviewPath -Force
      }
    }
  }
}

# Categorize failures.
foreach ($case in @($results.cases)) {
  if (-not $case.overallPassed) {
    switch ($case.failureCategory) {
      'structural' { $results.structuralFailures += $case.name }
      'visual' { $results.visualFailures += $case.name }
    }
  }
}

$results.summary = [ordered]@{
  caseCount = @($results.cases).Count
  passedCount = @($results.cases | Where-Object { $_.overallPassed }).Count
  failedCount = @($results.cases | Where-Object { -not $_.overallPassed }).Count
  structuralFailureCount = @($results.structuralFailures).Count
  visualFailureCount = @($results.visualFailures).Count
  outputCount = @($results.files | Where-Object { $_ -and (Test-Path -LiteralPath $_) }).Count
  goldenRoot = $goldenRoot
  goldenApproveMode = [bool] $ApproveGolden
}

$statsPath = Join-Path $OutputDirectory 'regression-stats.json'
($results | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $statsPath -Encoding utf8

$results | ConvertTo-Json -Depth 8
