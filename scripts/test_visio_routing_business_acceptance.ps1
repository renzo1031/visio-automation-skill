param(
  [string] $OutputDirectory = '',
  [switch] $Visible
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'skills\visio-automation\scripts\visio_helpers.ps1')

if (-not $OutputDirectory) {
  $OutputDirectory = Join-Path $repoRoot 'out\routing-business-acceptance'
}
$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

function Assert-Condition {
  param(
    [Parameter(Mandatory = $true)] [bool] $Condition,
    [Parameter(Mandatory = $true)] [string] $Message
  )

  if (-not $Condition) {
    throw $Message
  }
}

function Close-OpenOutputDocument {
  param([Parameter(Mandatory = $true)] [string] $Path)

  $openDocument = Get-OpenVisioDocumentByPath -Path $Path
  if ($null -ne $openDocument) {
    try { $openDocument.Saved = $true } catch {}
    try { $openDocument.Close() } catch {}
    Start-Sleep -Milliseconds 300
  }
}

function Assert-VisioOutputBasicStructure {
  param(
    [Parameter(Mandatory = $true)] [string] $VsdxPath,
    [Parameter(Mandatory = $true)] [int] $ExpectedNodeCount,
    [Parameter(Mandatory = $true)] [int] $ExpectedEdgeCount
  )

  $visio = $null
  $doc = $null
  try {
    $visio = New-InvisibleVisioApplication
    $doc = $visio.Documents.OpenEx($VsdxPath, 64 + 2)
    $page = $visio.ActivePage
    $shapes = @($page.Shapes)
    $connectors = @($shapes | Where-Object { $_.OneD -ne 0 })
    $nodes = @($shapes | Where-Object { $_.OneD -eq 0 })
    $unglued = @($connectors | Where-Object {
      $beginX = $_.CellsU('BeginX').FormulaU
      $endX = $_.CellsU('EndX').FormulaU
      -not (($beginX -match 'Connections|Glue|GUARD|Pin') -and ($endX -match 'Connections|Glue|GUARD|Pin'))
    })

    return [pscustomobject]@{
      passed = ($nodes.Count -eq $ExpectedNodeCount -and $connectors.Count -eq $ExpectedEdgeCount -and $unglued.Count -eq 0)
      nodeCount = $nodes.Count
      edgeCount = $connectors.Count
      ungluedConnectorCount = $unglued.Count
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

function Render-AcceptanceGraph {
  param(
    [Parameter(Mandatory = $true)] [string] $Name,
    [Parameter(Mandatory = $true)] [string] $Kind,
    [Parameter(Mandatory = $true)] $Graph,
    [Parameter(Mandatory = $true)] [string] $FileName,
    [string] $RoutingIntent = 'clean'
  )

  $outputPath = Join-Path $OutputDirectory $FileName
  $previewPath = [System.IO.Path]::ChangeExtension($outputPath, '.png')
  Close-OpenOutputDocument -Path $outputPath

  $renderArgs = @{
    GraphModel = $Graph
    OutputPath = $outputPath
    PreviewPath = $previewPath
    RoutingIntent = $RoutingIntent
    Force = $true
  }
  if ($Visible) {
    $renderArgs.Visible = $true
  } else {
    $renderArgs.Invisible = $true
  }

  $result = Render-VisioGraphModel @renderArgs
  $structure = Assert-VisioOutputBasicStructure -VsdxPath $outputPath -ExpectedNodeCount @($Graph.nodes).Count -ExpectedEdgeCount @($Graph.edges).Count

  Assert-Condition -Condition (Test-Path -LiteralPath $outputPath) -Message "$Name did not create a .vsdx file."
  Assert-Condition -Condition (Test-Path -LiteralPath $previewPath) -Message "$Name did not create a preview PNG."
  Assert-Condition -Condition $structure.passed -Message "$Name output structure failed: nodes=$($structure.nodeCount), edges=$($structure.edgeCount), unglued=$($structure.ungluedConnectorCount)."

  return [pscustomobject]@{
    name = $Name
    kind = $Kind
    outputPath = $outputPath
    previewPath = $previewPath
    routingIntent = $result.routingIntent
    nodes = @($result.nodes).Count
    edges = @($result.edges).Count
    renderedEdges = @($result.edges)
    structure = $structure
  }
}

$flowLoopGraph = [pscustomobject]@{
  diagramType = 'flowchart'
  routingIntent = 'balanced'
  page = [pscustomobject]@{ width = 9.0; height = 5.4 }
  nodes = @(
    [pscustomobject]@{ id = 'start'; text = '员工提交'; semanticType = 'start'; stencil = 'BASFLO_M.VSSX'; x = 1.0; y = 3.7; width = 1.35; height = 0.58 }
    [pscustomobject]@{ id = 'review'; text = '主管审批'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 3.0; y = 3.7; width = 1.45; height = 0.62 }
    [pscustomobject]@{ id = 'decision'; text = '是否通过?'; semanticType = 'decision'; stencil = 'BASFLO_M.VSSX'; x = 5.0; y = 3.7; width = 1.35; height = 0.85 }
    [pscustomobject]@{ id = 'end'; text = '归档'; semanticType = 'end'; stencil = 'BASFLO_M.VSSX'; x = 7.0; y = 3.7; width = 1.2; height = 0.58 }
    [pscustomobject]@{ id = 'revise'; text = '返回修改'; semanticType = 'document'; stencil = 'BASFLO_M.VSSX'; x = 5.0; y = 2.0; width = 1.45; height = 0.62 }
    [pscustomobject]@{ id = 'note'; text = '审批意见'; semanticType = 'document'; stencil = 'BASFLO_M.VSSX'; x = 4.0; y = 2.75; width = 1.0; height = 0.52 }
  )
  annotations = @(
    [pscustomobject]@{ id = 'loop-note'; text = '退回说明'; x = 4.0; y = 2.35; width = 1.2; height = 0.3 }
  )
  edges = @(
    [pscustomobject]@{ from = 'start'; to = 'review'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'review'; to = 'decision'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'decision'; to = 'end'; text = '通过'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'decision'; to = 'revise'; text = '拒绝'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'revise'; to = 'review'; text = '重新提交'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
  )
}

$swimlaneGraph = [pscustomobject]@{
  diagramType = 'workflow'
  routingIntent = 'clean'
  page = [pscustomobject]@{ width = 9.5; height = 5.4 }
  nodes = @(
    [pscustomobject]@{ id = 'request'; text = '提交申请'; semanticType = 'start'; stencil = 'BASFLO_M.VSSX'; x = 1.0; y = 3.4; width = 1.35; height = 0.58 }
    [pscustomobject]@{ id = 'archive'; text = '归档完成'; semanticType = 'end'; stencil = 'BASFLO_M.VSSX'; x = 6.0; y = 3.4; width = 1.35; height = 0.58 }
    [pscustomobject]@{ id = 'finance-review'; text = '财务复核'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 3.5; y = 2.0; width = 1.45; height = 0.62 }
    [pscustomobject]@{ id = 'finance-done'; text = '付款确认'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 6.0; y = 2.0; width = 1.45; height = 0.62 }
  )
  swimlanes = @(
    [pscustomobject]@{ id = 'finance-lane'; text = '财务泳道'; x = 3.5; y = 3.4; width = 1.35; height = 1.7 }
  )
  labels = @(
    [pscustomobject]@{ id = 'sla'; text = 'SLA 说明'; x = 3.5; y = 3.4; width = 1.1; height = 0.3 }
  )
  edges = @(
    [pscustomobject]@{ from = 'request'; to = 'archive'; text = '直接归档'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'finance-review'; to = 'finance-done'; text = '确认'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
  )
}

$networkGraph = [pscustomobject]@{
  diagramType = 'network'
  routingIntent = 'clean'
  page = [pscustomobject]@{ width = 8.8; height = 6.2 }
  nodes = @(
    [pscustomobject]@{ id = 'client'; text = '客户端'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 1.0; y = 3.6; width = 1.3; height = 0.6 }
    [pscustomobject]@{ id = 'api'; text = 'API 网关'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 5.0; y = 3.6; width = 1.35; height = 0.6 }
    [pscustomobject]@{ id = 'monitor'; text = '监控'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 3.0; y = 5.0; width = 1.2; height = 0.6 }
    [pscustomobject]@{ id = 'database'; text = '数据库'; semanticType = 'database'; stencil = 'BASFLO_M.VSSX'; x = 3.0; y = 2.2; width = 1.25; height = 0.72 }
    [pscustomobject]@{ id = 'cache'; text = '缓存'; semanticType = 'database'; stencil = 'BASFLO_M.VSSX'; x = 5.0; y = 2.2; width = 1.2; height = 0.72 }
  )
  edges = @(
    [pscustomobject]@{ from = 'client'; to = 'api'; text = 'HTTPS'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'monitor'; to = 'database'; text = '指标采集'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'api'; to = 'cache'; text = '读写缓存'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
  )
}

$caseResults = @()

$flowCase = Render-AcceptanceGraph -Name 'flowchart-loopback-avoidance' -Kind 'flowchart-loop' -Graph $flowLoopGraph -FileName 'flowchart-loopback-avoidance.vsdx' -RoutingIntent 'balanced'
$flowLoopEdge = @($flowCase.renderedEdges | Where-Object { $_.from -eq 'revise' -and $_.to -eq 'review' } | Select-Object -First 1)
Assert-Condition -Condition ($null -ne $flowLoopEdge) -Message 'Flow loopback case did not return the revise->review edge.'
Assert-Condition -Condition ($flowLoopEdge.isLoopback -eq $true) -Message 'Flow loopback edge was not detected as loopback.'
Assert-Condition -Condition ([int] $flowLoopEdge.obstacleHits -eq 0) -Message "Flow loopback edge should avoid node obstacles, got obstacleHits=$($flowLoopEdge.obstacleHits)."
Assert-Condition -Condition ([int] $flowLoopEdge.labelHits -eq 0) -Message "Flow loopback edge should avoid text labels, got labelHits=$($flowLoopEdge.labelHits)."
$flowCase | Add-Member -NotePropertyName keyEdge -NotePropertyValue $flowLoopEdge -Force
$caseResults += $flowCase

$swimlaneCase = Render-AcceptanceGraph -Name 'swimlane-boundary-avoidance' -Kind 'swimlane-workflow' -Graph $swimlaneGraph -FileName 'swimlane-boundary-avoidance.vsdx' -RoutingIntent 'clean'
$swimlaneEdge = @($swimlaneCase.renderedEdges | Where-Object { $_.from -eq 'request' -and $_.to -eq 'archive' } | Select-Object -First 1)
Assert-Condition -Condition ($null -ne $swimlaneEdge) -Message 'Swimlane case did not return the request->archive edge.'
Assert-Condition -Condition ([int] $swimlaneEdge.boundaryHits -eq 0) -Message "Swimlane edge should avoid unrelated lane bounds, got boundaryHits=$($swimlaneEdge.boundaryHits)."
Assert-Condition -Condition ([int] $swimlaneEdge.labelHits -eq 0) -Message "Swimlane edge should avoid lane label bounds, got labelHits=$($swimlaneEdge.labelHits)."
$swimlaneCase | Add-Member -NotePropertyName keyEdge -NotePropertyValue $swimlaneEdge -Force
$caseResults += $swimlaneCase

$networkCase = Render-AcceptanceGraph -Name 'network-crossing-avoidance' -Kind 'network-architecture' -Graph $networkGraph -FileName 'network-crossing-avoidance.vsdx' -RoutingIntent 'clean'
$networkEdge = @($networkCase.renderedEdges | Where-Object { $_.from -eq 'monitor' -and $_.to -eq 'database' } | Select-Object -First 1)
Assert-Condition -Condition ($null -ne $networkEdge) -Message 'Network case did not return the monitor->database edge.'
Assert-Condition -Condition ([int] $networkEdge.crossingHits -eq 0) -Message "Network monitoring edge should avoid existing route crossings, got crossingHits=$($networkEdge.crossingHits)."
Assert-Condition -Condition ([int] $networkEdge.obstacleHits -eq 0) -Message "Network monitoring edge should avoid node obstacles, got obstacleHits=$($networkEdge.obstacleHits)."
$networkCase | Add-Member -NotePropertyName keyEdge -NotePropertyValue $networkEdge -Force
$caseResults += $networkCase

$results = [ordered]@{
  outputDirectory = $OutputDirectory
  visibleMode = [bool] $Visible
  cases = @($caseResults | ForEach-Object {
    [pscustomobject]@{
      name = $_.name
      kind = $_.kind
      outputPath = $_.outputPath
      previewPath = $_.previewPath
      nodes = $_.nodes
      edges = $_.edges
      routingIntent = $_.routingIntent
      keyEdge = [pscustomobject]@{
        from = $_.keyEdge.from
        to = $_.keyEdge.to
        text = $_.keyEdge.text
        routeType = $_.keyEdge.routeType
        fromSide = $_.keyEdge.fromSide
        toSide = $_.keyEdge.toSide
        isLoopback = $_.keyEdge.isLoopback
        obstacleHits = $_.keyEdge.obstacleHits
        labelHits = $_.keyEdge.labelHits
        boundaryHits = $_.keyEdge.boundaryHits
        crossingHits = $_.keyEdge.crossingHits
        routeLength = $_.keyEdge.routeLength
      }
      structure = $_.structure
    }
  })
  summary = [ordered]@{
    caseCount = @($caseResults).Count
    outputCount = @($caseResults | ForEach-Object { @($_.outputPath, $_.previewPath) } | Where-Object { $_ -and (Test-Path -LiteralPath $_) }).Count
    allPassed = $true
  }
}

$resultPath = Join-Path $OutputDirectory 'routing-business-acceptance-result.json'
$results | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $resultPath -Encoding utf8
$results | ConvertTo-Json -Depth 8
