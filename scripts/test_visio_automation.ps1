param(
  [string] $OutputDirectory = (Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'out') 'visio-automation-skill-test')
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'skills\visio-automation\scripts\visio_helpers.ps1')

$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

$results = [ordered]@{
  outputDirectory = $OutputDirectory
  helperLoaded = $true
  knownMasterTest = $null
  explicitStraightTest = $null
  explicitCurvedTest = $null
  recoveryDiscoveryTest = $null
  environmentTest = $null
  masterIndexTest = $null
  graphModelTest = $null
  connectorGluePointTest = $null
  autoLayoutTest = $null
  mermaidTest = $null
  drawioTest = $null
  textFlowTest = $null
  naturalLanguageFlowTest = $null
  imageReconstructionTest = $null
  semanticMasterTest = $null
  sessionCacheTest = $null
  renderVisibilityTest = $null
  routingIntentTest = $null
  routingPolicyTest = $null
  files = @()
}

function Assert-Condition {
  param(
    [Parameter(Mandatory = $true)] [bool] $Condition,
    [Parameter(Mandatory = $true)] [string] $Message
  )

  if (-not $Condition) {
    throw $Message
  }
}

# Test -1: render helpers should default to a visible Visio application for user-facing output.
$visibilityProbe = & {
  $ctorCalls = New-Object System.Collections.Generic.List[string]
  $originalVisible = (Get-Command New-VisibleVisioApplication -CommandType Function).ScriptBlock
  $originalInvisible = (Get-Command New-InvisibleVisioApplication -CommandType Function).ScriptBlock

  try {
    Set-Item -Path function:New-VisibleVisioApplication -Value ([scriptblock]::Create(@"
`$null = `$ctorCalls.Add('visible')
throw 'visibility probe reached visible constructor'
"@))
    Set-Item -Path function:New-InvisibleVisioApplication -Value ([scriptblock]::Create(@"
`$null = `$ctorCalls.Add('invisible')
throw 'visibility probe reached invisible constructor'
"@))

    $graph = [pscustomobject]@{
      diagramType = 'flowchart'
      layout = [pscustomobject]@{
        strategy = 'Flowchart'
        direction = 'LR'
      }
      nodes = @()
      edges = @()
    }

    try {
      Render-VisioGraphModel -GraphModel $graph -OutputPath (Join-Path $OutputDirectory 'visibility-default-check.vsdx') -Force | Out-Null
    } catch {
      if ($_.Exception.Message -ne 'visibility probe reached visible constructor') {
        throw
      }
    }
    try {
      Render-VisioGraphModel -GraphModel $graph -OutputPath (Join-Path $OutputDirectory 'visibility-invisible-check.vsdx') -Invisible -Force | Out-Null
    } catch {
      if ($_.Exception.Message -ne 'visibility probe reached invisible constructor') {
        throw
      }
    }
  } catch {
    throw
  } finally {
    Set-Item -Path function:New-VisibleVisioApplication -Value $originalVisible
    Set-Item -Path function:New-InvisibleVisioApplication -Value $originalInvisible
  }

  Assert-Condition -Condition (@($ctorCalls).Count -gt 0) -Message 'Render-VisioGraphModel did not invoke any Visio application constructor.'
  Assert-Condition -Condition ($ctorCalls[0] -eq 'visible') -Message 'Render-VisioGraphModel should default to New-VisibleVisioApplication.'
  Assert-Condition -Condition ($ctorCalls[1] -eq 'invisible') -Message 'Render-VisioGraphModel -Invisible should use New-InvisibleVisioApplication.'

  [pscustomobject]@{
    constructors = @($ctorCalls)
  }
}

$results.renderVisibilityTest = [ordered]@{
  passed = $true
  constructors = @($visibilityProbe.constructors)
}

# Test 0.5: orthogonal connectors should prefer top/bottom when nodes are vertically stacked.
$verticalGlueProbe = Resolve-VisioConnectorGluePoints -From ([pscustomobject]@{ x = 2.0; y = 5.0 }) -To ([pscustomobject]@{ x = 2.0; y = 3.0 })
Assert-Condition -Condition (
  $verticalGlueProbe.fromX -eq 0.5 -and
  $verticalGlueProbe.fromY -eq 0.0 -and
  $verticalGlueProbe.toX -eq 0.5 -and
  $verticalGlueProbe.toY -eq 1.0
) -Message 'Vertical connectors should default to bottom-to-top glue points.'

$results.connectorGluePointTest = [ordered]@{
  passed = $true
  fromX = $verticalGlueProbe.fromX
  fromY = $verticalGlueProbe.fromY
  toX = $verticalGlueProbe.toX
  toY = $verticalGlueProbe.toY
}

# Test 0: inspect the local Visio environment and build a reusable master index.
$indexDirectory = Join-Path $OutputDirectory 'cache'
$indexPath = Join-Path $indexDirectory 'visio-master-index.json'
$environment = Get-VisioEnvironment
Assert-Condition -Condition ($null -ne $environment.Version) -Message 'Visio environment did not report a version.'
Assert-Condition -Condition (@($environment.ContentRoots).Count -gt 0) -Message 'Visio environment did not report any content roots.'
Assert-Condition -Condition (@($environment.StencilPatterns).Count -gt 0) -Message 'Visio environment did not report stencil patterns.'

$index = Build-VisioMasterIndex -OutputPath $indexPath -PreferredStencilRegex 'BPMN|DATFLO' -Query 'Gateway|Data process|Dynamic connector' -Force
Assert-Condition -Condition (Test-Path -LiteralPath $indexPath) -Message 'Master index JSON file was not created.'
Assert-Condition -Condition ($index.SchemaVersion -ge 1) -Message 'Master index schema version is missing.'
Assert-Condition -Condition ($index.Environment.Version -eq $environment.Version) -Message 'Master index did not preserve the Visio environment version.'
Assert-Condition -Condition (@($index.Masters).Count -gt 0) -Message 'Master index did not record any masters.'
Assert-Condition -Condition (@($index.Stencils).Count -gt 0) -Message 'Master index did not record stencil metadata.'
$indexedGateway = Find-VisioMasters -Query 'Gateway|网关' -PreferredStencilRegex 'BPMN' -IndexPath $indexPath -MaxResults 10 | Where-Object { $_.NameU -eq 'Gateway' } | Select-Object -First 1
Assert-Condition -Condition ($null -ne $indexedGateway) -Message 'Find-VisioMasters did not find Gateway from the master index.'
Assert-Condition -Condition ($indexedGateway.Source -eq 'Index') -Message 'Indexed Gateway result did not report Source=Index.'

$savedIndex = Get-Content -Raw -LiteralPath $indexPath | ConvertFrom-Json
$savedIndex.SchemaVersion = -1
$savedIndex | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $indexPath -Encoding utf8
Assert-Condition -Condition (-not (Test-VisioMasterIndexValid -IndexPath $indexPath -Environment $environment)) -Message 'Invalid schema version did not invalidate the master index.'
$index = Build-VisioMasterIndex -OutputPath $indexPath -PreferredStencilRegex 'BPMN|DATFLO' -Query 'Gateway|Data process|Dynamic connector' -Force

$results.environmentTest = [ordered]@{
  passed = $true
  version = $environment.Version
  language = $environment.Language
  contentRoots = @($environment.ContentRoots).Count
}
$results.masterIndexTest = [ordered]@{
  passed = $true
  indexPath = $indexPath
  schemaVersion = $index.SchemaVersion
  stencils = @($index.Stencils).Count
  masters = @($index.Masters).Count
  indexedGatewayStencil = $indexedGateway.Stencil
}
$results.files += $indexPath

# Test 1: render a tiny graph model directly into Visio.
$graphModelFile = Join-Path $OutputDirectory 'graph-model-smoke.vsdx'
$graphModelPreview = Join-Path $OutputDirectory 'graph-model-smoke.png'
$graphModel = [pscustomobject]@{
  diagramType = 'flowchart'
  nodes = @(
    [pscustomobject]@{
      id = 'start'
      text = '开始'
      semanticType = 'start'
      stencil = 'BASFLO_M.VSSX'
      x = 1.2
      y = 1.5
      width = 1.2
      height = 0.55
    }
    [pscustomobject]@{
      id = 'process'
      text = '处理'
      semanticType = 'process'
      stencil = 'BASFLO_M.VSSX'
      x = 3.7
      y = 1.5
      width = 1.7
      height = 0.62
    }
    [pscustomobject]@{
      id = 'decision'
      text = '通过?'
      semanticType = 'decision'
      stencil = 'BASFLO_M.VSSX'
      x = 5.8
      y = 1.5
      width = 1.2
      height = 0.85
    }
  )
  edges = @(
    [pscustomobject]@{
      from = 'start'
      to = 'process'
      connector = 'Dynamic connector'
      stencil = 'BASFLO_M.VSSX'
      routeType = 'orthogonal'
    }
    [pscustomobject]@{
      from = 'process'
      to = 'decision'
      connector = 'Dynamic connector'
      stencil = 'BASFLO_M.VSSX'
      routeType = 'straight'
    }
    [pscustomobject]@{
      from = 'decision'
      to = 'process'
      connector = 'Dynamic connector'
      stencil = 'BASFLO_M.VSSX'
      routeType = 'curved'
      fromX = 0.5
      fromY = 0.0
      toX = 0.5
      toY = 0.0
    }
  )
}

Assert-Condition -Condition ($null -ne (Get-Command Render-VisioGraphModel -ErrorAction SilentlyContinue)) -Message 'Render-VisioGraphModel is not defined yet.'
$renderedGraph = Render-VisioGraphModel -GraphModel $graphModel -OutputPath $graphModelFile -PreviewPath $graphModelPreview -Invisible -Force
Assert-Condition -Condition (Test-Path -LiteralPath $graphModelFile) -Message 'Graph model .vsdx file was not created.'
Assert-Condition -Condition ($renderedGraph.diagramType -eq 'flowchart') -Message 'Rendered graph model did not preserve the diagram type.'
Assert-Condition -Condition (@($renderedGraph.nodes).Count -eq 3) -Message 'Rendered graph model did not preserve node count.'
Assert-Condition -Condition (@($renderedGraph.edges).Count -eq 3) -Message 'Rendered graph model did not preserve edge count.'
Assert-Condition -Condition (@($renderedGraph.edges | Where-Object { $_.routeType -eq 'orthogonal' }).Count -eq 1) -Message 'Graph model did not render one orthogonal connector.'
Assert-Condition -Condition (@($renderedGraph.edges | Where-Object { $_.routeType -eq 'straight' }).Count -eq 1) -Message 'Graph model did not render one straight connector.'
Assert-Condition -Condition (@($renderedGraph.edges | Where-Object { $_.routeType -eq 'curved' }).Count -eq 1) -Message 'Graph model did not render one curved connector.'
$results.graphModelTest = [ordered]@{
  passed = $true
  file = $graphModelFile
  preview = $graphModelPreview
  diagramType = $renderedGraph.diagramType
  nodes = @($renderedGraph.nodes).Count
  edges = @($renderedGraph.edges).Count
  routeTypes = @($renderedGraph.edges | ForEach-Object { $_.routeType })
}
$results.files += $graphModelFile
$results.files += $graphModelPreview

# Test 1a: render a graph model without coordinates by applying a deterministic layout first.
$autoLayoutFile = Join-Path $OutputDirectory 'auto-layout-smoke.vsdx'
$autoLayoutPreview = Join-Path $OutputDirectory 'auto-layout-smoke.png'
$autoLayoutGraph = [pscustomobject]@{
  diagramType = 'flowchart'
  layout = [pscustomobject]@{
    strategy = 'Flowchart'
    direction = 'LR'
  }
  nodes = @(
    [pscustomobject]@{
      id = 'submit'
      text = '提交'
      semanticType = 'start'
      stencil = 'BASFLO_M.VSSX'
    }
    [pscustomobject]@{
      id = 'review'
      text = '审批'
      semanticType = 'process'
      stencil = 'BASFLO_M.VSSX'
    }
    [pscustomobject]@{
      id = 'archive'
      text = '归档'
      semanticType = 'end'
      stencil = 'BASFLO_M.VSSX'
    }
  )
  edges = @(
    [pscustomobject]@{
      from = 'submit'
      to = 'review'
      routeType = 'orthogonal'
      connector = 'Dynamic connector'
      stencil = 'BASFLO_M.VSSX'
    }
    [pscustomobject]@{
      from = 'review'
      to = 'archive'
      routeType = 'orthogonal'
      connector = 'Dynamic connector'
      stencil = 'BASFLO_M.VSSX'
    }
  )
}

Assert-Condition -Condition ($null -ne (Get-Command Invoke-VisioGraphLayout -ErrorAction SilentlyContinue)) -Message 'Invoke-VisioGraphLayout is not defined yet.'
$renderedAutoLayout = Render-VisioGraphModel -GraphModel $autoLayoutGraph -OutputPath $autoLayoutFile -PreviewPath $autoLayoutPreview -LayoutStrategy Flowchart -LayoutDirection LR -Invisible -Force
Assert-Condition -Condition (Test-Path -LiteralPath $autoLayoutFile) -Message 'Auto-layout graph .vsdx file was not created.'
Assert-Condition -Condition (Test-Path -LiteralPath $autoLayoutPreview) -Message 'Auto-layout graph preview was not created.'
Assert-Condition -Condition ($renderedAutoLayout.layout.strategy -eq 'Flowchart') -Message 'Auto-layout result did not report Flowchart strategy.'
Assert-Condition -Condition ($renderedAutoLayout.layout.direction -eq 'LR') -Message 'Auto-layout result did not report LR direction.'
Assert-Condition -Condition (@($renderedAutoLayout.nodes | Where-Object { $null -ne $_.x -and $null -ne $_.y }).Count -eq 3) -Message 'Auto-layout did not assign coordinates to every node.'
$autoLayoutXValues = @($renderedAutoLayout.nodes | Sort-Object id | ForEach-Object { [double] $_.x })
Assert-Condition -Condition (($autoLayoutXValues | Select-Object -Unique).Count -gt 1) -Message 'Auto-layout did not spread nodes horizontally.'
$results.autoLayoutTest = [ordered]@{
  passed = $true
  file = $autoLayoutFile
  preview = $autoLayoutPreview
  strategy = $renderedAutoLayout.layout.strategy
  direction = $renderedAutoLayout.layout.direction
  nodesWithCoordinates = @($renderedAutoLayout.nodes | Where-Object { $null -ne $_.x -and $null -ne $_.y }).Count
  edges = @($renderedAutoLayout.edges).Count
}
$results.files += $autoLayoutFile
$results.files += $autoLayoutPreview

# Test 1a.1: branch-aware flowchart layout should keep loopback edges from breaking the main reading order.
$branchLayoutGraph = [pscustomobject]@{
  diagramType = 'flowchart'
  layout = [pscustomobject]@{
    strategy = 'Flowchart'
    direction = 'LR'
  }
  nodes = @(
    [pscustomobject]@{ id = 'start'; text = '发起申请'; semanticType = 'start'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ id = 'review'; text = '部门审批'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ id = 'decision'; text = '是否通过?'; semanticType = 'decision'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ id = 'end'; text = '流程结束'; semanticType = 'end'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ id = 'revise'; text = '返回修改'; semanticType = 'document'; stencil = 'BASFLO_M.VSSX' }
  )
  edges = @(
    [pscustomobject]@{ from = 'start'; to = 'review'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'review'; to = 'decision'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'decision'; to = 'end'; text = '通过'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'decision'; to = 'revise'; text = '不通过'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'revise'; to = 'review'; text = '重新提交'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
  )
}
$branchLayout = Invoke-VisioGraphLayout -GraphModel $branchLayoutGraph -Strategy Flowchart -Direction LR
$branchNodesById = @{}
foreach ($node in @($branchLayout.nodes)) {
  $branchNodesById[[string] $node.id] = $node
}
Assert-Condition -Condition ([double] $branchNodesById.start.x -lt [double] $branchNodesById.review.x) -Message 'Branch layout should place review after start.'
Assert-Condition -Condition ([double] $branchNodesById.review.x -lt [double] $branchNodesById.decision.x) -Message 'Branch layout loopback should not push review after the decision node.'
Assert-Condition -Condition ([double] $branchNodesById.decision.x -lt [double] $branchNodesById.end.x) -Message 'Branch layout should keep the positive end branch to the right of decision.'
Assert-Condition -Condition ([double] $branchNodesById.revise.y -lt ([double] $branchNodesById.decision.y - 0.9)) -Message 'Branch layout should place the rejection/rework branch below the decision node.'
Assert-Condition -Condition ([double] $branchNodesById.revise.y -ge 1.0) -Message 'Branch layout should leave enough bottom margin for the lower branch.'
$results.autoLayoutTest.branchLayout = [ordered]@{
  start = [ordered]@{ x = $branchNodesById.start.x; y = $branchNodesById.start.y }
  review = [ordered]@{ x = $branchNodesById.review.x; y = $branchNodesById.review.y }
  decision = [ordered]@{ x = $branchNodesById.decision.x; y = $branchNodesById.decision.y }
  end = [ordered]@{ x = $branchNodesById.end.x; y = $branchNodesById.end.y }
  revise = [ordered]@{ x = $branchNodesById.revise.x; y = $branchNodesById.revise.y }
}

# Test 1b: convert a Mermaid flowchart to a Visio file.
$mermaidFile = Join-Path $OutputDirectory 'mermaid-smoke.vsdx'
$mermaidPreview = Join-Path $OutputDirectory 'mermaid-smoke.png'
$mermaidText = @'
graph TD
    A((开始)) --> B[审批]
    B -- yes --> C((结束))
'@

Assert-Condition -Condition ($null -ne (Get-Command Convert-MermaidToVisio -ErrorAction SilentlyContinue)) -Message 'Convert-MermaidToVisio is not defined yet.'
$mermaidResult = Convert-MermaidToVisio -MermaidText $mermaidText -OutputPath $mermaidFile -PreviewPath $mermaidPreview -Invisible -Force
Assert-Condition -Condition (Test-Path -LiteralPath $mermaidFile) -Message 'Mermaid .vsdx file was not created.'
Assert-Condition -Condition (Test-Path -LiteralPath $mermaidPreview) -Message 'Mermaid preview was not created.'
Assert-Condition -Condition ($mermaidResult.layout.direction -eq 'TB') -Message 'Mermaid graph TD direction was not converted to TB layout.'
Assert-Condition -Condition (@($mermaidResult.nodes).Count -eq 3) -Message 'Mermaid graph did not produce three nodes.'
Assert-Condition -Condition (@($mermaidResult.edges).Count -eq 2) -Message 'Mermaid graph did not produce two edges.'
Assert-Condition -Condition (@($mermaidResult.nodes | Where-Object { $_.semanticType -eq 'start' }).Count -eq 1) -Message 'Mermaid start node was not mapped to a Start/End master.'
Assert-Condition -Condition (@($mermaidResult.nodes | Where-Object { $_.semanticType -eq 'end' }).Count -eq 1) -Message 'Mermaid end node was not mapped to a Start/End master.'
Assert-Condition -Condition (@($mermaidResult.edges | Where-Object { $_.text -eq 'yes' }).Count -eq 1) -Message 'Mermaid edge label was not preserved.'
$results.mermaidTest = [ordered]@{
  passed = $true
  file = $mermaidFile
  preview = $mermaidPreview
  direction = $mermaidResult.layout.direction
  nodes = @($mermaidResult.nodes).Count
  edges = @($mermaidResult.edges).Count
  labeledEdges = @($mermaidResult.edges | Where-Object { $_.text -eq 'yes' }).Count
}
$results.files += $mermaidFile
$results.files += $mermaidPreview

# Test 1c: convert a simple uncompressed draw.io diagram to a Visio file.
$drawioFile = Join-Path $OutputDirectory 'drawio-smoke.vsdx'
$drawioPreview = Join-Path $OutputDirectory 'drawio-smoke.png'
$drawioXml = @'
<mxfile>
  <diagram name="Page-1">
    <mxGraphModel>
      <root>
        <mxCell id="0" />
        <mxCell id="1" parent="0" />
        <mxCell id="start" value="开始" style="ellipse;whiteSpace=wrap;html=1;" vertex="1" parent="1">
          <mxGeometry x="40" y="80" width="120" height="60" as="geometry" />
        </mxCell>
        <mxCell id="review" value="审批" style="rounded=1;whiteSpace=wrap;html=1;" vertex="1" parent="1">
          <mxGeometry x="240" y="80" width="120" height="60" as="geometry" />
        </mxCell>
        <mxCell id="end" value="结束" style="ellipse;whiteSpace=wrap;html=1;" vertex="1" parent="1">
          <mxGeometry x="440" y="80" width="120" height="60" as="geometry" />
        </mxCell>
        <mxCell id="edge-1" value="" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;" edge="1" parent="1" source="start" target="review">
          <mxGeometry relative="1" as="geometry" />
        </mxCell>
        <mxCell id="edge-2" value="yes" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;" edge="1" parent="1" source="review" target="end">
          <mxGeometry relative="1" as="geometry" />
        </mxCell>
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>
'@

Assert-Condition -Condition ($null -ne (Get-Command Convert-DrawIOToVisio -ErrorAction SilentlyContinue)) -Message 'Convert-DrawIOToVisio is not defined yet.'
$drawioResult = Convert-DrawIOToVisio -DrawIOXml $drawioXml -OutputPath $drawioFile -PreviewPath $drawioPreview -Invisible -Force
Assert-Condition -Condition (Test-Path -LiteralPath $drawioFile) -Message 'Draw.io .vsdx file was not created.'
Assert-Condition -Condition (Test-Path -LiteralPath $drawioPreview) -Message 'Draw.io preview was not created.'
Assert-Condition -Condition (@($drawioResult.nodes).Count -eq 3) -Message 'Draw.io graph did not produce three nodes.'
Assert-Condition -Condition (@($drawioResult.edges).Count -eq 2) -Message 'Draw.io graph did not produce two edges.'
Assert-Condition -Condition (@($drawioResult.nodes | Where-Object { $_.semanticType -eq 'start' }).Count -eq 1) -Message 'Draw.io start node was not mapped to a Start/End master.'
Assert-Condition -Condition (@($drawioResult.nodes | Where-Object { $_.semanticType -eq 'end' }).Count -eq 1) -Message 'Draw.io end node was not mapped to a Start/End master.'
Assert-Condition -Condition (@($drawioResult.edges | Where-Object { $_.text -eq 'yes' }).Count -eq 1) -Message 'Draw.io edge label was not preserved.'
$results.drawioTest = [ordered]@{
  passed = $true
  file = $drawioFile
  preview = $drawioPreview
  nodes = @($drawioResult.nodes).Count
  edges = @($drawioResult.edges).Count
  labeledEdges = @($drawioResult.edges | Where-Object { $_.text -eq 'yes' }).Count
}
$results.files += $drawioFile
$results.files += $drawioPreview

# Test 1c.5: convert a simple natural-language text flow to a Visio file.
$textFlowFile = Join-Path $OutputDirectory 'text-flow-smoke.vsdx'
$textFlowPreview = Join-Path $OutputDirectory 'text-flow-smoke.png'
$textFlow = '发起申请 -> 部门审批 -> 是否通过? --通过-> 流程结束; 是否通过? --拒绝-> 返回修改'

Assert-Condition -Condition ($null -ne (Get-Command Convert-TextFlowToVisio -ErrorAction SilentlyContinue)) -Message 'Convert-TextFlowToVisio is not defined yet.'
$textFlowGraph = Convert-TextFlowToVisioGraphModel -Text $textFlow
Assert-Condition -Condition ($textFlowGraph.diagramType -eq 'flowchart') -Message 'Text flow should produce a flowchart graph model.'
Assert-Condition -Condition (@($textFlowGraph.nodes).Count -eq 5) -Message 'Text flow graph did not produce five nodes.'
Assert-Condition -Condition (@($textFlowGraph.edges).Count -eq 4) -Message 'Text flow graph did not produce four edges.'
Assert-Condition -Condition (@($textFlowGraph.nodes | Where-Object { $_.semanticType -eq 'start' -and $_.text -eq '发起申请' }).Count -eq 1) -Message 'Text flow did not infer the start node from 发起申请.'
Assert-Condition -Condition (@($textFlowGraph.nodes | Where-Object { $_.semanticType -eq 'decision' -and $_.text -eq '是否通过?' }).Count -eq 1) -Message 'Text flow did not infer the decision node.'
Assert-Condition -Condition (@($textFlowGraph.edges | Where-Object { $_.text -eq '通过' }).Count -eq 1) -Message 'Text flow did not preserve the positive edge label.'
Assert-Condition -Condition (@($textFlowGraph.edges | Where-Object { $_.text -eq '拒绝' }).Count -eq 1) -Message 'Text flow did not preserve the negative edge label.'
$chineseTextFlowGraph = Convert-TextFlowToVisioGraphModel -Text '客户提交申请 到 系统校验 然后 人工复核 至 完成'
Assert-Condition -Condition (@($chineseTextFlowGraph.nodes).Count -eq 4) -Message 'Chinese text flow words did not produce four nodes.'
Assert-Condition -Condition (@($chineseTextFlowGraph.edges).Count -eq 3) -Message 'Chinese text flow words did not produce three edges.'
$textFlowResult = Convert-TextFlowToVisio -Text $textFlow -OutputPath $textFlowFile -PreviewPath $textFlowPreview -Invisible -Force
Assert-Condition -Condition (Test-Path -LiteralPath $textFlowFile) -Message 'Text flow .vsdx file was not created.'
Assert-Condition -Condition (Test-Path -LiteralPath $textFlowPreview) -Message 'Text flow preview was not created.'
Assert-Condition -Condition (@($textFlowResult.nodes).Count -eq 5) -Message 'Rendered text flow did not preserve node count.'
Assert-Condition -Condition (@($textFlowResult.edges).Count -eq 4) -Message 'Rendered text flow did not preserve edge count.'
$results.textFlowTest = [ordered]@{
  passed = $true
  file = $textFlowFile
  preview = $textFlowPreview
  nodes = @($textFlowResult.nodes).Count
  edges = @($textFlowResult.edges).Count
  labeledEdges = @($textFlowResult.edges | Where-Object { $_.text }).Count
}
$results.files += $textFlowFile
$results.files += $textFlowPreview

# Test 1c.6: convert a short natural-language business request to a Visio file.
$naturalLanguageFlowFile = Join-Path $OutputDirectory 'natural-language-flow-smoke.vsdx'
$naturalLanguageFlowPreview = Join-Path $OutputDirectory 'natural-language-flow-smoke.png'
$naturalLanguageFlowText = '帮我画一个请假审批流程：员工发起申请，部门经理审批，如果通过就流程结束，如果不通过就返回修改。'

Assert-Condition -Condition ($null -ne (Get-Command Convert-NaturalLanguageFlowToVisio -ErrorAction SilentlyContinue)) -Message 'Convert-NaturalLanguageFlowToVisio is not defined yet.'
$naturalLanguageFlowGraph = Convert-NaturalLanguageFlowToVisioGraphModel -Text $naturalLanguageFlowText
Assert-Condition -Condition ($naturalLanguageFlowGraph.diagramType -eq 'flowchart') -Message 'Natural language flow should produce a flowchart graph model.'
Assert-Condition -Condition (@($naturalLanguageFlowGraph.nodes).Count -eq 5) -Message 'Natural language flow did not produce five nodes.'
Assert-Condition -Condition (@($naturalLanguageFlowGraph.edges).Count -eq 4) -Message 'Natural language flow did not produce four edges.'
Assert-Condition -Condition (@($naturalLanguageFlowGraph.nodes | Where-Object { $_.semanticType -eq 'start' -and $_.text -match '发起申请' }).Count -eq 1) -Message 'Natural language flow did not infer the start application node.'
Assert-Condition -Condition (@($naturalLanguageFlowGraph.nodes | Where-Object { $_.semanticType -eq 'process' -and $_.text -match '审批' }).Count -eq 1) -Message 'Natural language flow did not infer the approval process node.'
Assert-Condition -Condition (@($naturalLanguageFlowGraph.nodes | Where-Object { $_.semanticType -eq 'decision' -and $_.text -eq '是否通过?' }).Count -eq 1) -Message 'Natural language flow did not infer the separate pass/fail decision node.'
Assert-Condition -Condition (@($naturalLanguageFlowGraph.edges | Where-Object { $_.text -eq '通过' }).Count -eq 1) -Message 'Natural language flow did not preserve the approval branch label.'
Assert-Condition -Condition (@($naturalLanguageFlowGraph.edges | Where-Object { $_.text -eq '不通过' }).Count -eq 1) -Message 'Natural language flow did not preserve the rejection branch label.'
$expenseFlowText = '画一个报销流程：员工提交报销单，系统校验发票，如果资料不完整就退回补充，如果完整就财务审核，审核通过后付款并归档，审核不通过就驳回。'
$expenseFlowGraph = Convert-NaturalLanguageFlowToVisioGraphModel -Text $expenseFlowText
Assert-Condition -Condition (@($expenseFlowGraph.nodes).Count -eq 9) -Message 'Expense natural language flow should produce nine nodes.'
Assert-Condition -Condition (@($expenseFlowGraph.edges).Count -eq 8) -Message 'Expense natural language flow should produce eight edges.'
Assert-Condition -Condition (@($expenseFlowGraph.nodes | Where-Object { $_.semanticType -eq 'decision' -and $_.text -eq '资料是否完整?' }).Count -eq 1) -Message 'Expense natural language flow did not infer the completeness decision.'
Assert-Condition -Condition (@($expenseFlowGraph.nodes | Where-Object { $_.semanticType -eq 'decision' -and $_.text -eq '审核是否通过?' }).Count -eq 1) -Message 'Expense natural language flow did not infer the audit decision.'
Assert-Condition -Condition (@($expenseFlowGraph.edges | Where-Object { $_.text -eq '不完整' }).Count -eq 1) -Message 'Expense natural language flow did not preserve the incomplete branch label.'
Assert-Condition -Condition (@($expenseFlowGraph.edges | Where-Object { $_.text -eq '完整' }).Count -eq 1) -Message 'Expense natural language flow did not preserve the complete branch label.'
Assert-Condition -Condition (@($expenseFlowGraph.edges | Where-Object { $_.text -eq '通过' }).Count -eq 1) -Message 'Expense natural language flow did not preserve the audit pass branch label.'
Assert-Condition -Condition (@($expenseFlowGraph.edges | Where-Object { $_.text -eq '不通过' }).Count -eq 1) -Message 'Expense natural language flow did not preserve the audit rejection branch label.'
$naturalLanguageFlowResult = Convert-NaturalLanguageFlowToVisio -Text $naturalLanguageFlowText -OutputPath $naturalLanguageFlowFile -PreviewPath $naturalLanguageFlowPreview -Invisible -Force
Assert-Condition -Condition (Test-Path -LiteralPath $naturalLanguageFlowFile) -Message 'Natural language flow .vsdx file was not created.'
Assert-Condition -Condition (Test-Path -LiteralPath $naturalLanguageFlowPreview) -Message 'Natural language flow preview was not created.'
Assert-Condition -Condition (@($naturalLanguageFlowResult.nodes).Count -eq 5) -Message 'Rendered natural language flow did not preserve node count.'
Assert-Condition -Condition (@($naturalLanguageFlowResult.edges).Count -eq 4) -Message 'Rendered natural language flow did not preserve edge count.'
$results.naturalLanguageFlowTest = [ordered]@{
  passed = $true
  file = $naturalLanguageFlowFile
  preview = $naturalLanguageFlowPreview
  nodes = @($naturalLanguageFlowResult.nodes).Count
  edges = @($naturalLanguageFlowResult.edges).Count
  labeledEdges = @($naturalLanguageFlowResult.edges | Where-Object { $_.text }).Count
  expenseNodes = @($expenseFlowGraph.nodes).Count
  expenseEdges = @($expenseFlowGraph.edges).Count
  expenseDecisions = @($expenseFlowGraph.nodes | Where-Object { $_.semanticType -eq 'decision' }).Count
}
$results.files += $naturalLanguageFlowFile
$results.files += $naturalLanguageFlowPreview

# Test 1d: convert a structured image reconstruction model to a Visio file.
$imageReconstructionFile = Join-Path $OutputDirectory 'image-reconstruction-smoke.vsdx'
$imageReconstructionPreview = Join-Path $OutputDirectory 'image-reconstruction-smoke.png'
$imageReconstructionModel = [pscustomobject]@{
  source = [pscustomobject]@{
    imagePath = 'sample-whiteboard.png'
    width = 600
    height = 300
  }
  page = [pscustomobject]@{
    width = 6.0
    height = 3.0
  }
  nodes = @(
    [pscustomobject]@{
      id = 'capture'
      text = '截图'
      semanticType = 'start'
      bounds = [pscustomobject]@{ x = 40; y = 80; width = 120; height = 60 }
    }
    [pscustomobject]@{
      id = 'extract'
      text = '识别'
      semanticType = 'process'
      bounds = [pscustomobject]@{ x = 240; y = 80; width = 120; height = 60 }
    }
    [pscustomobject]@{
      id = 'review'
      text = '人工复核?'
      semanticType = 'decision'
      bounds = [pscustomobject]@{ x = 430; y = 70; width = 120; height = 80 }
    }
  )
  connectors = @(
    [pscustomobject]@{
      id = 'c1'
      from = 'capture'
      to = 'extract'
      routeType = 'orthogonal'
      fromPoint = [pscustomobject]@{ x = 1.0; y = 0.5 }
      toPoint = [pscustomobject]@{ x = 0.0; y = 0.5 }
    }
    [pscustomobject]@{
      id = 'c2'
      from = 'extract'
      to = 'review'
      text = 'low confidence'
      routeType = 'straight'
      fromPoint = [pscustomobject]@{ x = 1.0; y = 0.5 }
      toPoint = [pscustomobject]@{ x = 0.0; y = 0.5 }
    }
    [pscustomobject]@{
      id = 'c3'
      from = 'review'
      to = 'extract'
      text = 'retry'
      routeType = 'curved'
      fromPoint = [pscustomobject]@{ x = 0.5; y = 0.0 }
      toPoint = [pscustomobject]@{ x = 0.5; y = 0.0 }
    }
  )
}

Assert-Condition -Condition ($null -ne (Get-Command Convert-ImageReconstructionToVisio -ErrorAction SilentlyContinue)) -Message 'Convert-ImageReconstructionToVisio is not defined yet.'
$imageReconstructionResult = Convert-ImageReconstructionToVisio -ReconstructionModel $imageReconstructionModel -OutputPath $imageReconstructionFile -PreviewPath $imageReconstructionPreview -Invisible -Force
Assert-Condition -Condition (Test-Path -LiteralPath $imageReconstructionFile) -Message 'Image reconstruction .vsdx file was not created.'
Assert-Condition -Condition (Test-Path -LiteralPath $imageReconstructionPreview) -Message 'Image reconstruction preview was not created.'
Assert-Condition -Condition (@($imageReconstructionResult.nodes).Count -eq 3) -Message 'Image reconstruction did not produce three nodes.'
Assert-Condition -Condition (@($imageReconstructionResult.edges).Count -eq 3) -Message 'Image reconstruction did not produce three connectors.'
Assert-Condition -Condition (@($imageReconstructionResult.edges | Where-Object { $_.routeType -eq 'orthogonal' }).Count -eq 1) -Message 'Image reconstruction did not preserve one orthogonal connector.'
Assert-Condition -Condition (@($imageReconstructionResult.edges | Where-Object { $_.routeType -eq 'straight' }).Count -eq 1) -Message 'Image reconstruction did not preserve one straight connector.'
Assert-Condition -Condition (@($imageReconstructionResult.edges | Where-Object { $_.routeType -eq 'curved' }).Count -eq 1) -Message 'Image reconstruction did not preserve one curved connector.'
Assert-Condition -Condition (@($imageReconstructionResult.edges | Where-Object { $_.fromX -eq 0.5 -and $_.fromY -eq 0.0 -and $_.toX -eq 0.5 -and $_.toY -eq 0.0 }).Count -eq 1) -Message 'Image reconstruction did not preserve relative glue points.'
Assert-Condition -Condition ($imageReconstructionResult.source.imagePath -eq 'sample-whiteboard.png') -Message 'Image reconstruction did not preserve source metadata.'
$results.imageReconstructionTest = [ordered]@{
  passed = $true
  file = $imageReconstructionFile
  preview = $imageReconstructionPreview
  sourceImage = $imageReconstructionResult.source.imagePath
  nodes = @($imageReconstructionResult.nodes).Count
  edges = @($imageReconstructionResult.edges).Count
  routeTypes = @($imageReconstructionResult.edges | ForEach-Object { $_.routeType })
  relativeGlueEdges = @($imageReconstructionResult.edges | Where-Object { $_.fromX -eq 0.5 -and $_.fromY -eq 0.0 -and $_.toX -eq 0.5 -and $_.toY -eq 0.0 }).Count
}
$results.files += $imageReconstructionFile
$results.files += $imageReconstructionPreview

# Test 1: create an editable Visio file from known built-in masters and verify default orthogonal glued connector.
$knownFile = Join-Path $OutputDirectory 'known-data-flow-smoke.vsdx'
$knownPreview = Join-Path $OutputDirectory 'known-data-flow-smoke.png'
$visio = New-InvisibleVisioApplication
try {
  $doc = $visio.Documents.Add('')
  $page = $visio.ActivePage
  $stencil = Open-VisioStencilReadOnly -Visio $visio -StencilNameOrPath 'DATFLO_M.VSSX'
  $externalMaster = $stencil.Masters.ItemU('External interactor')
  $processMaster = $stencil.Masters.ItemU('Data process')
  $connectorMaster = $stencil.Masters.ItemU('Dynamic connector')

  $user = $page.Drop($externalMaster, 1.0, 4.0)
  $user.Text = '用户'
  Set-VisioShapeFill -Shape $user -Fill 'RGB(70,118,242)'
  Set-VisioTextStyle -Shape $user -Color 'RGB(255,255,255)' -Bold 1

  $process = $page.Drop($processMaster, 3.2, 4.0)
  $process.Text = '处理'
  Set-VisioShapeFill -Shape $process -Fill 'RGB(239,129,219)'

  $connector = Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $user -To $process

  $null = $doc.SaveAs($knownFile)
  $null = $page.Export($knownPreview)

  Assert-Condition -Condition (Test-Path $knownFile) -Message 'Known master .vsdx file was not created.'
  Assert-Condition -Condition ($connector.OneD -ne 0) -Message 'Connector is not a OneD dynamic connector.'
  Assert-Condition -Condition ($connector.CellsU('ShapeRouteStyle').ResultIU -eq 0) -Message 'ShapeRouteStyle was not left at default orthogonal routing.'
  Assert-Condition -Condition ($connector.CellsU('ConLineRouteExt').ResultIU -eq 0) -Message 'ConLineRouteExt was not left at default orthogonal routing.'
  Assert-Condition -Condition ($connector.CellsU('BeginX').FormulaU -match 'Glue|GUARD|Connections|Pin') -Message 'BeginX does not appear to be glued.'
  Assert-Condition -Condition ($connector.CellsU('EndX').FormulaU -match 'Glue|GUARD|Connections|Pin') -Message 'EndX does not appear to be glued.'

  $results.knownMasterTest = [ordered]@{
    passed = $true
    file = $knownFile
    preview = $knownPreview
    shapes = $page.Shapes.Count
    connectorOneD = [int] $connector.OneD
    shapeRouteStyle = $connector.CellsU('ShapeRouteStyle').FormulaU
    conLineRouteExt = $connector.CellsU('ConLineRouteExt').FormulaU
    beginX = $connector.CellsU('BeginX').FormulaU
    endX = $connector.CellsU('EndX').FormulaU
  }
  $results.files += $knownFile
  $results.files += $knownPreview
  $null = $doc.Close()
} finally {
  $null = $visio.Quit()
}

# Test 1b: explicitly requested straight connectors still work.
$straightFile = Join-Path $OutputDirectory 'explicit-straight-smoke.vsdx'
$visio = New-InvisibleVisioApplication
try {
  $doc = $visio.Documents.Add('')
  $page = $visio.ActivePage
  $stencil = Open-VisioStencilReadOnly -Visio $visio -StencilNameOrPath 'DATFLO_M.VSSX'
  $externalMaster = $stencil.Masters.ItemU('External interactor')
  $processMaster = $stencil.Masters.ItemU('Data process')
  $connectorMaster = $stencil.Masters.ItemU('Dynamic connector')
  $from = $page.Drop($externalMaster, 1.0, 2.0)
  $to = $page.Drop($processMaster, 3.2, 2.0)
  $straightConnector = Connect-VisioShapesStraight -Page $page -ConnectorMaster $connectorMaster -From $from -To $to
  $null = $doc.SaveAs($straightFile)
  Assert-Condition -Condition ($straightConnector.CellsU('ShapeRouteStyle').ResultIU -eq 2) -Message 'Explicit straight connector did not set ShapeRouteStyle to 2.'
  Assert-Condition -Condition ($straightConnector.CellsU('ConLineRouteExt').ResultIU -eq 1) -Message 'Explicit straight connector did not set ConLineRouteExt to 1.'
  $results.explicitStraightTest = [ordered]@{
    passed = $true
    file = $straightFile
    shapeRouteStyle = $straightConnector.CellsU('ShapeRouteStyle').FormulaU
    conLineRouteExt = $straightConnector.CellsU('ConLineRouteExt').FormulaU
  }
  $results.files += $straightFile
  $null = $doc.Close()
} finally {
  $null = $visio.Quit()
}

# Test 1c: explicitly requested curved connectors still use a dynamic connector and preserve glued endpoints.
$curvedFile = Join-Path $OutputDirectory 'explicit-curved-smoke.vsdx'
$visio = New-InvisibleVisioApplication
try {
  $doc = $visio.Documents.Add('')
  $page = $visio.ActivePage
  $stencil = Open-VisioStencilReadOnly -Visio $visio -StencilNameOrPath 'DATFLO_M.VSSX'
  $externalMaster = $stencil.Masters.ItemU('External interactor')
  $processMaster = $stencil.Masters.ItemU('Data process')
  $connectorMaster = $stencil.Masters.ItemU('Dynamic connector')
  $from = $page.Drop($externalMaster, 1.0, 1.0)
  $to = $page.Drop($processMaster, 3.2, 1.4)
  $curvedConnector = Connect-VisioShapesCurved -Page $page -ConnectorMaster $connectorMaster -From $from -To $to
  $null = $doc.SaveAs($curvedFile)
  Assert-Condition -Condition ($curvedConnector.OneD -ne 0) -Message 'Curved connector is not a OneD dynamic connector.'
  Assert-Condition -Condition ($curvedConnector.CellsU('ConLineRouteExt').ResultIU -eq 2) -Message 'Explicit curved connector did not set ConLineRouteExt to 2.'
  Assert-Condition -Condition ($curvedConnector.CellsU('BeginX').FormulaU -match 'Glue|GUARD|Connections|Pin') -Message 'Curved connector BeginX does not appear to be glued.'
  Assert-Condition -Condition ($curvedConnector.CellsU('EndX').FormulaU -match 'Glue|GUARD|Connections|Pin') -Message 'Curved connector EndX does not appear to be glued.'
  $results.explicitCurvedTest = [ordered]@{
    passed = $true
    file = $curvedFile
    connectorOneD = [int] $curvedConnector.OneD
    shapeRouteStyle = $curvedConnector.CellsU('ShapeRouteStyle').FormulaU
    conLineRouteExt = $curvedConnector.CellsU('ConLineRouteExt').FormulaU
    beginX = $curvedConnector.CellsU('BeginX').FormulaU
    endX = $curvedConnector.CellsU('EndX').FormulaU
  }
  $results.files += $curvedFile
  $null = $doc.Close()
} finally {
  $null = $visio.Quit()
}

# Test 2: simulate an unknown request by discovering a BPMN gateway master that was not part of the original DFD workflow.
$matches = Find-VisioMasters -Query 'Gateway|网关' -PreferredStencilRegex 'BPMN' -MaxResults 10
$gatewayMatch = $matches | Where-Object { $_.NameU -eq 'Gateway' -or $_.Name -eq '网关' } | Select-Object -First 1
Assert-Condition -Condition ($null -ne $gatewayMatch) -Message 'Recovery discovery could not find a BPMN Gateway master.'

$recoveryFile = Join-Path $OutputDirectory 'recovery-bpmn-gateway-smoke.vsdx'
$recoveryPreview = Join-Path $OutputDirectory 'recovery-bpmn-gateway-smoke.png'
$visio = New-InvisibleVisioApplication
try {
  $doc = $visio.Documents.Add('')
  $page = $visio.ActivePage
  $stencil = Open-VisioStencilReadOnly -Visio $visio -StencilNameOrPath $gatewayMatch.FullName
  $gatewayMaster = $stencil.Masters.ItemU($gatewayMatch.NameU)
  $gateway = $page.Drop($gatewayMaster, 2.0, 2.0)
  $gateway.Text = '审批?'
  $null = $doc.SaveAs($recoveryFile)
  $null = $page.Export($recoveryPreview)

  Assert-Condition -Condition (Test-Path $recoveryFile) -Message 'Recovery proof .vsdx file was not created.'
  Assert-Condition -Condition ($gateway.NameU -match 'Gateway') -Message 'Dropped recovery shape is not the expected BPMN Gateway.'

  $results.recoveryDiscoveryTest = [ordered]@{
    passed = $true
    query = 'Gateway|网关'
    discoveredStencil = $gatewayMatch.Stencil
    discoveredStencilPath = $gatewayMatch.FullName
    discoveredName = $gatewayMatch.Name
    discoveredNameU = $gatewayMatch.NameU
    file = $recoveryFile
    preview = $recoveryPreview
    droppedShapeNameU = $gateway.NameU
  }
  $results.files += $recoveryFile
  $results.files += $recoveryPreview
  $null = $doc.Close()
} finally {
  $null = $visio.Quit()
}

# Test 3: semantic master search — find masters by natural language intent.
$semanticMapPath = Join-Path $repoRoot 'skills\visio-automation\references\semantic-master-map.json'
Assert-Condition -Condition (Test-Path -LiteralPath $semanticMapPath) -Message 'semantic-master-map.json was not found.'

$semanticResult1 = @(Find-SemanticVisioMaster -Query '数据库' -MapPath $semanticMapPath)
Assert-Condition -Condition (@($semanticResult1).Count -gt 0) -Message 'Find-SemanticVisioMaster did not return results for "数据库".'
Assert-Condition -Condition ($semanticResult1[0].Master -eq 'Database') -Message 'Semantic search for "数据库" did not map to Database master.'
Assert-Condition -Condition ($semanticResult1[0].Stencil -eq 'BASFLO_M.VSSX') -Message 'Semantic search for "数据库" did not map to BASFLO_M.VSSX stencil.'

$semanticResult2 = @(Find-SemanticVisioMaster -Query 'gateway' -MapPath $semanticMapPath -Category 'bpmn')
Assert-Condition -Condition (@($semanticResult2).Count -gt 0) -Message 'Find-SemanticVisioMaster did not return results for "gateway" in bpmn category.'
Assert-Condition -Condition ($semanticResult2[0].Master -eq 'Gateway') -Message 'Semantic search for "gateway" in bpmn did not map to Gateway master.'

$semanticResult3 = @(Find-SemanticVisioMaster -Query '网关' -MapPath $semanticMapPath)
Assert-Condition -Condition (@($semanticResult3).Count -gt 0) -Message 'Find-SemanticVisioMaster did not return results for "网关".'
$gatewayEntry = $semanticResult3 | Where-Object { $_.Master -eq 'Decision' -or $_.Master -eq 'Gateway' } | Select-Object -First 1
Assert-Condition -Condition ($null -ne $gatewayEntry) -Message 'Semantic search for "网关" did not match a decision/gateway master.'

$semanticResult4 = @(Find-SemanticVisioMaster -Query '负载均衡器' -MapPath $semanticMapPath -MaxResults 3)
# No fallback was requested, so an unmapped query should be an empty result set.
Assert-Condition -Condition (@($semanticResult4).Count -eq 0) -Message 'Find-SemanticVisioMaster should return an empty result set for unmapped queries without fallback.'

$semanticResult5 = @(Find-SemanticVisioMaster -Query '库存' -MapPath $semanticMapPath -Category 'vsm')
Assert-Condition -Condition (@($semanticResult5).Count -gt 0) -Message 'Find-SemanticVisioMaster did not return results for "库存" in vsm category.'
Assert-Condition -Condition ($semanticResult5[0].Master -eq 'Inventory') -Message 'Semantic search for "库存" in vsm did not map to Inventory master.'

$results.semanticMasterTest = [ordered]@{
  passed = $true
  mapPath = $semanticMapPath
  databaseQuery = @($semanticResult1).Count
  databaseMaster = $semanticResult1[0].Master
  gatewayCategoryQuery = @($semanticResult2).Count
  gatewayCategoryMaster = $semanticResult2[0].Master
  chineseGatewayQuery = @($semanticResult3).Count
  unmatchedQuery = @($semanticResult4).Count
  vsmInventoryQuery = @($semanticResult5).Count
  vsmInventoryMaster = $semanticResult5[0].Master
}

# Test 4: session cache — save, load, and update workspace state.
$testWorkspace = Join-Path $OutputDirectory 'test-workspace'
New-Item -ItemType Directory -Path $testWorkspace -Force | Out-Null

# Initial cache should be null
$initialCache = Get-VisioSessionCache -WorkspacePath $testWorkspace
Assert-Condition -Condition ($null -eq $initialCache) -Message 'Initial session cache should be null for new workspace.'

# Update cache with diagram type and stencil
$updatedCache = Update-VisioSessionCache -WorkspacePath $testWorkspace -DiagramType 'flowchart' -Stencil 'BASFLO_M.VSSX' -LastLayoutStrategy 'Flowchart' -LastLayoutDirection 'LR'
Assert-Condition -Condition ($null -ne $updatedCache) -Message 'Update-VisioSessionCache returned null.'
Assert-Condition -Condition ($updatedCache.diagramType -eq 'flowchart') -Message 'Session cache diagramType was not saved.'
Assert-Condition -Condition ($updatedCache.stencil -eq 'BASFLO_M.VSSX') -Message 'Session cache stencil was not saved.'
Assert-Condition -Condition ($updatedCache.lastLayoutStrategy -eq 'Flowchart') -Message 'Session cache layout strategy was not saved.'
Assert-Condition -Condition ($updatedCache.lastLayoutDirection -eq 'LR') -Message 'Session cache layout direction was not saved.'
Assert-Condition -Condition ($updatedCache.schemaVersion -eq 1) -Message 'Session cache schema version was not set.'

# Load cache and verify persistence
$loadedCache = Get-VisioSessionCache -WorkspacePath $testWorkspace
Assert-Condition -Condition ($null -ne $loadedCache) -Message 'Get-VisioSessionCache returned null after save.'
Assert-Condition -Condition ($loadedCache.diagramType -eq 'flowchart') -Message 'Session cache diagramType did not persist.'
Assert-Condition -Condition ($loadedCache.stencil -eq 'BASFLO_M.VSSX') -Message 'Session cache stencil did not persist.'

# Update with output path
$updateWithPath = Update-VisioSessionCache -WorkspacePath $testWorkspace -LastOutputPath 'E:\output\test.vsdx' -LastPreviewPath 'E:\output\test.png'
Assert-Condition -Condition ($updateWithPath.lastOutputPath -eq 'E:\output\test.vsdx') -Message 'Session cache lastOutputPath was not saved.'
Assert-Condition -Condition ($updateWithPath.lastPreviewPath -eq 'E:\output\test.png') -Message 'Session cache lastPreviewPath was not saved.'
Assert-Condition -Condition ($updateWithPath.diagramType -eq 'flowchart') -Message 'Session cache lost diagramType after partial update.'

# Add recent masters
$withRecent = Add-VisioSessionRecentMaster -WorkspacePath $testWorkspace -Stencil 'BASFLO_M.VSSX' -MasterNameU 'Process'
Assert-Condition -Condition (@($withRecent.recentMasters).Count -eq 1) -Message 'Session cache recentMasters did not track one master.'
Assert-Condition -Condition ($withRecent.recentMasters[0].master -eq 'Process') -Message 'Session cache recent master was not recorded.'

$withRecent2 = Add-VisioSessionRecentMaster -WorkspacePath $testWorkspace -Stencil 'BPMN_M.VSSX' -MasterNameU 'Gateway'
Assert-Condition -Condition (@($withRecent2.recentMasters).Count -eq 2) -Message 'Session cache recentMasters did not track two masters.'
Assert-Condition -Condition ($withRecent2.recentMasters[0].master -eq 'Gateway') -Message 'Most recent master should be Gateway.'

# Adding same master again should move it to front without duplicating
$withRecent3 = Add-VisioSessionRecentMaster -WorkspacePath $testWorkspace -Stencil 'BASFLO_M.VSSX' -MasterNameU 'Process'
Assert-Condition -Condition (@($withRecent3.recentMasters).Count -eq 2) -Message 'Session cache duplicated a recent master entry.'
Assert-Condition -Condition ($withRecent3.recentMasters[0].master -eq 'Process') -Message 'Re-added master should be at front of recent list.'

# Verify cache file exists in .cache directory
$cachePath = Get-VisioSessionCachePath -WorkspacePath $testWorkspace
Assert-Condition -Condition (Test-Path -LiteralPath $cachePath) -Message 'Session cache file was not created on disk.'

$results.sessionCacheTest = [ordered]@{
  passed = $true
  cachePath = $cachePath
  diagramType = $loadedCache.diagramType
  stencil = $loadedCache.stencil
  lastOutputPath = $updateWithPath.lastOutputPath
  lastPreviewPath = $updateWithPath.lastPreviewPath
  recentMastersCount = @($withRecent3.recentMasters).Count
  recentMasters = @($withRecent3.recentMasters | ForEach-Object { $_.master })
}

# Test 5: routing intent — auto-detection from sourceFormat and diagramType.
Assert-Condition -Condition ($null -ne (Get-Command Resolve-VisioRoutingIntent -ErrorAction SilentlyContinue)) -Message 'Resolve-VisioRoutingIntent is not defined yet.'

$imageReconModel = [pscustomobject]@{ diagramType = 'image-reconstruction'; sourceFormat = 'image-reconstruction'; nodes = @(); edges = @() }
$mermaidModel = [pscustomobject]@{ diagramType = 'flowchart'; sourceFormat = 'mermaid'; nodes = @(); edges = @() }
$drawioModel = [pscustomobject]@{ diagramType = 'flowchart'; sourceFormat = 'drawio'; nodes = @(); edges = @() }
$networkModel = [pscustomobject]@{ diagramType = 'network'; nodes = @(); edges = @() }
$blankFlowchartModel = [pscustomobject]@{ diagramType = 'flowchart'; nodes = @(); edges = @() }

$imageReconIntent = Resolve-VisioRoutingIntent -GraphModel $imageReconModel
Assert-Condition -Condition ($imageReconIntent -eq 'fidelity') -Message "Image reconstruction should resolve to fidelity, got $imageReconIntent"

$mermaidIntent = Resolve-VisioRoutingIntent -GraphModel $mermaidModel
Assert-Condition -Condition ($mermaidIntent -eq 'balanced') -Message "Mermaid should resolve to balanced, got $mermaidIntent"

$drawioIntent = Resolve-VisioRoutingIntent -GraphModel $drawioModel
Assert-Condition -Condition ($drawioIntent -eq 'balanced') -Message "DrawIO should resolve to balanced, got $drawioIntent"

$networkIntent = Resolve-VisioRoutingIntent -GraphModel $networkModel
Assert-Condition -Condition ($networkIntent -eq 'clean') -Message "Network should resolve to clean, got $networkIntent"

$flowchartIntent = Resolve-VisioRoutingIntent -GraphModel $blankFlowchartModel
Assert-Condition -Condition ($flowchartIntent -eq 'balanced') -Message "Plain flowchart should resolve to balanced, got $flowchartIntent"

$overrideFidelity = Resolve-VisioRoutingIntent -GraphModel $networkModel -OverrideIntent 'fidelity'
Assert-Condition -Condition ($overrideFidelity -eq 'fidelity') -Message "Override should force fidelity intent."

$declaredIntent = Resolve-VisioRoutingIntent -GraphModel ([pscustomobject]@{ diagramType = 'flowchart'; routingIntent = 'clean'; nodes = @(); edges = @() })
Assert-Condition -Condition ($declaredIntent -eq 'clean') -Message "Declared routingIntent on graph should be respected."

$results.routingIntentTest = [ordered]@{
  passed = $true
  imageReconstruction = $imageReconIntent
  mermaid = $mermaidIntent
  drawio = $drawioIntent
  network = $networkIntent
  flowchart = $flowchartIntent
  overrideFidelity = $overrideFidelity
  declaredClean = $declaredIntent
}

# Test 6: routing policy — same Graph Model produces different route decisions under different intents.
Assert-Condition -Condition ($null -ne (Get-Command Resolve-VisioRouteDecision -ErrorAction SilentlyContinue)) -Message 'Resolve-VisioRouteDecision is not defined yet.'

$policyGraph = [pscustomobject]@{
  diagramType = 'flowchart'
  layout = [pscustomobject]@{ strategy = 'Flowchart'; direction = 'LR' }
  nodes = @(
    [pscustomobject]@{ id = 'start'; text = '开始'; semanticType = 'start'; stencil = 'BASFLO_M.VSSX'; x = 1.0; y = 3.0 }
    [pscustomobject]@{ id = 'process'; text = '处理'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 3.0; y = 3.0 }
    [pscustomobject]@{ id = 'decision'; text = '通过?'; semanticType = 'decision'; stencil = 'BASFLO_M.VSSX'; x = 5.0; y = 3.0 }
    [pscustomobject]@{ id = 'end'; text = '结束'; semanticType = 'end'; stencil = 'BASFLO_M.VSSX'; x = 7.0; y = 3.0 }
    [pscustomobject]@{ id = 'reject'; text = '退回'; semanticType = 'process'; stencil = 'BASFLO_M.VSSX'; x = 5.0; y = 1.5 }
  )
  edges = @(
    [pscustomobject]@{ from = 'start'; to = 'process'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'process'; to = 'decision'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'decision'; to = 'end'; routeType = 'orthogonal'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'decision'; to = 'reject'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'reject'; to = 'process'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
  )
}

$fidelityDecision = Resolve-VisioRouteDecision -Edge ($policyGraph.edges[4]) -GraphModel $policyGraph -RoutingIntent 'fidelity'
$balancedDecision = Resolve-VisioRouteDecision -Edge ($policyGraph.edges[4]) -GraphModel $policyGraph -RoutingIntent 'balanced'
$cleanDecision = Resolve-VisioRouteDecision -Edge ($policyGraph.edges[4]) -GraphModel $policyGraph -RoutingIntent 'clean'

Assert-Condition -Condition ($fidelityDecision.isLoopback -eq $true) -Message 'Reject->process should be detected as loopback.'
Assert-Condition -Condition ($balancedDecision.isLoopback -eq $true) -Message 'Reject->process should be detected as loopback in balanced.'
Assert-Condition -Condition ($cleanDecision.isLoopback -eq $true) -Message 'Reject->process should be detected as loopback in clean.'

Assert-Condition -Condition ($balancedDecision.routeType -eq 'curved') -Message "Balanced loopback should default to curved, got $($balancedDecision.routeType)"
Assert-Condition -Condition ($cleanDecision.routeType -eq 'curved') -Message "Clean loopback should default to curved for flowchart, got $($cleanDecision.routeType)"

$decisionEdgeNoType = $policyGraph.edges[3]
$fidelityDecisionEdge = Resolve-VisioRouteDecision -Edge $decisionEdgeNoType -GraphModel $policyGraph -RoutingIntent 'fidelity'
$balancedDecisionEdge = Resolve-VisioRouteDecision -Edge $decisionEdgeNoType -GraphModel $policyGraph -RoutingIntent 'balanced'
$cleanDecisionEdge = Resolve-VisioRouteDecision -Edge $decisionEdgeNoType -GraphModel $policyGraph -RoutingIntent 'clean'

Assert-Condition -Condition ($fidelityDecisionEdge.routeType -eq 'orthogonal') -Message "Fidelity default for flowchart should be orthogonal, got $($fidelityDecisionEdge.routeType)"
Assert-Condition -Condition ($balancedDecisionEdge.routeType -eq 'orthogonal') -Message "Balanced default for flowchart should be orthogonal, got $($balancedDecisionEdge.routeType)"
Assert-Condition -Condition ($cleanDecisionEdge.routeType -eq 'orthogonal') -Message "Clean default for flowchart should be orthogonal, got $($cleanDecisionEdge.routeType)"

$explicitStraightEdge = [pscustomobject]@{ from = 'start'; to = 'process'; routeType = 'straight'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
$cleanExplicit = Resolve-VisioRouteDecision -Edge $explicitStraightEdge -GraphModel $policyGraph -RoutingIntent 'clean'
Assert-Condition -Condition ($cleanExplicit.routeType -eq 'straight') -Message "Explicit straight routeType should be preserved even under clean intent, got $($cleanExplicit.routeType)"

$explicitCurvedEdge = [pscustomobject]@{ from = 'reject'; to = 'process'; routeType = 'curved'; fromX = 0.5; fromY = 0.0; toX = 0.5; toY = 1.0; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
$cleanCurved = Resolve-VisioRouteDecision -Edge $explicitCurvedEdge -GraphModel $policyGraph -RoutingIntent 'clean'
Assert-Condition -Condition ($cleanCurved.routeType -eq 'curved') -Message "Explicit curved loopback should be preserved under clean intent."
Assert-Condition -Condition ($cleanCurved.fromX -eq 0.5 -and $cleanCurved.fromY -eq 0.0) -Message "Explicit glue points should be preserved."

$noRouteEdge = [pscustomobject]@{ from = 'start'; to = 'process'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
$cleanNoRoute = Resolve-VisioRouteDecision -Edge $noRouteEdge -GraphModel $policyGraph -RoutingIntent 'clean'
Assert-Condition -Condition ($cleanNoRoute.routeType -eq 'orthogonal') -Message "No routeType with clean intent should default to orthogonal, got $($cleanNoRoute.routeType)"

$networkGraph = [pscustomobject]@{
  diagramType = 'network'
  layout = [pscustomobject]@{ strategy = 'Hierarchical'; direction = 'LR' }
  nodes = @(
    [pscustomobject]@{ id = 'a'; text = 'A'; semanticType = 'process'; x = 1.0; y = 3.0 }
    [pscustomobject]@{ id = 'b'; text = 'B'; semanticType = 'process'; x = 3.0; y = 3.0 }
    [pscustomobject]@{ id = 'c'; text = 'C'; semanticType = 'process'; x = 3.0; y = 1.5 }
  )
  edges = @(
    [pscustomobject]@{ from = 'a'; to = 'b'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'a'; to = 'c'; routeType = 'straight'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
  )
}
$networkDefaults = Get-VisioRoutingDefault -DiagramType 'network'
Assert-Condition -Condition ($networkDefaults.defaultRouteType -eq 'orthogonal') -Message "Network default should be orthogonal"
Assert-Condition -Condition ($networkDefaults.loopbackRouteType -eq 'orthogonal') -Message "Network loopback should be orthogonal"

$decisionNetworkFidelity = Resolve-VisioRouteDecision -Edge ($networkGraph.edges[1]) -GraphModel $networkGraph -RoutingIntent 'fidelity'
Assert-Condition -Condition ($decisionNetworkFidelity.routeType -eq 'straight') -Message "Fidelity should preserve explicit straight routeType on network edge."

$results.routingPolicyTest = [ordered]@{
  passed = $true
  loopbackDetected = $fidelityDecision.isLoopback
  balancedLoopbackRoute = $balancedDecision.routeType
  cleanLoopbackRoute = $cleanDecision.routeType
  fidelityDefaultRoute = $fidelityDecisionEdge.routeType
  explicitStraightPreservedInClean = $cleanExplicit.routeType
  explicitCurvedPreservedInClean = $cleanCurved.routeType
  explicitGluePreserved = "$($cleanCurved.fromX),$($cleanCurved.fromY)"
  noRouteCleanDefault = $cleanNoRoute.routeType
  networkDefault = $networkDefaults.defaultRouteType
  networkExplicitStraightPreserved = $decisionNetworkFidelity.routeType
}

$jsonPath = Join-Path $OutputDirectory 'test-results.json'
($results | ConvertTo-Json -Depth 8) | Set-Content -Path $jsonPath -Encoding utf8
$results.files += $jsonPath

$results | ConvertTo-Json -Depth 8
