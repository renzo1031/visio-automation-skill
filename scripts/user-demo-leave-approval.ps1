param(
  [string] $OutputDirectory = ''
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'skills\visio-automation\scripts\visio_helpers.ps1')

if (-not $OutputDirectory) {
  $OutputDirectory = Join-Path $repoRoot 'out\user-demo'
}
$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

$outputPath = Join-Path $OutputDirectory 'leave-approval-flow.vsdx'
$previewPath = Join-Path $OutputDirectory 'leave-approval-flow.png'

$graph = [pscustomobject]@{
  diagramType = 'flowchart'
  routingIntent = 'balanced'
  layout = [pscustomobject]@{
    strategy = 'Flowchart'
    direction = 'LR'
  }
  page = [pscustomobject]@{
    width = 9.0
    height = 5.0
  }
  nodes = @(
    [pscustomobject]@{
      id = 'start'
      text = '发起申请'
      semanticType = 'start'
      stencil = 'BASFLO_M.VSSX'
      x = 1.1
      y = 3.5
      width = 1.6
      height = 0.58
    }
    [pscustomobject]@{
      id = 'manager'
      text = '部门审批'
      semanticType = 'process'
      stencil = 'BASFLO_M.VSSX'
      x = 3.1
      y = 3.5
      width = 1.55
      height = 0.62
    }
    [pscustomobject]@{
      id = 'decision'
      text = '是否通过?'
      semanticType = 'decision'
      stencil = 'BASFLO_M.VSSX'
      x = 5.1
      y = 3.5
      width = 1.35
      height = 0.88
    }
    [pscustomobject]@{
      id = 'revise'
      text = '返回修改'
      semanticType = 'document'
      stencil = 'BASFLO_M.VSSX'
      x = 5.1
      y = 2.0
      width = 1.45
      height = 0.62
    }
    [pscustomobject]@{
      id = 'end'
      text = '流程结束'
      semanticType = 'end'
      stencil = 'BASFLO_M.VSSX'
      x = 7.1
      y = 3.5
      width = 1.35
      height = 0.58
    }
  )
  edges = @(
    [pscustomobject]@{
      from = 'start'
      to = 'manager'
      routeType = 'orthogonal'
      connector = 'Dynamic connector'
      stencil = 'BASFLO_M.VSSX'
    }
    [pscustomobject]@{
      from = 'manager'
      to = 'decision'
      routeType = 'orthogonal'
      connector = 'Dynamic connector'
      stencil = 'BASFLO_M.VSSX'
    }
    [pscustomobject]@{
      from = 'decision'
      to = 'end'
      text = '通过'
      connector = 'Dynamic connector'
      stencil = 'BASFLO_M.VSSX'
    }
    [pscustomobject]@{
      from = 'decision'
      to = 'revise'
      text = '拒绝'
      connector = 'Dynamic connector'
      stencil = 'BASFLO_M.VSSX'
    }
    [pscustomobject]@{
      from = 'revise'
      to = 'manager'
      text = '重新提交'
      connector = 'Dynamic connector'
      stencil = 'BASFLO_M.VSSX'
    }
  )
}

$result = Render-VisioGraphModel -GraphModel $graph -OutputPath $outputPath -PreviewPath $previewPath -Force

Update-VisioSessionCache `
  -WorkspacePath $repoRoot `
  -DiagramType 'flowchart' `
  -Stencil 'BASFLO_M.VSSX' `
  -LastOutputPath $outputPath `
  -LastPreviewPath $previewPath `
  -LastLayoutStrategy 'Manual' `
  -LastLayoutDirection 'LR' | Out-Null

Add-VisioSessionRecentMaster -WorkspacePath $repoRoot -Stencil 'BASFLO_M.VSSX' -MasterNameU 'Start/End' | Out-Null
Add-VisioSessionRecentMaster -WorkspacePath $repoRoot -Stencil 'BASFLO_M.VSSX' -MasterNameU 'Process' | Out-Null
Add-VisioSessionRecentMaster -WorkspacePath $repoRoot -Stencil 'BASFLO_M.VSSX' -MasterNameU 'Decision' | Out-Null
Add-VisioSessionRecentMaster -WorkspacePath $repoRoot -Stencil 'BASFLO_M.VSSX' -MasterNameU 'Document' | Out-Null

[pscustomobject]@{
  OutputPath = $outputPath
  PreviewPath = $previewPath
  Nodes = @($result.nodes).Count
  Edges = @($result.edges).Count
  RoutingIntent = $result.routingIntent
  RouteTypes = @($result.edges | ForEach-Object { $_.routeType })
  EdgesDetail = @($result.edges | ForEach-Object {
    [pscustomobject]@{
      from = $_.from
      to = $_.to
      text = $_.text
      routeType = $_.routeType
      isLoopback = $_.isLoopback
      fromSide = $_.fromSide
      toSide = $_.toSide
      routeScore = $_.routeScore
      routeLength = $_.routeLength
      obstacleHits = $_.obstacleHits
      avoidancePenalty = $_.avoidancePenalty
      labelHits = $_.labelHits
      labelPenalty = $_.labelPenalty
      boundaryHits = $_.boundaryHits
      boundaryPenalty = $_.boundaryPenalty
      crossingHits = $_.crossingHits
      crossingPenalty = $_.crossingPenalty
      fromX = $_.fromX
      fromY = $_.fromY
      toX = $_.toX
      toY = $_.toY
    }
  })
} | ConvertTo-Json -Depth 5
