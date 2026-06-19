param(
  [string] $OutputDirectory = ''
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'skills\visio-automation\scripts\visio_helpers.ps1')

function Assert-Condition {
  param(
    [Parameter(Mandatory = $true)] [bool] $Condition,
    [Parameter(Mandatory = $true)] [string] $Message
  )

  if (-not $Condition) {
    throw $Message
  }
}

if (-not $OutputDirectory) {
  $OutputDirectory = Join-Path $repoRoot 'out\routing-policy-test'
}
$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

$flowGraph = [pscustomobject]@{
  diagramType = 'flowchart'
  layout = [pscustomobject]@{ strategy = 'Flowchart'; direction = 'LR' }
  nodes = @(
    [pscustomobject]@{ id = 'start'; text = '发起申请'; semanticType = 'start'; x = 1.1; y = 3.5; width = 1.6; height = 0.58 }
    [pscustomobject]@{ id = 'manager'; text = '部门审批'; semanticType = 'process'; x = 3.1; y = 3.5; width = 1.55; height = 0.62 }
    [pscustomobject]@{ id = 'decision'; text = '是否通过?'; semanticType = 'decision'; x = 5.1; y = 3.5; width = 1.35; height = 0.88 }
    [pscustomobject]@{ id = 'revise'; text = '返回修改'; semanticType = 'document'; x = 5.1; y = 2.0; width = 1.45; height = 0.62 }
    [pscustomobject]@{ id = 'end'; text = '流程结束'; semanticType = 'end'; x = 7.1; y = 3.5; width = 1.35; height = 0.58 }
  )
  edges = @(
    [pscustomobject]@{ from = 'start'; to = 'manager'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'manager'; to = 'decision'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'decision'; to = 'end'; text = '通过'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'decision'; to = 'revise'; text = '拒绝'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'revise'; to = 'manager'; text = '重新提交'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
  )
}

$forwardDecision = Resolve-VisioRouteDecision -Edge $flowGraph.edges[0] -GraphModel $flowGraph -RoutingIntent 'balanced'
Assert-Condition -Condition ($forwardDecision.fromSide -eq 'right' -and $forwardDecision.toSide -eq 'left') -Message "Forward LR edge should use right->left ports, got $($forwardDecision.fromSide)->$($forwardDecision.toSide)."
Assert-Condition -Condition ($forwardDecision.fromX -eq 1.0 -and $forwardDecision.fromY -eq 0.5 -and $forwardDecision.toX -eq 0.0 -and $forwardDecision.toY -eq 0.5) -Message 'Forward LR edge glue points are not right->left.'

$rejectDecision = Resolve-VisioRouteDecision -Edge $flowGraph.edges[3] -GraphModel $flowGraph -RoutingIntent 'balanced'
Assert-Condition -Condition ($rejectDecision.fromSide -eq 'bottom' -and $rejectDecision.toSide -eq 'top') -Message "Downward rejection edge should use bottom->top ports, got $($rejectDecision.fromSide)->$($rejectDecision.toSide)."

$loopDecision = Resolve-VisioRouteDecision -Edge $flowGraph.edges[4] -GraphModel $flowGraph -RoutingIntent 'balanced'
Assert-Condition -Condition ($loopDecision.isLoopback -eq $true) -Message 'Revise->manager should be detected as a loopback.'
Assert-Condition -Condition ($loopDecision.routeType -eq 'curved') -Message "Loopback should keep curved route type, got $($loopDecision.routeType)."
Assert-Condition -Condition ($loopDecision.fromSide -eq 'left' -and $loopDecision.toSide -eq 'bottom') -Message "Loopback should choose source-left -> target-bottom ports instead of an outer detour, got $($loopDecision.fromSide)->$($loopDecision.toSide)."
Assert-Condition -Condition ($loopDecision.fromX -eq 0.0 -and $loopDecision.fromY -eq 0.5 -and $loopDecision.toX -eq 0.5 -and $loopDecision.toY -eq 0.0) -Message 'Loopback glue points should be source left to target bottom.'

$explicitEdge = [pscustomobject]@{ from = 'revise'; to = 'manager'; routeType = 'orthogonal'; fromX = 0.0; fromY = 0.5; toX = 1.0; toY = 0.5; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
$explicitDecision = Resolve-VisioRouteDecision -Edge $explicitEdge -GraphModel $flowGraph -RoutingIntent 'balanced'
Assert-Condition -Condition ($explicitDecision.fromX -eq 0.0 -and $explicitDecision.fromY -eq 0.5 -and $explicitDecision.toX -eq 1.0 -and $explicitDecision.toY -eq 0.5) -Message 'Explicit glue points should still be preserved.'

$blockedGraph = [pscustomobject]@{
  diagramType = 'flowchart'
  layout = [pscustomobject]@{ strategy = 'Flowchart'; direction = 'LR' }
  nodes = @(
    [pscustomobject]@{ id = 'left'; text = '左侧步骤'; semanticType = 'process'; x = 1.0; y = 3.0; width = 1.4; height = 0.6 }
    [pscustomobject]@{ id = 'blocker'; text = '中间障碍'; semanticType = 'process'; x = 3.0; y = 3.0; width = 1.2; height = 0.9 }
    [pscustomobject]@{ id = 'right'; text = '右侧步骤'; semanticType = 'process'; x = 5.0; y = 3.0; width = 1.4; height = 0.6 }
  )
  edges = @(
    [pscustomobject]@{ from = 'left'; to = 'right'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
  )
}

$blockedDecision = Resolve-VisioRouteDecision -Edge $blockedGraph.edges[0] -GraphModel $blockedGraph -RoutingIntent 'clean'
Assert-Condition -Condition ($blockedDecision.obstacleHits -eq 0) -Message "Blocked edge should choose a port pair whose candidate path avoids the middle node, got $($blockedDecision.fromSide)->$($blockedDecision.toSide) with $($blockedDecision.obstacleHits) obstacle hits."
Assert-Condition -Condition ($blockedDecision.avoidancePenalty -eq 0) -Message "Blocked edge should not carry an avoidance penalty after choosing a clear candidate path, got $($blockedDecision.avoidancePenalty)."
Assert-Condition -Condition ($blockedDecision.routeLength -gt 0) -Message 'Blocked edge should expose routeLength diagnostics.'

$loopbackBlockedGraph = [pscustomobject]@{
  diagramType = 'flowchart'
  layout = [pscustomobject]@{ strategy = 'Flowchart'; direction = 'LR' }
  nodes = @(
    [pscustomobject]@{ id = 'approve'; text = '审批'; semanticType = 'process'; x = 3.0; y = 3.4; width = 1.5; height = 0.6 }
    [pscustomobject]@{ id = 'revise'; text = '修改'; semanticType = 'process'; x = 5.2; y = 2.0; width = 1.5; height = 0.6 }
    [pscustomobject]@{ id = 'blocker'; text = '备注'; semanticType = 'document'; x = 4.1; y = 2.7; width = 1.1; height = 0.8 }
  )
  edges = @(
    [pscustomobject]@{ from = 'revise'; to = 'approve'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
  )
}

$loopbackBlockedDecision = Resolve-VisioRouteDecision -Edge $loopbackBlockedGraph.edges[0] -GraphModel $loopbackBlockedGraph -RoutingIntent 'balanced'
Assert-Condition -Condition ($loopbackBlockedDecision.isLoopback -eq $true) -Message 'Blocked loopback edge should still be detected as loopback.'
Assert-Condition -Condition ($loopbackBlockedDecision.obstacleHits -eq 0) -Message "Blocked loopback edge should choose a candidate path that avoids the middle obstacle, got $($loopbackBlockedDecision.fromSide)->$($loopbackBlockedDecision.toSide) with $($loopbackBlockedDecision.obstacleHits) obstacle hits."

$crossingGraph = [pscustomobject]@{
  diagramType = 'flowchart'
  layout = [pscustomobject]@{ strategy = 'Flowchart'; direction = 'LR' }
  nodes = @(
    [pscustomobject]@{ id = 'a'; text = 'A'; semanticType = 'process'; x = 1.0; y = 4.0; width = 1.2; height = 0.6 }
    [pscustomobject]@{ id = 'b'; text = 'B'; semanticType = 'process'; x = 5.0; y = 4.0; width = 1.2; height = 0.6 }
    [pscustomobject]@{ id = 'c'; text = 'C'; semanticType = 'process'; x = 3.0; y = 5.6; width = 1.2; height = 0.6 }
    [pscustomobject]@{ id = 'd'; text = 'D'; semanticType = 'process'; x = 3.0; y = 2.4; width = 1.2; height = 0.6 }
  )
  edges = @(
    [pscustomobject]@{ from = 'a'; to = 'b'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
    [pscustomobject]@{ from = 'c'; to = 'd'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
  )
}

$existingCrossRoute = [pscustomobject]@{
  from = 'a'
  to = 'b'
  points = @(
    [pscustomobject]@{ x = 1.6; y = 4.0 }
    [pscustomobject]@{ x = 4.4; y = 4.0 }
  )
}

$crossingDecision = Resolve-VisioRouteDecision -Edge $crossingGraph.edges[1] -GraphModel $crossingGraph -RoutingIntent 'clean' -ExistingRoutes @($existingCrossRoute)
Assert-Condition -Condition ($crossingDecision.crossingHits -eq 0) -Message "Crossing-aware edge should choose a candidate route with no crossings, got $($crossingDecision.fromSide)->$($crossingDecision.toSide) with $($crossingDecision.crossingHits) crossings."
Assert-Condition -Condition ($crossingDecision.crossingPenalty -eq 0) -Message "Crossing-aware edge should not carry a crossing penalty after choosing a clear route, got $($crossingDecision.crossingPenalty)."

$labelBlockedGraph = [pscustomobject]@{
  diagramType = 'flowchart'
  layout = [pscustomobject]@{ strategy = 'Flowchart'; direction = 'LR' }
  nodes = @(
    [pscustomobject]@{ id = 'source'; text = '提交'; semanticType = 'process'; x = 1.0; y = 3.0; width = 1.3; height = 0.6 }
    [pscustomobject]@{ id = 'target'; text = '审批'; semanticType = 'process'; x = 5.0; y = 3.0; width = 1.3; height = 0.6 }
  )
  annotations = @(
    [pscustomobject]@{ id = 'note'; text = '优先级说明'; x = 3.0; y = 3.0; width = 1.4; height = 0.35 }
  )
  edges = @(
    [pscustomobject]@{ from = 'source'; to = 'target'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
  )
}

$labelBlockedDecision = Resolve-VisioRouteDecision -Edge $labelBlockedGraph.edges[0] -GraphModel $labelBlockedGraph -RoutingIntent 'clean'
Assert-Condition -Condition ($labelBlockedDecision.labelHits -eq 0) -Message "Label-aware edge should avoid text label bounds, got $($labelBlockedDecision.fromSide)->$($labelBlockedDecision.toSide) with $($labelBlockedDecision.labelHits) label hits."
Assert-Condition -Condition ($labelBlockedDecision.labelPenalty -eq 0) -Message "Label-aware edge should not carry a label penalty after choosing a clear candidate path, got $($labelBlockedDecision.labelPenalty)."

$boundaryBlockedGraph = [pscustomobject]@{
  diagramType = 'workflow'
  layout = [pscustomobject]@{ strategy = 'Flowchart'; direction = 'LR' }
  nodes = @(
    [pscustomobject]@{ id = 'left'; text = '申请'; semanticType = 'process'; x = 1.0; y = 3.0; width = 1.2; height = 0.6 }
    [pscustomobject]@{ id = 'right'; text = '归档'; semanticType = 'process'; x = 5.2; y = 3.0; width = 1.2; height = 0.6 }
  )
  swimlanes = @(
    [pscustomobject]@{ id = 'finance'; text = '财务泳道'; x = 3.1; y = 3.0; width = 1.3; height = 1.8 }
  )
  edges = @(
    [pscustomobject]@{ from = 'left'; to = 'right'; connector = 'Dynamic connector'; stencil = 'BASFLO_M.VSSX' }
  )
}

$boundaryBlockedDecision = Resolve-VisioRouteDecision -Edge $boundaryBlockedGraph.edges[0] -GraphModel $boundaryBlockedGraph -RoutingIntent 'clean'
Assert-Condition -Condition ($boundaryBlockedDecision.boundaryHits -eq 0) -Message "Boundary-aware edge should avoid unrelated swimlane/container bounds, got $($boundaryBlockedDecision.fromSide)->$($boundaryBlockedDecision.toSide) with $($boundaryBlockedDecision.boundaryHits) boundary hits."
Assert-Condition -Condition ($boundaryBlockedDecision.boundaryPenalty -eq 0) -Message "Boundary-aware edge should not carry a boundary penalty after choosing a clear candidate path, got $($boundaryBlockedDecision.boundaryPenalty)."

$result = [pscustomobject]@{
  passed = $true
  forward = $forwardDecision
  rejection = $rejectDecision
  loopback = $loopDecision
  explicit = $explicitDecision
  blocked = $blockedDecision
  loopbackBlocked = $loopbackBlockedDecision
  crossing = $crossingDecision
  labelBlocked = $labelBlockedDecision
  boundaryBlocked = $boundaryBlockedDecision
}

$jsonPath = Join-Path $OutputDirectory 'routing-policy-result.json'
$result | ConvertTo-Json -Depth 8 | Set-Content -Path $jsonPath -Encoding utf8
$result | ConvertTo-Json -Depth 8
