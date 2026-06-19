param(
  [string] $OutputDirectory = '',
  [string] $Text = '帮我画一个请假审批流程：员工发起申请，部门经理审批，如果通过就流程结束，如果不通过就返回修改。',
  [switch] $Invisible
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'skills\visio-automation\scripts\visio_helpers.ps1')

if (-not $OutputDirectory) {
  $OutputDirectory = Join-Path $repoRoot 'out\user-demo-natural-language-flow'
}
$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

$outputPath = Join-Path $OutputDirectory 'natural-language-flow.vsdx'
$previewPath = Join-Path $OutputDirectory 'natural-language-flow.png'
$normalizedTextFlow = Convert-NaturalLanguageFlowToTextFlow -Text $Text

$renderArgs = @{
  Text = $Text
  OutputPath = $outputPath
  PreviewPath = $previewPath
  LayoutDirection = 'LR'
  RoutingIntent = 'balanced'
  Force = $true
}
if ($Invisible) {
  $renderArgs.Invisible = $true
}

$result = Convert-NaturalLanguageFlowToVisio @renderArgs

Update-VisioSessionCache `
  -WorkspacePath $repoRoot `
  -DiagramType 'flowchart' `
  -Stencil 'BASFLO_M.VSSX' `
  -LastOutputPath $outputPath `
  -LastPreviewPath $previewPath `
  -LastLayoutStrategy 'Flowchart' `
  -LastLayoutDirection 'LR' | Out-Null

[pscustomobject]@{
  InputText = $Text
  NormalizedTextFlow = $normalizedTextFlow
  OutputPath = $outputPath
  PreviewPath = $previewPath
  Nodes = @($result.nodes).Count
  Edges = @($result.edges).Count
  RoutingIntent = $result.routingIntent
  NodeTypes = @($result.nodes | ForEach-Object {
    [pscustomobject]@{
      id = $_.id
      text = $_.text
      semanticType = $_.semanticType
      master = $_.master
    }
  })
  EdgesDetail = @($result.edges | ForEach-Object {
    [pscustomobject]@{
      from = $_.from
      to = $_.to
      text = $_.text
      routeType = $_.routeType
      fromSide = $_.fromSide
      toSide = $_.toSide
      obstacleHits = $_.obstacleHits
      labelHits = $_.labelHits
      boundaryHits = $_.boundaryHits
      crossingHits = $_.crossingHits
    }
  })
} | ConvertTo-Json -Depth 6
