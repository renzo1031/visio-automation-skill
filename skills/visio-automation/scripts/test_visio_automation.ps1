param(
  [string] $OutputDirectory = (Join-Path $env:TEMP 'visio-automation-skill-test')
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'visio_helpers.ps1')

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

$results = [ordered]@{
  outputDirectory = $OutputDirectory
  helperLoaded = $true
  knownMasterTest = $null
  explicitStraightTest = $null
  recoveryDiscoveryTest = $null
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

  $doc.SaveAs($knownFile)
  $page.Export($knownPreview)

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
  $doc.Close()
} finally {
  $visio.Quit()
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
  $doc.SaveAs($straightFile)
  Assert-Condition -Condition ($straightConnector.CellsU('ShapeRouteStyle').ResultIU -eq 2) -Message 'Explicit straight connector did not set ShapeRouteStyle to 2.'
  Assert-Condition -Condition ($straightConnector.CellsU('ConLineRouteExt').ResultIU -eq 1) -Message 'Explicit straight connector did not set ConLineRouteExt to 1.'
  $results.explicitStraightTest = [ordered]@{
    passed = $true
    file = $straightFile
    shapeRouteStyle = $straightConnector.CellsU('ShapeRouteStyle').FormulaU
    conLineRouteExt = $straightConnector.CellsU('ConLineRouteExt').FormulaU
  }
  $results.files += $straightFile
  $doc.Close()
} finally {
  $visio.Quit()
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
  $doc.SaveAs($recoveryFile)
  $page.Export($recoveryPreview)

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
  $doc.Close()
} finally {
  $visio.Quit()
}

$jsonPath = Join-Path $OutputDirectory 'test-results.json'
($results | ConvertTo-Json -Depth 8) | Set-Content -Path $jsonPath -Encoding utf8
$results.files += $jsonPath

$results | ConvertTo-Json -Depth 8
