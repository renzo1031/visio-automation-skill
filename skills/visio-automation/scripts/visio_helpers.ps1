$ErrorActionPreference = 'Stop'

$script:VisioMasterIndexSchemaVersion = 1

function New-VisibleVisioApplication {
  $visio = New-Object -ComObject Visio.Application
  $visio.Visible = $true
  return $visio
}

function New-InvisibleVisioApplication {
  return New-Object -ComObject Visio.InvisibleApp
}

function Open-VisioStencilReadOnly {
  param(
    [Parameter(Mandatory = $true)] $Visio,
    [Parameter(Mandatory = $true)] [string] $StencilNameOrPath
  )

  # 66 = visOpenHidden (64) + visOpenRO (2). Keep stencils hidden and read-only.
  return $Visio.Documents.OpenEx($StencilNameOrPath, 66)
}

function Get-VisioContentRoots {
  $candidates = @(
    (Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\Visio Content'),
    (Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\root\Office16\Visio Content'),
    (Join-Path $env:ProgramFiles 'Microsoft Office\Office16\Visio Content'),
    (Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\Office16\Visio Content')
  )

  foreach ($candidate in $candidates) {
    if ($candidate -and (Test-Path $candidate)) {
      Get-Item -LiteralPath $candidate
    }
  }
}

function Get-VisioEnvironment {
  $visio = $null
  $version = $null
  $language = [System.Globalization.CultureInfo]::CurrentUICulture.Name
  try {
    $visio = New-InvisibleVisioApplication
    $version = [string] $visio.Version
    try {
      $languageId = [int] $visio.Language
      if ($languageId -gt 0) {
        $language = ([System.Globalization.CultureInfo]::GetCultureInfo($languageId)).Name
      }
    } catch {
      # Keep the current UI culture when Visio does not expose a usable language id.
    }
  } finally {
    if ($null -ne $visio) {
      $null = $visio.Quit()
    }
  }

  $contentRoots = @(Get-VisioContentRoots | ForEach-Object { $_.FullName })
  [pscustomobject]@{
    Version = $version
    Language = $language
    InstallRoots = @(
      (Join-Path $env:ProgramFiles 'Microsoft Office'),
      (Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
    ContentRoots = $contentRoots
    StencilPatterns = @('*.vssx', '*.vss')
  }
}

function Get-VisioIndexContentRootSummary {
  param(
    [Parameter(Mandatory = $true)] $Environment
  )

  $roots = @($Environment.ContentRoots)
  return ($roots | Sort-Object) -join '|'
}

function Find-VisioStencilFiles {
  param(
    [string[]] $Pattern = @('*.vssx', '*.vss')
  )

  foreach ($root in Get-VisioContentRoots) {
    foreach ($itemPattern in $Pattern) {
      Get-ChildItem -Path $root.FullName -Recurse -File -Filter $itemPattern -ErrorAction SilentlyContinue
    }
  }
}

function Find-VisioStencilFilesByName {
  param(
    [Parameter(Mandatory = $true)] [string] $NameRegex,
    [string[]] $Pattern = @('*.vssx', '*.vss')
  )

  foreach ($root in Get-VisioContentRoots) {
    foreach ($itemPattern in $Pattern) {
      Get-ChildItem -Path $root.FullName -File -Recurse -Filter $itemPattern -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match $NameRegex }
    }
  }
}

function Test-VisioMasterIndexValid {
  param(
    [Parameter(Mandatory = $true)] [string] $IndexPath,
    $Environment = (Get-VisioEnvironment)
  )

  if (-not (Test-Path -LiteralPath $IndexPath)) {
    return $false
  }

  try {
    $index = Get-Content -Raw -LiteralPath $IndexPath | ConvertFrom-Json
  } catch {
    return $false
  }

  if ([int] $index.SchemaVersion -ne $script:VisioMasterIndexSchemaVersion) {
    return $false
  }

  if ([string] $index.Environment.Version -ne [string] $Environment.Version) {
    return $false
  }

  if ([string] $index.Environment.ContentRootSummary -ne (Get-VisioIndexContentRootSummary -Environment $Environment)) {
    return $false
  }

  foreach ($stencil in @($index.Stencils)) {
    if (-not (Test-Path -LiteralPath $stencil.FullName)) {
      return $false
    }
    $file = Get-Item -LiteralPath $stencil.FullName
    if ([string] $file.LastWriteTimeUtc.Ticks -ne [string] $stencil.LastWriteTimeUtcTicks) {
      return $false
    }
  }

  return $true
}

function Build-VisioMasterIndex {
  param(
    [Parameter(Mandatory = $true)] [string] $OutputPath,
    [string[]] $StencilPattern = @('*.vssx', '*.vss'),
    [string] $PreferredStencilRegex = '',
    [string] $Query = '',
    [switch] $Force
  )

  $environment = Get-VisioEnvironment
  if (-not $Force -and (Test-VisioMasterIndexValid -IndexPath $OutputPath -Environment $environment)) {
    return Get-Content -Raw -LiteralPath $OutputPath | ConvertFrom-Json
  }

  $outputDirectory = Split-Path -Parent $OutputPath
  if ($outputDirectory) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
  }

  $files = @(Find-VisioStencilFiles -Pattern $StencilPattern)
  if ($PreferredStencilRegex) {
    $files = @($files | Where-Object { $_.Name -match $PreferredStencilRegex })
  }

  $visio = New-InvisibleVisioApplication
  $stencils = New-Object System.Collections.ArrayList
  $masters = New-Object System.Collections.ArrayList

  try {
    foreach ($file in $files) {
      try {
        $stencil = Open-VisioStencilReadOnly -Visio $visio -StencilNameOrPath $file.FullName
        $languageDirectory = Split-Path -Path $file.DirectoryName -Leaf
        $null = $stencils.Add([pscustomobject]@{
          Stencil = $file.Name
          FullName = $file.FullName
          Directory = $file.DirectoryName
          LanguageDirectory = $languageDirectory
          LastWriteTimeUtcTicks = [string] $file.LastWriteTimeUtc.Ticks
          Length = $file.Length
        })

        foreach ($master in @($stencil.Masters)) {
          if ($Query -and -not ($master.Name -match $Query -or $master.NameU -match $Query -or $file.Name -match $Query)) {
            continue
          }
          $null = $masters.Add([pscustomobject]@{
            Stencil = $file.Name
            FullName = $file.FullName
            Name = $master.Name
            NameU = $master.NameU
            LanguageDirectory = $languageDirectory
          })
        }
        $null = $stencil.Close()
      } catch {
        # Skip protected or incompatible Visio content files.
      }
    }
  } finally {
    $null = $visio.Quit()
  }

  $contentRootSummary = Get-VisioIndexContentRootSummary -Environment $environment

  $index = [pscustomobject]@{
    SchemaVersion = $script:VisioMasterIndexSchemaVersion
    GeneratedAt = (Get-Date).ToUniversalTime().ToString('o')
    Environment = [pscustomobject]@{
      Version = $environment.Version
      Language = $environment.Language
      ContentRoots = @($environment.ContentRoots)
      ContentRootSummary = $contentRootSummary
      StencilPatterns = @($StencilPattern)
    }
    Stencils = @($stencils.ToArray())
    Masters = @($masters.ToArray())
  }

  $index | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputPath -Encoding utf8
  return $index
}

function Get-VisioMasters {
  param(
    [Parameter(Mandatory = $true)] [string] $StencilNameOrPath
  )

  $visio = New-InvisibleVisioApplication
  try {
    $stencil = Open-VisioStencilReadOnly -Visio $visio -StencilNameOrPath $StencilNameOrPath
    foreach ($master in @($stencil.Masters)) {
      [pscustomobject]@{
        Stencil = $StencilNameOrPath
        Name = $master.Name
        NameU = $master.NameU
      }
    }
    $null = $stencil.Close()
  } finally {
    $null = $visio.Quit()
  }
}

function Find-VisioMasters {
  param(
    [Parameter(Mandatory = $true)] [string] $Query,
    [string[]] $StencilPattern = @('*.vssx', '*.vss'),
    [string] $PreferredStencilRegex = '',
    [string] $IndexPath = '',
    [int] $MaxResults = 50
  )

  $exactTerms = $Query -split '\|' | ForEach-Object { $_.Trim(' ', '^', '$') } | Where-Object { $_ }

  if ($IndexPath -and (Test-VisioMasterIndexValid -IndexPath $IndexPath)) {
    $index = Get-Content -Raw -LiteralPath $IndexPath | ConvertFrom-Json
    $indexedResults = foreach ($master in @($index.Masters)) {
      if ($PreferredStencilRegex -and $master.Stencil -notmatch $PreferredStencilRegex -and $master.FullName -notmatch $PreferredStencilRegex) {
        continue
      }
      if ($master.Name -match $Query -or $master.NameU -match $Query -or $master.Stencil -match $Query) {
        $score = 3
        foreach ($term in $exactTerms) {
          if ($master.NameU -ieq $term -or $master.Name -ieq $term) {
            $score = 0
            break
          }
        }
        if ($score -ne 0 -and ($master.NameU -match $Query -or $master.Name -match $Query)) {
          $score = 1
        } elseif ($score -ne 0 -and $master.Stencil -match $Query) {
          $score = 2
        }
        [pscustomobject]@{
          Score = $score
          Source = 'Index'
          Stencil = $master.Stencil
          FullName = $master.FullName
          Name = $master.Name
          NameU = $master.NameU
        }
      }
    }

    $ordered = @($indexedResults | Sort-Object Score, Stencil, NameU | Select-Object -First $MaxResults)
    if ($ordered.Count -gt 0) {
      return $ordered
    }
  }

  $visio = New-InvisibleVisioApplication
  $results = New-Object System.Collections.Generic.List[object]
  function Search-VisioMasterFiles {
    param(
      [Parameter(Mandatory = $true)] [array] $FilesToSearch
    )

    foreach ($file in $FilesToSearch) {
      try {
        $stencil = Open-VisioStencilReadOnly -Visio $visio -StencilNameOrPath $file.FullName
        foreach ($master in @($stencil.Masters)) {
          if ($master.Name -match $Query -or $master.NameU -match $Query -or $file.Name -match $Query) {
            $score = 3
            foreach ($term in $exactTerms) {
              if ($master.NameU -ieq $term -or $master.Name -ieq $term) {
                $score = 0
                break
              }
            }
            if ($score -ne 0 -and ($master.NameU -match $Query -or $master.Name -match $Query)) {
              $score = 1
            } elseif ($score -ne 0 -and $file.Name -match $Query) {
              $score = 2
            }
            $results.Add([pscustomobject]@{
              Score = $score
              Source = 'Scan'
              Stencil = $file.Name
              FullName = $file.FullName
              Name = $master.Name
              NameU = $master.NameU
            })
          }
        }
        $null = $stencil.Close()
      } catch {
        # Some Visio content files can be templates or protected files. Skip and continue discovery.
      }
    }
  }

  try {
    if ($PreferredStencilRegex) {
      $preferred = @(Find-VisioStencilFilesByName -NameRegex $PreferredStencilRegex -Pattern $StencilPattern)
      Search-VisioMasterFiles -FilesToSearch $preferred
      if ($results.Count -eq 0) {
        $files = @(Find-VisioStencilFiles -Pattern $StencilPattern)
        $other = @($files | Where-Object { $_.Name -notmatch $PreferredStencilRegex })
        Search-VisioMasterFiles -FilesToSearch $other
      }
    } else {
      $files = @(Find-VisioStencilFiles -Pattern $StencilPattern)
      Search-VisioMasterFiles -FilesToSearch $files
    }
  } finally {
    $null = $visio.Quit()
  }

  return $results | Sort-Object Score, Stencil, NameU | Select-Object -First $MaxResults
}

function Set-VisioTextStyle {
  param(
    [Parameter(Mandatory = $true)] $Shape,
    [string] $Color = 'RGB(45,45,45)',
    [double] $Size = 11,
    [int] $Bold = 0
  )

  $null = $Shape.CellsU('Char.Color').FormulaU = $Color
  $null = $Shape.CellsU('Char.Size').FormulaU = "$Size pt"
  $null = $Shape.CellsU('Char.Style').FormulaU = "$Bold"
  $null = $Shape.CellsU('Para.HorzAlign').FormulaU = '1'
  $null = $Shape.CellsU('VerticalAlign').FormulaU = '1'
}

function Set-VisioShapeFill {
  param(
    [Parameter(Mandatory = $true)] $Shape,
    [Parameter(Mandatory = $true)] [string] $Fill,
    [string] $Line = $Fill,
    [string] $LineWeight = '1 pt'
  )

  $null = $Shape.CellsU('FillForegnd').FormulaU = $Fill
  $null = $Shape.CellsU('LineColor').FormulaU = $Line
  $null = $Shape.CellsU('LineWeight').FormulaU = $LineWeight
}

function Set-VisioShapeBounds {
  param(
    [Parameter(Mandatory = $true)] $Shape,
    [double] $X,
    [double] $Y,
    [double] $Width,
    [double] $Height
  )

  if ($Width -gt 0) {
    $null = $Shape.CellsU('Width').FormulaU = "$Width in"
  }
  if ($Height -gt 0) {
    $null = $Shape.CellsU('Height').FormulaU = "$Height in"
  }
  $null = $Shape.CellsU('PinX').FormulaU = "$X in"
  $null = $Shape.CellsU('PinY').FormulaU = "$Y in"
}

function Get-VisioShapeCenter {
  param(
    [Parameter(Mandatory = $true)] $Shape
  )

  if ($null -ne $Shape -and $Shape.PSObject.Properties.Match('x').Count -gt 0 -and $Shape.PSObject.Properties.Match('y').Count -gt 0) {
    return [pscustomobject]@{
      X = [double] $Shape.x
      Y = [double] $Shape.y
    }
  }

  try {
    return [pscustomobject]@{
      X = [double] $Shape.CellsU('PinX').ResultIU
      Y = [double] $Shape.CellsU('PinY').ResultIU
    }
  } catch {
    throw 'Unable to resolve Visio shape center.'
  }
}

function Resolve-VisioConnectorGluePoints {
  param(
    [Parameter(Mandatory = $true)] $From,
    [Parameter(Mandatory = $true)] $To,
    [Nullable[double]] $FromX = $null,
    [Nullable[double]] $FromY = $null,
    [Nullable[double]] $ToX = $null,
    [Nullable[double]] $ToY = $null
  )

  $fromCenter = Get-VisioShapeCenter -Shape $From
  $toCenter = Get-VisioShapeCenter -Shape $To

  $deltaX = [Math]::Abs([double] $fromCenter.X - [double] $toCenter.X)
  $deltaY = [Math]::Abs([double] $fromCenter.Y - [double] $toCenter.Y)
  $preferVertical = $deltaY -ge $deltaX

  if ($preferVertical) {
    if ([double] $fromCenter.Y -ge [double] $toCenter.Y) {
      $resolved = [ordered]@{
        FromX = 0.5
        FromY = 0.0
        ToX = 0.5
        ToY = 1.0
      }
    } else {
      $resolved = [ordered]@{
        FromX = 0.5
        FromY = 1.0
        ToX = 0.5
        ToY = 0.0
      }
    }
  } else {
    if ([double] $fromCenter.X -le [double] $toCenter.X) {
      $resolved = [ordered]@{
        FromX = 1.0
        FromY = 0.5
        ToX = 0.0
        ToY = 0.5
      }
    } else {
      $resolved = [ordered]@{
        FromX = 0.0
        FromY = 0.5
        ToX = 1.0
        ToY = 0.5
      }
    }
  }

  if ($null -ne $FromX) { $resolved.FromX = [double] $FromX }
  if ($null -ne $FromY) { $resolved.FromY = [double] $FromY }
  if ($null -ne $ToX) { $resolved.ToX = [double] $ToX }
  if ($null -ne $ToY) { $resolved.ToY = [double] $ToY }

  return [pscustomobject]$resolved
}

function Connect-VisioShapesOrthogonal {
  param(
    [Parameter(Mandatory = $true)] $Page,
    [Parameter(Mandatory = $true)] $ConnectorMaster,
    [Parameter(Mandatory = $true)] $From,
    [Parameter(Mandatory = $true)] $To,
    [Nullable[double]] $FromX = $null,
    [Nullable[double]] $FromY = $null,
    [Nullable[double]] $ToX = $null,
    [Nullable[double]] $ToY = $null,
    [int] $EndArrow = 4,
    [string] $LineColor = 'RGB(45,45,45)'
  )

  $gluePoints = Resolve-VisioConnectorGluePoints -From $From -To $To -FromX $FromX -FromY $FromY -ToX $ToX -ToY $ToY
  $connector = $Page.Drop($ConnectorMaster, 0, 0)
  $null = $connector.CellsU('BeginX').GlueToPos($From, $gluePoints.FromX, $gluePoints.FromY)
  $null = $connector.CellsU('EndX').GlueToPos($To, $gluePoints.ToX, $gluePoints.ToY)
  # Default to Visio's orthogonal/right-angle dynamic connector behavior.
  if ($connector.CellExistsU('ShapeRouteStyle', 0) -ne 0) {
    $null = $connector.CellsU('ShapeRouteStyle').FormulaU = '0'
  }
  if ($connector.CellExistsU('ConLineRouteExt', 0) -ne 0) {
    $null = $connector.CellsU('ConLineRouteExt').FormulaU = '0'
  }
  $null = $connector.CellsU('LineColor').FormulaU = $LineColor
  $null = $connector.CellsU('LineWeight').FormulaU = '1 pt'
  $null = $connector.CellsU('EndArrow').FormulaU = "$EndArrow"
  return $connector
}

function Set-VisioConnectorStraight {
  param(
    [Parameter(Mandatory = $true)] $Connector
  )

  $null = $Connector.CellsU('ShapeRouteStyle').FormulaU = '2'
  $null = $Connector.CellsU('ConLineRouteExt').FormulaU = '1'
  return $Connector
}

function Set-VisioConnectorCurved {
  param(
    [Parameter(Mandatory = $true)] $Connector
  )

  if ($Connector.CellExistsU('ShapeRouteStyle', 0) -ne 0) {
    $null = $Connector.CellsU('ShapeRouteStyle').FormulaU = '0'
  }
  $null = $Connector.CellsU('ConLineRouteExt').FormulaU = '2'
  return $Connector
}

function Connect-VisioShapesStraight {
  param(
    [Parameter(Mandatory = $true)] $Page,
    [Parameter(Mandatory = $true)] $ConnectorMaster,
    [Parameter(Mandatory = $true)] $From,
    [Parameter(Mandatory = $true)] $To,
    [Nullable[double]] $FromX = $null,
    [Nullable[double]] $FromY = $null,
    [Nullable[double]] $ToX = $null,
    [Nullable[double]] $ToY = $null,
    [int] $EndArrow = 4,
    [string] $LineColor = 'RGB(45,45,45)'
  )

  $connector = Connect-VisioShapesOrthogonal -Page $Page -ConnectorMaster $ConnectorMaster -From $From -To $To -FromX $FromX -FromY $FromY -ToX $ToX -ToY $ToY -EndArrow $EndArrow -LineColor $LineColor
  Set-VisioConnectorStraight -Connector $connector | Out-Null
  return $connector
}

function Connect-VisioShapesCurved {
  param(
    [Parameter(Mandatory = $true)] $Page,
    [Parameter(Mandatory = $true)] $ConnectorMaster,
    [Parameter(Mandatory = $true)] $From,
    [Parameter(Mandatory = $true)] $To,
    [Nullable[double]] $FromX = $null,
    [Nullable[double]] $FromY = $null,
    [Nullable[double]] $ToX = $null,
    [Nullable[double]] $ToY = $null,
    [int] $EndArrow = 4,
    [string] $LineColor = 'RGB(45,45,45)'
  )

  $connector = Connect-VisioShapesOrthogonal -Page $Page -ConnectorMaster $ConnectorMaster -From $From -To $To -FromX $FromX -FromY $FromY -ToX $ToX -ToY $ToY -EndArrow $EndArrow -LineColor $LineColor
  Set-VisioConnectorCurved -Connector $connector | Out-Null
  return $connector
}

function Get-VisioGraphModelSemanticMaster {
  param(
    [Parameter(Mandatory = $true)] [string] $SemanticType
  )

  switch -Regex ($SemanticType) {
    '^(start|end|terminator)$' { return 'Start/End' }
    '^(decision|gateway)$' { return 'Decision' }
    '^(database|data-store|datastore)$' { return 'Database' }
    '^(document)$' { return 'Document' }
    default { return 'Process' }
  }
}

function Copy-VisioGraphModel {
  param(
    [Parameter(Mandatory = $true)] $GraphModel
  )

  $json = $GraphModel | ConvertTo-Json -Depth 20
  return $json | ConvertFrom-Json
}

function New-VisioGraphModelNode {
  param(
    [Parameter(Mandatory = $true)] [string] $Id,
    [string] $Text = '',
    [string] $SemanticType = '',
    [string] $Stencil = 'BASFLO_M.VSSX',
    [string] $Master = '',
    [string] $PreferredMaster = ''
  )

  $node = [ordered]@{
    id = $Id
    text = $Text
    stencil = $Stencil
  }
  if ($SemanticType) { $node.semanticType = $SemanticType }
  if ($Master) { $node.master = $Master }
  if ($PreferredMaster) { $node.preferredMaster = $PreferredMaster }
  return [pscustomobject]$node
}

function New-VisioGraphModelEdge {
  param(
    [Parameter(Mandatory = $true)] [string] $From,
    [Parameter(Mandatory = $true)] [string] $To,
    [string] $Text = '',
    [string] $RouteType = 'orthogonal'
  )

  $edge = [ordered]@{
    from = $From
    to = $To
    connector = 'Dynamic connector'
    stencil = 'BASFLO_M.VSSX'
    routeType = $RouteType
  }
  if ($Text) {
    $edge.text = $Text
  }
  return [pscustomobject]$edge
}

function Invoke-VisioGraphLayout {
  param(
    [Parameter(Mandatory = $true)] $GraphModel,
    [ValidateSet('Flowchart', 'Hierarchical', 'Tree', 'Network', 'Radial')] [string] $Strategy = 'Flowchart',
    [ValidateSet('LR', 'RL', 'TB', 'BT')] [string] $Direction = 'LR',
    [double] $StartX = 1.2,
    [double] $StartY = 3.5,
    [double] $HorizontalSpacing = 2.4,
    [double] $VerticalSpacing = 1.4
  )

  $layoutGraph = Copy-VisioGraphModel -GraphModel $GraphModel
  if (-not $layoutGraph.layout) {
    $layoutGraph | Add-Member -NotePropertyName layout -NotePropertyValue ([pscustomobject]@{}) -Force
  }
  $layoutGraph.layout | Add-Member -NotePropertyName strategy -NotePropertyValue $Strategy -Force
  $layoutGraph.layout | Add-Member -NotePropertyName direction -NotePropertyValue $Direction -Force
  $layoutGraph.layout | Add-Member -NotePropertyName engine -NotePropertyValue 'GraphModel' -Force

  $nodes = @($layoutGraph.nodes)
  if ($nodes.Count -eq 0) {
    return $layoutGraph
  }

  $indexById = @{}
  for ($i = 0; $i -lt $nodes.Count; $i++) {
    if ($nodes[$i].id) {
      $indexById[[string] $nodes[$i].id] = $i
    }
  }

  $depthById = @{}
  foreach ($node in $nodes) {
    if ($node.id) {
      $depthById[[string] $node.id] = 0
    }
  }

  $layoutEdges = New-Object System.Collections.ArrayList
  $branchPriorityById = @{}
  $branchAdvanceByEdgeKey = @{}
  foreach ($node in $nodes) {
    if ($node.id) {
      $branchPriorityById[[string] $node.id] = 0
    }
  }

  function Test-VisioLayoutPathExists {
    param(
      [Parameter(Mandatory = $true)] [string] $StartId,
      [Parameter(Mandatory = $true)] [string] $TargetId,
      [string] $SkipFrom = '',
      [string] $SkipTo = ''
    )

    if ($StartId -eq $TargetId) {
      return $true
    }

    $visited = @{}
    $queue = New-Object System.Collections.Queue
    $queue.Enqueue($StartId)
    $visited[$StartId] = $true

    while ($queue.Count -gt 0) {
      $current = [string] $queue.Dequeue()
      foreach ($candidateEdge in @($layoutGraph.edges)) {
        if (-not $candidateEdge.from -or -not $candidateEdge.to) { continue }
        $candidateFrom = [string] $candidateEdge.from
        $candidateTo = [string] $candidateEdge.to
        if ($candidateFrom -eq $SkipFrom -and $candidateTo -eq $SkipTo) { continue }
        if ($candidateFrom -ne $current) { continue }
        if ($candidateTo -eq $TargetId) {
          return $true
        }
        if (-not $visited.ContainsKey($candidateTo)) {
          $visited[$candidateTo] = $true
          $queue.Enqueue($candidateTo)
        }
      }
    }

    return $false
  }

  foreach ($edge in @($layoutGraph.edges)) {
    if (-not $edge.from -or -not $edge.to) {
      continue
    }
    $from = [string] $edge.from
    $to = [string] $edge.to
    if (-not $depthById.ContainsKey($from) -or -not $depthById.ContainsKey($to)) {
      continue
    }

    $edgeText = if ($edge.text) { [string] $edge.text } else { '' }
    $isNegativeBranch = $edgeText -match '(不通过|拒绝|驳回|退回|返回|修改|否|no|false|reject)'
    if ($isNegativeBranch) {
      $branchPriorityById[$to] = [Math]::Max([int] $branchPriorityById[$to], 2)
    } elseif ($edgeText -match '(通过|批准|同意|完成|是|yes|true|approve)') {
      $branchPriorityById[$to] = [Math]::Min([int] $branchPriorityById[$to], 0)
    }

    $fromIndex = if ($indexById.ContainsKey($from)) { [int] $indexById[$from] } else { -1 }
    $toIndex = if ($indexById.ContainsKey($to)) { [int] $indexById[$to] } else { -1 }
    $isCycleEdge = Test-VisioLayoutPathExists -StartId $to -TargetId $from -SkipFrom $from -SkipTo $to
    $looksLikeFeedbackEdge = (
      ($fromIndex -ge 0 -and $toIndex -ge 0 -and $toIndex -le $fromIndex) -or
      ($edgeText -match '(重新提交|重试|返工|回退|返回|修改|retry|redo|rework)')
    )
    if (-not ($isCycleEdge -and $looksLikeFeedbackEdge)) {
      $null = $layoutEdges.Add($edge)
      $hasTargetLoopback = Test-VisioLayoutPathExists -StartId $to -TargetId $from -SkipFrom $from -SkipTo $to
      $branchAdvanceByEdgeKey["$from->$to"] = if ($isNegativeBranch -and $hasTargetLoopback) { 0 } else { 1 }
    }
  }

  for ($pass = 0; $pass -lt $nodes.Count; $pass++) {
    $changed = $false
    foreach ($edge in @($layoutEdges.ToArray())) {
      $from = [string] $edge.from
      $to = [string] $edge.to
      $advance = if ($branchAdvanceByEdgeKey.ContainsKey("$from->$to")) { [int] $branchAdvanceByEdgeKey["$from->$to"] } else { 1 }
      $candidateDepth = [int] $depthById[$from] + $advance
      if ($candidateDepth -gt [int] $depthById[$to]) {
        $depthById[$to] = $candidateDepth
        $changed = $true
      }
    }
    if (-not $changed) {
      break
    }
  }

  $slotById = @{}
  $rankGroups = @{}
  foreach ($node in $nodes) {
    $id = [string] $node.id
    $rank = if ($depthById.ContainsKey($id)) { [int] $depthById[$id] } else { 0 }
    if (-not $rankGroups.ContainsKey($rank)) {
      $rankGroups[$rank] = New-Object System.Collections.ArrayList
    }
    $null = $rankGroups[$rank].Add($node)
  }

  foreach ($rankKey in @($rankGroups.Keys)) {
    $rankNodes = @($rankGroups[$rankKey].ToArray())
    $orderedRankNodes = @($rankNodes | Sort-Object `
      @{ Expression = { if ($branchPriorityById.ContainsKey([string] $_.id)) { [int] $branchPriorityById[[string] $_.id] } else { 0 } } }, `
      @{ Expression = { if ($indexById.ContainsKey([string] $_.id)) { [int] $indexById[[string] $_.id] } else { [int]::MaxValue } } })
    for ($slotIndex = 0; $slotIndex -lt $orderedRankNodes.Count; $slotIndex++) {
      $slotById[[string] $orderedRankNodes[$slotIndex].id] = $slotIndex
    }
  }

  foreach ($node in $nodes) {
    $id = [string] $node.id
    $rank = if ($depthById.ContainsKey($id)) { [int] $depthById[$id] } else { 0 }
    $slot = if ($slotById.ContainsKey($id)) { [int] $slotById[$id] } else { 0 }

    $hasX = $null -ne $node.x
    $hasY = $null -ne $node.y
    if ($hasX -and $hasY) {
      continue
    }

    switch ($Direction) {
      'RL' {
        $x = $StartX - ($rank * $HorizontalSpacing)
        $y = $StartY - ($slot * $VerticalSpacing)
      }
      'TB' {
        $x = $StartX + ($slot * $HorizontalSpacing)
        $y = $StartY - ($rank * $VerticalSpacing)
      }
      'BT' {
        $x = $StartX + ($slot * $HorizontalSpacing)
        $y = $StartY + ($rank * $VerticalSpacing)
      }
      default {
        $x = $StartX + ($rank * $HorizontalSpacing)
        $y = $StartY - ($slot * $VerticalSpacing)
      }
    }

    if (-not $hasX) {
      $node | Add-Member -NotePropertyName x -NotePropertyValue $x -Force
    }
    if (-not $hasY) {
      $node | Add-Member -NotePropertyName y -NotePropertyValue $y -Force
    }
    if ($null -eq $node.width) {
      $node | Add-Member -NotePropertyName width -NotePropertyValue 1.4 -Force
    }
    if ($null -eq $node.height) {
      $node | Add-Member -NotePropertyName height -NotePropertyValue 0.6 -Force
    }
  }

  return $layoutGraph
}

function Convert-TextFlowLabelToSemanticType {
  param(
    [Parameter(Mandatory = $true)] [string] $Text
  )

  $normalized = $Text.Trim()
  if ($normalized -match '^(开始|发起|启动|提交|Start|Begin)' -or $normalized -match '(发起申请|提交申请|提交报销单|提交报销|发起报销|提交请求|发起请求)$') { return 'start' }
  if ($normalized -match '(结束|完成|终止|End|Finish)$') { return 'end' }
  if ($normalized -match '[?？]$' -or $normalized -match '^(是否|判断).*[?？]?$') { return 'decision' }
  if ($normalized -match '(文档|单据|表单|报告|申请单|Document)') { return 'document' }
  if ($normalized -match '(数据库|数据表|数据存储|数据仓库|Database|DB)') { return 'database' }
  return 'process'
}

function Convert-TextFlowStatementToEdgeParts {
  param(
    [Parameter(Mandatory = $true)] [string] $Statement
  )

  $trimmed = $Statement.Trim()
  if (-not $trimmed) { return @() }

  $connectorPattern = '\s*(?:--\s*(?<label>.*?)\s*--?>|(?<plain>-->|->|=>|→|再到|然后|到|至))\s*'
  $matches = @([regex]::Matches($trimmed, $connectorPattern))
  if ($matches.Count -eq 0) {
    return @()
  }

  $edgeParts = New-Object System.Collections.ArrayList
  $currentNode = $trimmed.Substring(0, $matches[0].Index).Trim()

  for ($i = 0; $i -lt $matches.Count; $i++) {
    $match = $matches[$i]
    $targetStart = $match.Index + $match.Length
    $targetEnd = if ($i -lt ($matches.Count - 1)) { $matches[$i + 1].Index } else { $trimmed.Length }
    $target = $trimmed.Substring($targetStart, $targetEnd - $targetStart).Trim()
    if (-not $currentNode -or -not $target) {
      continue
    }

    $label = ''
    if ($match.Groups['label'].Success) {
      $label = $match.Groups['label'].Value.Trim()
    }

    $null = $edgeParts.Add([pscustomobject]@{
      from = $currentNode
      to = $target
      label = $label
    })
    $currentNode = $target
  }

  return @($edgeParts.ToArray())
}

function Convert-TextFlowToVisioGraphModel {
  param(
    [Parameter(Mandatory = $true)] [string] $Text,
    [ValidateSet('LR', 'RL', 'TB', 'BT')] [string] $LayoutDirection = 'LR',
    [ValidateSet('fidelity', 'balanced', 'clean')] [string] $RoutingIntent = 'balanced'
  )

  $cleanText = $Text.Trim()
  if (-not $cleanText) {
    throw 'Text flow is empty.'
  }

  $nodeByText = [ordered]@{}
  $edges = New-Object System.Collections.ArrayList
  $nextNodeNumber = 1

  function Ensure-TextFlowNode {
    param(
      [Parameter(Mandatory = $true)] [string] $Label
    )

    $key = $Label.Trim()
    if (-not $key) {
      throw 'Text flow contains an empty node label.'
    }
    if (-not $nodeByText.Contains($key)) {
      $id = ('n{0}' -f $script:TextFlowNextNodeNumber)
      $script:TextFlowNextNodeNumber += 1
      $semanticType = Convert-TextFlowLabelToSemanticType -Text $key
      $nodeByText[$key] = [pscustomobject]@{
        id = $id
        text = $key
        semanticType = $semanticType
        stencil = 'BASFLO_M.VSSX'
        preferredMaster = Get-VisioGraphModelSemanticMaster -SemanticType $semanticType
      }
    }
    return $nodeByText[$key]
  }

  $script:TextFlowNextNodeNumber = $nextNodeNumber
  try {
    foreach ($statement in @($cleanText -split '[;\r\n]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })) {
      $edgeParts = @(Convert-TextFlowStatementToEdgeParts -Statement $statement)
      if ($edgeParts.Count -eq 0) {
        $null = Ensure-TextFlowNode -Label $statement
        continue
      }

      foreach ($edgePart in $edgeParts) {
        $from = Ensure-TextFlowNode -Label ([string] $edgePart.from)
        $to = Ensure-TextFlowNode -Label ([string] $edgePart.to)
        $null = $edges.Add((New-VisioGraphModelEdge -From ([string] $from.id) -To ([string] $to.id) -Text ([string] $edgePart.label) -RouteType 'orthogonal'))
      }
    }
  } finally {
    Remove-Variable -Scope Script -Name TextFlowNextNodeNumber -ErrorAction SilentlyContinue
  }

  if ($nodeByText.Count -eq 0) {
    throw 'Text flow did not contain any nodes.'
  }
  if ($edges.Count -eq 0 -and $nodeByText.Count -gt 1) {
    throw 'Text flow did not contain any recognizable edges.'
  }

  return [pscustomobject]@{
    diagramType = 'flowchart'
    sourceFormat = 'text-flow'
    routingIntent = $RoutingIntent
    layout = [pscustomobject]@{
      strategy = 'Flowchart'
      direction = $LayoutDirection
    }
    nodes = @($nodeByText.Values)
    edges = @($edges.ToArray())
  }
}

function Convert-TextFlowToVisio {
  param(
    [Parameter(Mandatory = $true)] [string] $Text,
    [Parameter(Mandatory = $true)] [string] $OutputPath,
    [string] $PreviewPath = '',
    [ValidateSet('LR', 'RL', 'TB', 'BT')] [string] $LayoutDirection = 'LR',
    [ValidateSet('fidelity', 'balanced', 'clean')] [string] $RoutingIntent = 'balanced',
    [switch] $Visible,
    [switch] $Invisible,
    [switch] $Force
  )

  $graphModel = Convert-TextFlowToVisioGraphModel -Text $Text -LayoutDirection $LayoutDirection -RoutingIntent $RoutingIntent
  return Render-VisioGraphModel -GraphModel $graphModel -OutputPath $OutputPath -PreviewPath $PreviewPath -LayoutStrategy Flowchart -LayoutDirection $LayoutDirection -RoutingIntent $RoutingIntent -Visible:$Visible -Invisible:$Invisible -Force:$Force
}

function Convert-NaturalLanguageFlowToTextFlow {
  param(
    [Parameter(Mandatory = $true)] [string] $Text
  )

  $clean = $Text.Trim()
  if (-not $clean) {
    throw 'Natural language flow text is empty.'
  }

  $content = $clean
  if ($content -match '[:：](?<body>.+)$') {
    $content = $Matches.body
  }
  $content = $content -replace '[。.!！]+', ';'
  $content = $content -replace '\s+', ''

  function Convert-NaturalLanguageActionChainToTextFlow {
    param(
      [string] $ActionText
    )

    $action = $ActionText.Trim()
    if (-not $action) { return '' }
    $action = $action -replace '^(后|之后|然后|再|接着|就|则)', ''
    $action = $action -replace '(后|之后)$', ''
    $parts = @($action -split '(?:并且|并|然后|再|接着|后|之后)' | ForEach-Object {
      $_.Trim()
    } | Where-Object { $_ })
    if ($parts.Count -eq 0) {
      return $action
    }
    return ($parts -join ' -> ')
  }

  $startLabel = '发起申请'
  if ($content -match '(?<actor>员工|用户|申请人|客户|发起人)?(?<action>发起申请|提交申请|提交报销单|提交报销|发起报销|提交请求|发起请求)') {
    $actor = if ($Matches.actor) { $Matches.actor } else { '' }
    $startLabel = "$actor$($Matches.action)"
  }

  $validationLabel = '系统校验资料'
  $validationMatch = [regex]::Match($content, '(?<actor>系统|平台|服务|人工)?(?<action>校验发票|校验资料|检查发票|检查资料|校验报销单|校验单据)')
  if ($validationMatch.Success) {
    $actor = if ($validationMatch.Groups['actor'].Success) { $validationMatch.Groups['actor'].Value } else { '' }
    $validationLabel = "$actor$($validationMatch.Groups['action'].Value)"
  }

  $approvalLabel = '审批'
  if ($content -match '(?<role>部门经理|经理|主管|负责人|财务|管理员|HR|人事)?(?<action>审批|审核|复核)') {
    $role = if ($Matches.role) { $Matches.role } else { '' }
    $approvalLabel = "$role$($Matches.action)"
  }

  $hasCompletenessBranch = $content -match '(资料|材料|发票|单据).{0,6}(不完整|完整|齐全|缺失)'
  $hasAuditBranch = $content -match '(审核|审批|复核).{0,6}(通过|不通过|拒绝|驳回)'
  if ($hasCompletenessBranch -and $hasAuditBranch) {
    $incompleteLabel = '退回补充'
    $incompleteMatch = [regex]::Match($content, '(?:资料|材料|发票|单据)?不完整.{0,4}(?<target>退回补充|补充资料|补充材料|退回修改|返回修改|驳回)')
    if ($incompleteMatch.Success) {
      $incompleteLabel = $incompleteMatch.Groups['target'].Value
      if ($incompleteLabel -eq '补充材料') { $incompleteLabel = '补充资料' }
    }

    $completeLabel = $approvalLabel
    $completeMatch = [regex]::Match($content, '(?<!不)(?:资料|材料|发票|单据)?完整(?:就|则|后)?(?<target>(?:财务|主管|经理|人工|部门经理|负责人)?(?:审核|审批|复核))')
    if ($completeMatch.Success) {
      $completeLabel = $completeMatch.Groups['target'].Value
    }

    $auditSuccessChain = '流程结束'
    $auditSuccessMatch = [regex]::Match($content, '(?:审核|审批|复核)通过(?:后|就|则)?(?<target>[^,，;；。.!！]+)')
    if ($auditSuccessMatch.Success -and $auditSuccessMatch.Groups['target'].Value.Trim()) {
      $auditSuccessChain = Convert-NaturalLanguageActionChainToTextFlow -ActionText $auditSuccessMatch.Groups['target'].Value
    } elseif ($content -match '(付款|支付).{0,4}(归档|完成|结束)') {
      $auditSuccessChain = '付款 -> 归档'
    }

    $auditRejectLabel = '驳回'
    $auditRejectMatch = [regex]::Match($content, '(?:审核|审批|复核)不通过(?:就|则|后)?(?<target>驳回|拒绝|流程结束|结束|退回修改|返回修改)')
    if ($auditRejectMatch.Success) {
      $auditRejectLabel = $auditRejectMatch.Groups['target'].Value
      if ($auditRejectLabel -eq '结束') { $auditRejectLabel = '流程结束' }
    }

    return "$startLabel -> $validationLabel -> 资料是否完整? --不完整-> $incompleteLabel; 资料是否完整? --完整-> $completeLabel -> 审核是否通过? --通过-> $auditSuccessChain; 审核是否通过? --不通过-> $auditRejectLabel"
  }

  $successLabel = '流程结束'
  if ($content -match '(?:通过|批准|同意|审核通过).{0,8}(?<success>流程结束|结束|完成|归档|通知申请人)') {
    $successLabel = $Matches.success
    if ($successLabel -eq '结束') { $successLabel = '流程结束' }
  }

  $rejectLabel = ''
  if ($content -match '(?:不通过|未通过|拒绝|驳回|不同意).{0,10}(?<reject>返回修改|退回修改|重新提交|补充材料|结束|流程结束)') {
    $rejectLabel = $Matches.reject
    if ($rejectLabel -eq '退回修改') { $rejectLabel = '返回修改' }
    if ($rejectLabel -eq '结束') { $rejectLabel = '流程结束' }
  }

  if ($content -match '(如果|若|如).*?(通过|批准|同意)' -and $rejectLabel) {
    return "$startLabel -> $approvalLabel -> 是否通过? --通过-> $successLabel; 是否通过? --不通过-> $rejectLabel"
  }

  $steps = New-Object System.Collections.ArrayList
  foreach ($clause in @($content -split '[;；,，]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })) {
    $step = $clause
    $step = $step -replace '^(先|然后|再|接着|最后|如果|若|如)', ''
    $step = $step -replace '^(帮我画一个|画一个|生成一个|创建一个).*(流程|流程图)', ''
    $step = $step -replace '^就', ''
    if (-not $step) { continue }
    if ($step -match '(通过|不通过|拒绝|驳回|否则)') { continue }
    $null = $steps.Add($step)
  }

  if ($steps.Count -ge 2) {
    return (@($steps.ToArray()) -join ' -> ')
  }

  throw 'Natural language flow could not be normalized. Use arrow-style text such as: 发起申请 -> 审批 -> 是否通过? --通过-> 结束.'
}

function Convert-NaturalLanguageFlowToVisioGraphModel {
  param(
    [Parameter(Mandatory = $true)] [string] $Text,
    [ValidateSet('LR', 'RL', 'TB', 'BT')] [string] $LayoutDirection = 'LR',
    [ValidateSet('fidelity', 'balanced', 'clean')] [string] $RoutingIntent = 'balanced'
  )

  $textFlow = Convert-NaturalLanguageFlowToTextFlow -Text $Text
  $graphModel = Convert-TextFlowToVisioGraphModel -Text $textFlow -LayoutDirection $LayoutDirection -RoutingIntent $RoutingIntent
  $graphModel | Add-Member -NotePropertyName sourceFormat -NotePropertyValue 'natural-language' -Force
  $graphModel | Add-Member -NotePropertyName sourceText -NotePropertyValue $Text -Force
  $graphModel | Add-Member -NotePropertyName normalizedTextFlow -NotePropertyValue $textFlow -Force
  return $graphModel
}

function Convert-NaturalLanguageFlowToVisio {
  param(
    [Parameter(Mandatory = $true)] [string] $Text,
    [Parameter(Mandatory = $true)] [string] $OutputPath,
    [string] $PreviewPath = '',
    [ValidateSet('LR', 'RL', 'TB', 'BT')] [string] $LayoutDirection = 'LR',
    [ValidateSet('fidelity', 'balanced', 'clean')] [string] $RoutingIntent = 'balanced',
    [switch] $Visible,
    [switch] $Invisible,
    [switch] $Force
  )

  $graphModel = Convert-NaturalLanguageFlowToVisioGraphModel -Text $Text -LayoutDirection $LayoutDirection -RoutingIntent $RoutingIntent
  return Render-VisioGraphModel -GraphModel $graphModel -OutputPath $OutputPath -PreviewPath $PreviewPath -LayoutStrategy Flowchart -LayoutDirection $LayoutDirection -RoutingIntent $RoutingIntent -Visible:$Visible -Invisible:$Invisible -Force:$Force
}

function Convert-MermaidToVisioGraphModel {
  param(
    [Parameter(Mandatory = $true)] [string] $MermaidText
  )

  $lines = @($MermaidText -split "\r?\n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
  if ($lines.Count -eq 0) {
    throw 'Mermaid text is empty.'
  }

  $header = $lines[0]
  if ($header -notmatch '^graph\s+(?<dir>TD|LR|TB|BT)\s*$') {
    throw 'Only Mermaid graph TD/LR/TB/BT flowcharts are supported in this first version.'
  }

  $direction = $Matches.dir
  $nodes = [ordered]@{}
  $edges = New-Object System.Collections.ArrayList
  $diagramType = 'flowchart'

  function Ensure-MermaidNode {
    param(
      [Parameter(Mandatory = $true)] [string] $NodeId,
      [string] $Label = ''
    )

    if (-not $nodes.Contains($NodeId)) {
      $nodes[$NodeId] = [ordered]@{
        id = $NodeId
        text = $Label
        stencil = 'BASFLO_M.VSSX'
      }
    } elseif ($Label -and -not $nodes[$NodeId].text) {
      $nodes[$NodeId].text = $Label
    }
  }

  function Convert-MermaidNodeExpression {
    param([Parameter(Mandatory = $true)] [string] $Expression)

    $expr = $Expression.Trim()
    if ($expr -match '^(?<id>[A-Za-z0-9_]+)\(\((?<label>.*?)\)\)$') {
      $id = $Matches.id
      $label = $Matches.label
      $semanticType = if ($label -match '结束|end|stop') { 'end' } else { 'start' }
      return [pscustomobject]@{
        id = $id
        label = $label
        semanticType = $semanticType
      }
    }
    if ($expr -match '^(?<id>[A-Za-z0-9_]+)\[\[(?<label>.*?)\]\]$') {
      $id = $Matches.id
      $label = $Matches.label
      return [pscustomobject]@{
        id = $id
        label = $label
        semanticType = 'process'
      }
    }
    if ($expr -match '^(?<id>[A-Za-z0-9_]+)\[(?<label>.*?)\]$') {
      $id = $Matches.id
      $label = $Matches.label
      return [pscustomobject]@{
        id = $id
        label = $label
        semanticType = 'process'
      }
    }
    if ($expr -match '^(?<id>[A-Za-z0-9_]+)\{(?<label>.*?)\}$') {
      $id = $Matches.id
      $label = $Matches.label
      return [pscustomobject]@{
        id = $id
        label = $label
        semanticType = 'decision'
      }
    }
    if ($expr -match '^(?<id>[A-Za-z0-9_]+)$') {
      $id = $Matches.id
      return [pscustomobject]@{
        id = $id
        label = ''
        semanticType = 'process'
      }
    }

    throw "Unsupported Mermaid node expression: $Expression"
  }

  function Split-MermaidEdgeExpression {
    param([Parameter(Mandatory = $true)] [string] $Expression)

    $text = $Expression.Trim()
    if ($text -match '^(?<left>.+?)\s+--\s*(?<label>.+?)\s*--$') {
      $left = $Matches.left.Trim()
      $label = $Matches.label.Trim()
      return [pscustomobject]@{
        node = $left
        label = $label
      }
    }
    if ($text -match '^(?<left>.+?)\s+--\s*(?<label>.+?)$') {
      $left = $Matches.left.Trim()
      $label = $Matches.label.Trim()
      return [pscustomobject]@{
        node = $left
        label = $label
      }
    }
    return [pscustomobject]@{
      node = $text
      label = ''
    }
  }

  foreach ($line in $lines[1..($lines.Count - 1)]) {
    if ($line -match '^(?<left>.+?)\s*(?<arrow>-->|---|-.->)\s*(?<right>.+)$') {
      $leftSplit = Split-MermaidEdgeExpression -Expression $Matches.left
      $rightSplit = Split-MermaidEdgeExpression -Expression $Matches.right
      $leftExpr = Convert-MermaidNodeExpression -Expression $leftSplit.node
      $rightExpr = Convert-MermaidNodeExpression -Expression $rightSplit.node

      Ensure-MermaidNode -NodeId $leftExpr.id -Label $leftExpr.label
      Ensure-MermaidNode -NodeId $rightExpr.id -Label $rightExpr.label
      if (-not $nodes[$leftExpr.id].semanticType) {
        $nodes[$leftExpr.id].semanticType = $leftExpr.semanticType
      }
      if (-not $nodes[$rightExpr.id].semanticType) {
        $nodes[$rightExpr.id].semanticType = $rightExpr.semanticType
      }
      if (-not $nodes[$leftExpr.id].semanticType -and $nodes[$leftExpr.id].text -match '开始|start|begin') {
        $nodes[$leftExpr.id].semanticType = 'start'
      }
      if (-not $nodes[$rightExpr.id].semanticType -and $nodes[$rightExpr.id].text -match '结束|end|stop') {
        $nodes[$rightExpr.id].semanticType = 'end'
      }

      $edgeText = ''
      $routeType = 'orthogonal'
      if ($Matches.arrow -eq '-.->') {
        $routeType = 'straight'
      }
      if ($leftSplit.label) {
        $edgeText = $leftSplit.label
      } elseif ($rightSplit.label) {
        $edgeText = $rightSplit.label
      } elseif ($line -match '--\s*(?<label>[^-]+?)\s*-->') {
        $edgeText = $Matches.label.Trim()
      }

      $null = $edges.Add((New-VisioGraphModelEdge -From $leftExpr.id -To $rightExpr.id -Text $edgeText -RouteType $routeType))
      continue
    }

    if ($line -match '^(?<expr>[A-Za-z0-9_]+(?:\(\(.*?\)\)|\[\[.*?\]\]|\[.*?\]|\{.*?\})?)$') {
      $expr = Convert-MermaidNodeExpression -Expression $Matches.expr
      Ensure-MermaidNode -NodeId $expr.id -Label $expr.label
      if (-not $nodes[$expr.id].semanticType) {
        $nodes[$expr.id].semanticType = $expr.semanticType
      }
    }
  }

  $diagramNodes = foreach ($node in $nodes.Values) {
    $semanticType = if ($node.semanticType) { [string] $node.semanticType } else { 'process' }
    $preferredMaster = Get-VisioGraphModelSemanticMaster -SemanticType $semanticType
    New-VisioGraphModelNode -Id ([string] $node.id) -Text ([string] $node.text) -SemanticType $semanticType -Stencil ([string] $node.stencil) -PreferredMaster $preferredMaster
  }

  [pscustomobject]@{
    diagramType = $diagramType
    layout = [pscustomobject]@{
      strategy = 'Flowchart'
      direction = $direction
    }
    nodes = @($diagramNodes)
    edges = @($edges.ToArray())
    sourceFormat = 'mermaid'
  }
}

function Convert-MermaidToVisio {
  param(
    [Parameter(Mandatory = $true)] [string] $MermaidText,
    [Parameter(Mandatory = $true)] [string] $OutputPath,
    [string] $PreviewPath = '',
    [switch] $Visible,
    [switch] $Invisible,
    [switch] $Force
  )

  $graphModel = Convert-MermaidToVisioGraphModel -MermaidText $MermaidText
  $layoutDirection = [string] $graphModel.layout.direction
  if ($layoutDirection -eq 'TD') { $layoutDirection = 'TB' }
  return Render-VisioGraphModel -GraphModel $graphModel -OutputPath $OutputPath -PreviewPath $PreviewPath -LayoutStrategy Flowchart -LayoutDirection $layoutDirection -Visible:$Visible -Invisible:$Invisible -Force:$Force
}

function Get-DrawIOGraphModelSemanticType {
  param(
    [string] $Style = '',
    [string] $Text = ''
  )

  $styleValue = if ($Style) { [string] $Style } else { '' }
  $textValue = if ($Text) { [string] $Text } else { '' }

  if ($styleValue -match 'ellipse|terminator|start') {
    if ($textValue -match '开始|start|begin|起始') {
      return 'start'
    }
    if ($textValue -match '结束|end|stop|finish|终止') {
      return 'end'
    }
    return 'process'
  }

  if ($styleValue -match 'rhombus|diamond|decision') {
    return 'decision'
  }

  if ($styleValue -match 'rounded=1') {
    return 'process'
  }

  return 'process'
}

function Convert-DrawIOToVisioGraphModel {
  param(
    [string] $DrawIOXml = '',
    [string] $DrawIOPath = ''
  )

  if (-not $DrawIOXml) {
    if (-not $DrawIOPath) {
      throw 'Either DrawIOXml or DrawIOPath must be provided.'
    }
    $DrawIOXml = Get-Content -Raw -LiteralPath $DrawIOPath
  }

  try {
    [xml] $xml = $DrawIOXml
  } catch {
    throw 'Draw.io input is not valid XML. This first version only supports raw uncompressed draw.io XML.'
  }

  $diagram = $xml.mxfile.diagram | Select-Object -First 1
  if (-not $diagram) {
    throw 'Draw.io XML does not contain a diagram element.'
  }

  $graphModel = $diagram.mxGraphModel
  if (-not $graphModel) {
    throw 'Draw.io diagram does not contain an mxGraphModel element.'
  }

  $root = $graphModel.root
  if (-not $root) {
    throw 'Draw.io mxGraphModel does not contain a root element.'
  }

  $scale = 100.0
  $rawNodes = New-Object System.Collections.ArrayList
  $rawEdges = New-Object System.Collections.ArrayList
  $maxRight = 0.0
  $maxBottom = 0.0

  foreach ($cell in @($root.mxCell)) {
    if (-not $cell) {
      continue
    }

    $cellId = [string] $cell.id
    if (-not $cellId -or $cellId -in @('0', '1')) {
      continue
    }

    if ([string] $cell.vertex -eq '1') {
      $geometry = $cell.mxGeometry
      if (-not $geometry) {
        continue
      }

      $xIn = if ($null -ne $geometry.x) { [double] $geometry.x / $scale } else { 0.0 }
      $yIn = if ($null -ne $geometry.y) { [double] $geometry.y / $scale } else { 0.0 }
      $widthIn = if ($null -ne $geometry.width) { [double] $geometry.width / $scale } else { 1.4 }
      $heightIn = if ($null -ne $geometry.height) { [double] $geometry.height / $scale } else { 0.6 }
      $text = if ($null -ne $cell.value) { [string] $cell.value } else { '' }
      $style = if ($null -ne $cell.style) { [string] $cell.style } else { '' }
      $semanticType = Get-DrawIOGraphModelSemanticType -Style $style -Text $text

      $maxRight = [Math]::Max($maxRight, $xIn + $widthIn)
      $maxBottom = [Math]::Max($maxBottom, $yIn + $heightIn)

      $null = $rawNodes.Add([pscustomobject]@{
        id = $cellId
        sourceId = $cellId
        text = $text
        style = $style
        semanticType = $semanticType
        x = $xIn
        y = $yIn
        width = $widthIn
        height = $heightIn
      })
      continue
    }

    if ([string] $cell.edge -eq '1') {
      $style = if ($null -ne $cell.style) { [string] $cell.style } else { '' }
      $routeType = 'orthogonal'
      if ($style -match 'straightEdgeStyle') {
        $routeType = 'straight'
      } elseif ($style -match 'curved|elbowEdgeStyle') {
        $routeType = 'curved'
      }

      $null = $rawEdges.Add([pscustomobject]@{
        id = $cellId
        sourceId = $cellId
        from = [string] $cell.source
        to = [string] $cell.target
        text = if ($null -ne $cell.value) { [string] $cell.value } else { '' }
        style = $style
        routeType = $routeType
      })
    }
  }

  $pageWidth = if ($maxRight -gt 0.0) { $maxRight + 0.5 } else { 8.5 }
  $pageHeight = if ($maxBottom -gt 0.0) { $maxBottom + 0.5 } else { 6.0 }

  $nodes = foreach ($node in $rawNodes) {
    [pscustomobject]@{
      id = [string] $node.id
      sourceId = [string] $node.sourceId
      text = [string] $node.text
      semanticType = [string] $node.semanticType
      stencil = 'BASFLO_M.VSSX'
      preferredMaster = Get-VisioGraphModelSemanticMaster -SemanticType ([string] $node.semanticType)
      x = [double] $node.x + ([double] $node.width / 2.0)
      y = $pageHeight - ([double] $node.y + ([double] $node.height / 2.0))
      width = [double] $node.width
      height = [double] $node.height
    }
  }

  $edges = foreach ($edge in $rawEdges) {
    [pscustomobject]@{
      id = [string] $edge.id
      sourceId = [string] $edge.sourceId
      from = [string] $edge.from
      to = [string] $edge.to
      text = [string] $edge.text
      stencil = 'BASFLO_M.VSSX'
      connector = 'Dynamic connector'
      routeType = [string] $edge.routeType
    }
  }

  [pscustomobject]@{
    diagramType = 'flowchart'
    layout = [pscustomobject]@{
      strategy = 'Flowchart'
      direction = 'LR'
      engine = 'DrawIO'
    }
    page = [pscustomobject]@{
      width = $pageWidth
      height = $pageHeight
    }
    nodes = @($nodes)
    edges = @($edges)
    sourceFormat = 'drawio'
  }
}

function Convert-DrawIOToVisio {
  param(
    [string] $DrawIOXml = '',
    [string] $DrawIOPath = '',
    [Parameter(Mandatory = $true)] [string] $OutputPath,
    [string] $PreviewPath = '',
    [switch] $Visible,
    [switch] $Invisible,
    [switch] $Force
  )

  $graphModel = Convert-DrawIOToVisioGraphModel -DrawIOXml $DrawIOXml -DrawIOPath $DrawIOPath
  return Render-VisioGraphModel -GraphModel $graphModel -OutputPath $OutputPath -PreviewPath $PreviewPath -Visible:$Visible -Invisible:$Invisible -Force:$Force
}

function Convert-ImageReconstructionToVisioGraphModel {
  param(
    [Parameter(Mandatory = $true)] $ReconstructionModel
  )

  if (-not $ReconstructionModel.source) {
    throw 'Image reconstruction model is missing source metadata.'
  }
  if (-not $ReconstructionModel.source.width -or -not $ReconstructionModel.source.height) {
    throw 'Image reconstruction model source must include width and height in pixels.'
  }

  $sourceWidth = [double] $ReconstructionModel.source.width
  $sourceHeight = [double] $ReconstructionModel.source.height
  if ($sourceWidth -le 0 -or $sourceHeight -le 0) {
    throw 'Image reconstruction source width and height must be positive.'
  }

  $pageWidth = if ($ReconstructionModel.page -and $ReconstructionModel.page.width) { [double] $ReconstructionModel.page.width } else { 8.5 }
  $pageHeight = if ($ReconstructionModel.page -and $ReconstructionModel.page.height) { [double] $ReconstructionModel.page.height } else { $pageWidth * ($sourceHeight / $sourceWidth) }
  $scaleX = $pageWidth / $sourceWidth
  $scaleY = $pageHeight / $sourceHeight

  $nodes = foreach ($node in @($ReconstructionModel.nodes)) {
    if (-not $node.id) {
      throw 'Image reconstruction node is missing id.'
    }
    if (-not $node.bounds) {
      throw "Image reconstruction node is missing bounds: $($node.id)"
    }

    $bounds = $node.bounds
    $semanticType = if ($node.semanticType) { [string] $node.semanticType } else { 'process' }
    $width = [double] $bounds.width * $scaleX
    $height = [double] $bounds.height * $scaleY
    $centerX = ([double] $bounds.x + ([double] $bounds.width / 2.0)) * $scaleX
    $centerY = $pageHeight - (([double] $bounds.y + ([double] $bounds.height / 2.0)) * $scaleY)

    [pscustomobject]@{
      id = [string] $node.id
      sourceId = if ($node.sourceId) { [string] $node.sourceId } else { [string] $node.id }
      text = if ($null -ne $node.text) { [string] $node.text } else { '' }
      semanticType = $semanticType
      stencil = if ($node.stencil) { [string] $node.stencil } else { 'BASFLO_M.VSSX' }
      preferredMaster = if ($node.preferredMaster) { [string] $node.preferredMaster } else { Get-VisioGraphModelSemanticMaster -SemanticType $semanticType }
      x = $centerX
      y = $centerY
      width = $width
      height = $height
      style = $node.style
    }
  }

  $edges = foreach ($connector in @($ReconstructionModel.connectors)) {
    if (-not $connector.from -or -not $connector.to) {
      throw 'Image reconstruction connector is missing from/to.'
    }

    $fromPoint = if ($connector.fromPoint) { $connector.fromPoint } else { [pscustomobject]@{ x = 1.0; y = 0.5 } }
    $toPoint = if ($connector.toPoint) { $connector.toPoint } else { [pscustomobject]@{ x = 0.0; y = 0.5 } }
    $routeType = if ($connector.routeType) { ([string] $connector.routeType).ToLowerInvariant() } else { 'orthogonal' }
    if ($routeType -notin @('orthogonal', 'straight', 'curved')) {
      $routeType = 'orthogonal'
    }

    [pscustomobject]@{
      id = if ($connector.id) { [string] $connector.id } else { '' }
      sourceId = if ($connector.sourceId) { [string] $connector.sourceId } elseif ($connector.id) { [string] $connector.id } else { '' }
      from = [string] $connector.from
      to = [string] $connector.to
      text = if ($null -ne $connector.text) { [string] $connector.text } else { '' }
      stencil = if ($connector.stencil) { [string] $connector.stencil } else { 'BASFLO_M.VSSX' }
      connector = if ($connector.connector) { [string] $connector.connector } else { 'Dynamic connector' }
      routeType = $routeType
      fromX = if ($null -ne $fromPoint.x) { [double] $fromPoint.x } else { 1.0 }
      fromY = if ($null -ne $fromPoint.y) { [double] $fromPoint.y } else { 0.5 }
      toX = if ($null -ne $toPoint.x) { [double] $toPoint.x } else { 0.0 }
      toY = if ($null -ne $toPoint.y) { [double] $toPoint.y } else { 0.5 }
      style = $connector.style
    }
  }

  [pscustomobject]@{
    diagramType = if ($ReconstructionModel.diagramType) { [string] $ReconstructionModel.diagramType } else { 'image-reconstruction' }
    layout = [pscustomobject]@{
      direction = 'absolute'
      engine = 'StructuredImageModel'
    }
    page = [pscustomobject]@{
      width = $pageWidth
      height = $pageHeight
    }
    source = $ReconstructionModel.source
    nodes = @($nodes)
    edges = @($edges)
    sourceFormat = 'image-reconstruction'
  }
}

function Convert-ImageReconstructionToVisio {
  param(
    [Parameter(Mandatory = $true)] $ReconstructionModel,
    [Parameter(Mandatory = $true)] [string] $OutputPath,
    [string] $PreviewPath = '',
    [switch] $Visible,
    [switch] $Invisible,
    [switch] $Force
  )

  $graphModel = Convert-ImageReconstructionToVisioGraphModel -ReconstructionModel $ReconstructionModel
  $rendered = Render-VisioGraphModel -GraphModel $graphModel -OutputPath $OutputPath -PreviewPath $PreviewPath -Visible:$Visible -Invisible:$Invisible -Force:$Force
  $rendered | Add-Member -NotePropertyName source -NotePropertyValue $graphModel.source -Force
  return $rendered
}

function Render-VisioGraphModel {
  param(
    [Parameter(Mandatory = $true)] $GraphModel,
    [Parameter(Mandatory = $true)] [string] $OutputPath,
    [string] $PreviewPath = '',
    [ValidateSet('', 'Flowchart', 'Hierarchical', 'Tree', 'Network', 'Radial')] [string] $LayoutStrategy = '',
    [ValidateSet('LR', 'RL', 'TB', 'BT')] [string] $LayoutDirection = 'LR',
    [ValidateSet('', 'fidelity', 'balanced', 'clean')] [string] $RoutingIntent = '',
    [switch] $Visible,
    [switch] $Invisible,
    [switch] $Force
  )

  if ($LayoutStrategy) {
    $GraphModel = Invoke-VisioGraphLayout -GraphModel $GraphModel -Strategy $LayoutStrategy -Direction $LayoutDirection
  } elseif ($GraphModel.layout -and $GraphModel.layout.strategy) {
    $direction = if ($GraphModel.layout.direction) { [string] $GraphModel.layout.direction } else { $LayoutDirection }
    $GraphModel = Invoke-VisioGraphLayout -GraphModel $GraphModel -Strategy ([string] $GraphModel.layout.strategy) -Direction $direction
  }

  $OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
  if ($PreviewPath) {
    $PreviewPath = [System.IO.Path]::GetFullPath($PreviewPath)
  }

  if ((Test-Path -LiteralPath $OutputPath) -and -not $Force) {
    throw "OutputPath already exists: $OutputPath. Use -Force to overwrite."
  }

  $outputDirectory = Split-Path -Path $OutputPath -Parent
  if ($outputDirectory) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
  }
  if ($PreviewPath) {
    $previewDirectory = Split-Path -Path $PreviewPath -Parent
    if ($previewDirectory) {
      New-Item -ItemType Directory -Path $previewDirectory -Force | Out-Null
    }
  }

  if ((Test-Path -LiteralPath $OutputPath) -and $Force) {
    Remove-Item -LiteralPath $OutputPath -Force
  }
  if ($PreviewPath -and (Test-Path -LiteralPath $PreviewPath) -and $Force) {
    Remove-Item -LiteralPath $PreviewPath -Force
  }

  if ($Visible -and $Invisible) {
    throw 'Specify only one of -Visible or -Invisible.'
  }

  $runInvisible = [bool] $Invisible
  if ($Invisible) {
    $visio = New-InvisibleVisioApplication
  } else {
    $visio = New-VisibleVisioApplication
  }

  $openedStencils = @{}
  $shapeById = @{}
  $renderedNodes = New-Object System.Collections.ArrayList
  $renderedEdges = New-Object System.Collections.ArrayList
  $existingRoutes = New-Object System.Collections.ArrayList
  $doc = $null

  $effectiveRoutingIntent = Resolve-VisioRoutingIntent -GraphModel $GraphModel -OverrideIntent $RoutingIntent

  $nodeIndex = @{}
  foreach ($node in @($GraphModel.nodes)) {
    if ($node.id) { $nodeIndex[[string] $node.id] = $node }
  }

  function Get-GraphStencil {
    param(
      [string] $StencilName
    )

    if (-not $StencilName) {
      $StencilName = 'BASFLO_M.VSSX'
    }
    if (-not $openedStencils.ContainsKey($StencilName)) {
      $openedStencils[$StencilName] = Open-VisioStencilReadOnly -Visio $visio -StencilNameOrPath $StencilName
    }
    return $openedStencils[$StencilName]
  }

  function Get-GraphMaster {
    param(
      [string] $StencilName,
      [string] $MasterName
    )

    $stencil = Get-GraphStencil -StencilName $StencilName
    return $stencil.Masters.ItemU($MasterName)
  }

  try {
    $doc = $visio.Documents.Add('')
    $page = $visio.ActivePage

    if ($GraphModel.page) {
      if ($GraphModel.page.width) {
        $null = $page.PageSheet.CellsU('PageWidth').FormulaU = "$($GraphModel.page.width) in"
      }
      if ($GraphModel.page.height) {
        $null = $page.PageSheet.CellsU('PageHeight').FormulaU = "$($GraphModel.page.height) in"
      }
    }

    foreach ($node in @($GraphModel.nodes)) {
      if (-not $node.id) {
        throw 'Graph node is missing id.'
      }

      $stencilName = if ($node.stencil) { [string] $node.stencil } else { 'BASFLO_M.VSSX' }
      $masterName = if ($node.master) {
        [string] $node.master
      } elseif ($node.preferredMaster) {
        [string] $node.preferredMaster
      } elseif ($node.semanticType) {
        Get-VisioGraphModelSemanticMaster -SemanticType ([string] $node.semanticType)
      } else {
        'Process'
      }

      $x = if ($null -ne $node.x) { [double] $node.x } else { 1.0 }
      $y = if ($null -ne $node.y) { [double] $node.y } else { 1.0 }
      $width = if ($null -ne $node.width) { [double] $node.width } else { 1.4 }
      $height = if ($null -ne $node.height) { [double] $node.height } else { 0.6 }

      $master = Get-GraphMaster -StencilName $stencilName -MasterName $masterName
      $shape = $page.Drop($master, $x, $y)
      Set-VisioShapeBounds -Shape $shape -X $x -Y $y -Width $width -Height $height
      if ($null -ne $node.text) {
        $shape.Text = [string] $node.text
      }
      if ($node.style) {
        if ($node.style.fill) {
          $line = if ($node.style.line) { [string] $node.style.line } else { [string] $node.style.fill }
          $lineWeight = if ($node.style.lineWeight) { [string] $node.style.lineWeight } else { '1 pt' }
          Set-VisioShapeFill -Shape $shape -Fill ([string] $node.style.fill) -Line $line -LineWeight $lineWeight
        }
        if ($node.style.textColor -or $node.style.textSize -or $node.style.bold) {
          $textColor = if ($node.style.textColor) { [string] $node.style.textColor } else { 'RGB(45,45,45)' }
          $textSize = if ($node.style.textSize) { [double] $node.style.textSize } else { 11 }
          $bold = if ($node.style.bold) { [int] $node.style.bold } else { 0 }
          Set-VisioTextStyle -Shape $shape -Color $textColor -Size $textSize -Bold $bold
        }
      }
      $shapeById[[string] $node.id] = $shape
      $null = $renderedNodes.Add([pscustomobject]@{
        id = [string] $node.id
        semanticType = [string] $node.semanticType
        stencil = $stencilName
        master = $masterName
        nameU = $shape.NameU
        text = [string] $node.text
        x = $x
        y = $y
        width = $width
        height = $height
      })
    }

    foreach ($edge in @($GraphModel.edges)) {
      if (-not $edge.from -or -not $edge.to) {
        throw 'Graph edge is missing from/to.'
      }
      if (-not $shapeById.ContainsKey([string] $edge.from)) {
        throw "Graph edge references missing source node: $($edge.from)"
      }
      if (-not $shapeById.ContainsKey([string] $edge.to)) {
        throw "Graph edge references missing target node: $($edge.to)"
      }

      $routeDecision = Resolve-VisioRouteDecision -Edge $edge -GraphModel $GraphModel -RoutingIntent $effectiveRoutingIntent -NodeIndex $nodeIndex -ExistingRoutes @($existingRoutes.ToArray())
      $routeType = $routeDecision.routeType

      $stencilName = if ($edge.stencil) { [string] $edge.stencil } else { 'BASFLO_M.VSSX' }
      $connectorName = if ($edge.connector) { [string] $edge.connector } else { 'Dynamic connector' }
      $connectorMaster = Get-GraphMaster -StencilName $stencilName -MasterName $connectorName

      $fromX = $routeDecision.fromX
      $fromY = $routeDecision.fromY
      $toX = $routeDecision.toX
      $toY = $routeDecision.toY

      if ($null -eq $fromX -or $null -eq $fromY -or $null -eq $toX -or $null -eq $toY) {
        $rawFromX = if ($null -ne $edge.fromX) { [double] $edge.fromX } else { $null }
        $rawFromY = if ($null -ne $edge.fromY) { [double] $edge.fromY } else { $null }
        $rawToX = if ($null -ne $edge.toX) { [double] $edge.toX } else { $null }
        $rawToY = if ($null -ne $edge.toY) { [double] $edge.toY } else { $null }
        $gluePoints = Resolve-VisioConnectorGluePoints -From $shapeById[[string] $edge.from] -To $shapeById[[string] $edge.to] -FromX $rawFromX -FromY $rawFromY -ToX $rawToX -ToY $rawToY
        $fromX = [double] $gluePoints.FromX
        $fromY = [double] $gluePoints.FromY
        $toX = [double] $gluePoints.ToX
        $toY = [double] $gluePoints.ToY
      }

      switch ($routeType) {
        'straight' {
          $connector = Connect-VisioShapesStraight -Page $page -ConnectorMaster $connectorMaster -From $shapeById[[string] $edge.from] -To $shapeById[[string] $edge.to] -FromX $fromX -FromY $fromY -ToX $toX -ToY $toY
        }
        'curved' {
          $connector = Connect-VisioShapesCurved -Page $page -ConnectorMaster $connectorMaster -From $shapeById[[string] $edge.from] -To $shapeById[[string] $edge.to] -FromX $fromX -FromY $fromY -ToX $toX -ToY $toY
        }
        default {
          $connector = Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $shapeById[[string] $edge.from] -To $shapeById[[string] $edge.to] -FromX $fromX -FromY $fromY -ToX $toX -ToY $toY
          $routeType = 'orthogonal'
        }
      }

      if ($null -ne $edge.text) {
        $connector.Text = [string] $edge.text
      }
      if ($edge.style) {
        if ($edge.style.lineColor) {
          $null = $connector.CellsU('LineColor').FormulaU = [string] $edge.style.lineColor
        }
        if ($edge.style.lineWeight) {
          $null = $connector.CellsU('LineWeight').FormulaU = [string] $edge.style.lineWeight
        }
        if ($null -ne $edge.style.endArrow) {
          $null = $connector.CellsU('EndArrow').FormulaU = "$([int] $edge.style.endArrow)"
        }
      }

      $null = $renderedEdges.Add([pscustomobject]@{
        from = [string] $edge.from
        to = [string] $edge.to
        text = [string] $edge.text
        stencil = $stencilName
        connector = $connectorName
        routeType = $routeType
        fromX = $fromX
        fromY = $fromY
        toX = $toX
        toY = $toY
        fromSide = $routeDecision.fromSide
        toSide = $routeDecision.toSide
        routeScore = $routeDecision.routeScore
        routeLength = $routeDecision.routeLength
        obstacleHits = $routeDecision.obstacleHits
        avoidancePenalty = $routeDecision.avoidancePenalty
        labelHits = $routeDecision.labelHits
        labelPenalty = $routeDecision.labelPenalty
        boundaryHits = $routeDecision.boundaryHits
        boundaryPenalty = $routeDecision.boundaryPenalty
        crossingHits = $routeDecision.crossingHits
        crossingPenalty = $routeDecision.crossingPenalty
        oneD = [int] $connector.OneD
        isLoopback = $routeDecision.isLoopback
        routingIntent = $effectiveRoutingIntent
      })

      if ($routeDecision.routePoints -and @($routeDecision.routePoints).Count -gt 1) {
        $null = $existingRoutes.Add([pscustomobject]@{
          from = [string] $edge.from
          to = [string] $edge.to
          points = @($routeDecision.routePoints)
        })
      }
    }

    $null = $doc.SaveAs($OutputPath)
    if ($PreviewPath) {
      $null = $page.Export($PreviewPath)
    }

    return [pscustomobject]@{
      diagramType = [string] $GraphModel.diagramType
      layout = $GraphModel.layout
      routingIntent = $effectiveRoutingIntent
      outputPath = $OutputPath
      previewPath = $PreviewPath
      nodes = @($renderedNodes.ToArray())
      edges = @($renderedEdges.ToArray())
    }
  } finally {
    foreach ($stencil in $openedStencils.Values) {
      try {
        $null = $stencil.Close()
      } catch {
      }
    }
    if ($null -ne $doc -and $runInvisible) {
      try {
        $null = $doc.Close()
      } catch {
      }
    }
    if ($runInvisible) {
      $null = $visio.Quit()
    }
  }
}

function Add-VisioLabel {
  param(
    [Parameter(Mandatory = $true)] $Page,
    [Parameter(Mandatory = $true)] [string] $Text,
    [double] $X,
    [double] $Y,
    [double] $Width = 0.9,
    [double] $Height = 0.25
  )

  $label = $Page.DrawRectangle($X - ($Width / 2), $Y - ($Height / 2), $X + ($Width / 2), $Y + ($Height / 2))
  $label.Text = $Text
  $null = $label.CellsU('LinePattern').FormulaU = '0'
  $null = $label.CellsU('FillPattern').FormulaU = '0'
  Set-VisioTextStyle -Shape $label -Size 9.5
  return $label
}

function Add-VisioRotHelperType {
  if ('RotVisioDocumentHelper' -as [type]) {
    return
  }

  $code = @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public static class RotVisioDocumentHelper {
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

    public static object GetObjectByDisplayName(string targetName) {
        IRunningObjectTable rot;
        IEnumMoniker enumMoniker;
        IBindCtx bindCtx;
        GetRunningObjectTable(0, out rot);
        CreateBindCtx(0, out bindCtx);
        rot.EnumRunning(out enumMoniker);
        IMoniker[] monikers = new IMoniker[1];
        IntPtr fetched = IntPtr.Zero;
        while (enumMoniker.Next(1, monikers, fetched) == 0) {
            string name;
            monikers[0].GetDisplayName(bindCtx, null, out name);
            if (String.Equals(name, targetName, StringComparison.OrdinalIgnoreCase)) {
                object obj;
                rot.GetObject(monikers[0], out obj);
                return obj;
            }
        }
        return null;
    }
}
'@
  Add-Type -TypeDefinition $code
}

function Get-OpenVisioDocumentByPath {
  param(
    [Parameter(Mandatory = $true)] [string] $Path
  )

  Add-VisioRotHelperType
  $resolved = [System.IO.Path]::GetFullPath($Path)
  return [RotVisioDocumentHelper]::GetObjectByDisplayName($resolved)
}

# ---------------------------------------------------------------------------
# Phase 6: Semantic Master Search
# ---------------------------------------------------------------------------

function Get-SemanticMasterMap {
  param(
    [string] $MapPath = ''
  )

  if (-not $MapPath) {
    $MapPath = Join-Path $PSScriptRoot '..\references\semantic-master-map.json'
  }
  if (-not (Test-Path -LiteralPath $MapPath)) {
    throw "Semantic master map not found: $MapPath"
  }
  return Get-Content -Raw -LiteralPath $MapPath | ConvertFrom-Json
}

function Find-SemanticVisioMaster {
  param(
    [Parameter(Mandatory = $true)] [string] $Query,
    [string] $MapPath = '',
    [string] $Category = '',
    [int] $MaxResults = 5,
    [switch] $IncludeFallback
  )

  $map = Get-SemanticMasterMap -MapPath $MapPath
  $queryLower = $Query.ToLowerInvariant().Trim()

  # Score each mapping entry against the query
  $scored = New-Object System.Collections.ArrayList

  foreach ($entry in @($map.mappings)) {
    if ($Category -and $entry.category -ne $Category) {
      continue
    }

    $bestScore = 0
    $bestTerm = ''
    foreach ($term in @($entry.terms)) {
      $termLower = $term.ToLowerInvariant()
      $score = 0

      # Exact match on a term
      if ($queryLower -eq $termLower) {
        $score = 100
      }
      # Query contains the term as a whole word
      elseif ($queryLower -match [regex]::Escape($termLower)) {
        $score = 80
      }
      # Term contains the query
      elseif ($termLower -match [regex]::Escape($queryLower)) {
        $score = 60
      }
      # Fuzzy: query words overlap with term words
      else {
        $queryWords = $queryLower -split '[\s,;/]+'
        $termWords = $termLower -split '[\s,;/]+'
        $overlap = 0
        foreach ($qw in $queryWords) {
          if ($qw.Length -lt 2) { continue }
          foreach ($tw in $termWords) {
            if ($tw -match [regex]::Escape($qw)) {
              $overlap++
              break
            }
          }
        }
        if ($overlap -gt 0) {
          $score = 30 + ($overlap * 10)
        }
      }

      # Apply priority penalty: lower priority number = higher preference
      if ($score -gt 0) {
        $priorityPenalty = if ($entry.priority) { [int] $entry.priority } else { 1 }
        $score = $score - ($priorityPenalty * 2)
      }

      if ($score -gt $bestScore) {
        $bestScore = $score
        $bestTerm = $term
      }
    }

    if ($bestScore -gt 0) {
      $null = $scored.Add([pscustomobject]@{
        Score = $bestScore
        MatchedTerm = $bestTerm
        Stencil = $entry.stencil
        Master = $entry.master
        SemanticType = $entry.semanticType
        Category = $entry.category
        Priority = $entry.priority
      })
    }
  }

  # Sort by score descending, then priority ascending
  $results = @($scored | Sort-Object -Property @{Expression={$_.Score}; Descending=$true}, @{Expression={$_.Priority}; Ascending=$true} |
    Select-Object -First $MaxResults)

  if ($results.Count -gt 0) {
    return $results
  }

  # Fallback: use Find-VisioMasters with the raw query
  if ($IncludeFallback) {
    $fallback = Find-VisioMasters -Query $Query -MaxResults $MaxResults
    if ($fallback.Count -gt 0) {
      return @($fallback | ForEach-Object {
        [pscustomobject]@{
          Score = 10
          MatchedTerm = $_.NameU
          Stencil = $_.Stencil
          Master = $_.NameU
          SemanticType = ''
          Category = 'discovered'
          Priority = 1
          Source = 'FallbackScan'
          FullName = $_.FullName
        }
      })
    }
  }

  return @()
}

# ---------------------------------------------------------------------------
# Phase 6: Local Session Cache
# ---------------------------------------------------------------------------

$script:VisioSessionCacheSchemaVersion = 1

function Get-VisioSessionCachePath {
  param(
    [string] $WorkspacePath = ''
  )

  if (-not $WorkspacePath) {
    $WorkspacePath = (Get-Location).Path
  }
  $cacheDir = Join-Path $WorkspacePath '.cache\visio-automation'
  if (-not (Test-Path -LiteralPath $cacheDir)) {
    New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null
  }
  return Join-Path $cacheDir 'session-cache.json'
}

function Get-VisioSessionCache {
  param(
    [string] $WorkspacePath = ''
  )

  $cachePath = Get-VisioSessionCachePath -WorkspacePath $WorkspacePath
  if (-not (Test-Path -LiteralPath $cachePath)) {
    return $null
  }

  try {
    $cache = Get-Content -Raw -LiteralPath $cachePath | ConvertFrom-Json
    if ([int] $cache.schemaVersion -ne $script:VisioSessionCacheSchemaVersion) {
      return $null
    }
    return $cache
  } catch {
    return $null
  }
}

function Save-VisioSessionCache {
  param(
    [Parameter(Mandatory = $true)] $CacheData,
    [string] $WorkspacePath = ''
  )

  $cachePath = Get-VisioSessionCachePath -WorkspacePath $WorkspacePath
  $CacheData | Add-Member -NotePropertyName schemaVersion -NotePropertyValue $script:VisioSessionCacheSchemaVersion -Force
  $CacheData | Add-Member -NotePropertyName updatedAt -NotePropertyValue (Get-Date).ToUniversalTime().ToString('o') -Force
  $CacheData | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $cachePath -Encoding utf8
  return $CacheData
}

function Update-VisioSessionCache {
  param(
    [string] $WorkspacePath = '',
    [string] $DiagramType = '',
    [string] $Stencil = '',
    [string] $MasterMapping = '',
    [string] $LastOutputPath = '',
    [string] $LastPreviewPath = '',
    [string] $LastLayoutStrategy = '',
    [string] $LastLayoutDirection = ''
  )

  $cache = Get-VisioSessionCache -WorkspacePath $WorkspacePath
  if (-not $cache) {
    $cache = [pscustomobject]@{
      diagramType = ''
      stencil = ''
      masterMapping = ''
      lastOutputPath = ''
      lastPreviewPath = ''
      lastLayoutStrategy = ''
      lastLayoutDirection = ''
      recentMasters = @()
    }
  }

  if ($DiagramType) { $cache | Add-Member -NotePropertyName diagramType -NotePropertyValue $DiagramType -Force }
  if ($Stencil) { $cache | Add-Member -NotePropertyName stencil -NotePropertyValue $Stencil -Force }
  if ($MasterMapping) { $cache | Add-Member -NotePropertyName masterMapping -NotePropertyValue $MasterMapping -Force }
  if ($LastOutputPath) { $cache | Add-Member -NotePropertyName lastOutputPath -NotePropertyValue $LastOutputPath -Force }
  if ($LastPreviewPath) { $cache | Add-Member -NotePropertyName lastPreviewPath -NotePropertyValue $LastPreviewPath -Force }
  if ($LastLayoutStrategy) { $cache | Add-Member -NotePropertyName lastLayoutStrategy -NotePropertyValue $LastLayoutStrategy -Force }
  if ($LastLayoutDirection) { $cache | Add-Member -NotePropertyName lastLayoutDirection -NotePropertyValue $LastLayoutDirection -Force }

  return Save-VisioSessionCache -CacheData $cache -WorkspacePath $WorkspacePath
}

function Add-VisioSessionRecentMaster {
  param(
    [string] $WorkspacePath = '',
    [Parameter(Mandatory = $true)] [string] $Stencil,
    [Parameter(Mandatory = $true)] [string] $MasterNameU,
    [int] $MaxRecent = 10
  )

  $cache = Get-VisioSessionCache -WorkspacePath $WorkspacePath
  if (-not $cache) {
    $cache = [pscustomobject]@{
      diagramType = ''
      stencil = ''
      masterMapping = ''
      lastOutputPath = ''
      lastPreviewPath = ''
      lastLayoutStrategy = ''
      lastLayoutDirection = ''
      recentMasters = @()
    }
  }

  $recentList = [System.Collections.ArrayList]@($cache.recentMasters)
  # Remove existing entry for the same master
  $recentList = [System.Collections.ArrayList]@($recentList | Where-Object { $_.master -ne $MasterNameU })
  # Add to front
  $null = $recentList.Insert(0, [pscustomobject]@{
    stencil = $Stencil
    master = $MasterNameU
    usedAt = (Get-Date).ToUniversalTime().ToString('o')
  })
  # Trim to max
  if ($recentList.Count -gt $MaxRecent) {
    $recentList = [System.Collections.ArrayList]@($recentList[0..($MaxRecent - 1)])
  }

  $cache | Add-Member -NotePropertyName recentMasters -NotePropertyValue @($recentList.ToArray()) -Force
  return Save-VisioSessionCache -CacheData $cache -WorkspacePath $WorkspacePath
}

# ---------------------------------------------------------------------------
# Phase 8: Smart Routing Policy
# ---------------------------------------------------------------------------

function Get-VisioRoutingDefault {
  param(
    [Parameter(Mandatory = $true)] [string] $DiagramType
  )

  switch -Regex ($DiagramType) {
    '^(bpmn|workflow)$' {
      return [pscustomobject]@{
        defaultRouteType = 'orthogonal'
        loopbackRouteType = 'curved'
        loopbackFromSide = 'top'
        loopbackToSide = 'top'
        crossEdgesPreferCurved = $true
      }
    }
    '^(network|uml|uml-component|architecture)$' {
      return [pscustomobject]@{
        defaultRouteType = 'orthogonal'
        loopbackRouteType = 'orthogonal'
        loopbackFromSide = 'top'
        loopbackToSide = 'bottom'
        crossEdgesPreferCurved = $true
      }
    }
    '^(orgchart|org-chart|hierarchy)$' {
      return [pscustomobject]@{
        defaultRouteType = 'orthogonal'
        loopbackRouteType = 'orthogonal'
        loopbackFromSide = 'auto'
        loopbackToSide = 'auto'
        crossEdgesPreferCurved = $false
      }
    }
    '^(dfd|data-flow)$' {
      return [pscustomobject]@{
        defaultRouteType = 'orthogonal'
        loopbackRouteType = 'curved'
        loopbackFromSide = 'auto'
        loopbackToSide = 'auto'
        crossEdgesPreferCurved = $true
      }
    }
    '^(image-reconstruction)$' {
      return [pscustomobject]@{
        defaultRouteType = 'preserve'
        loopbackRouteType = 'preserve'
        loopbackFromSide = 'preserve'
        loopbackToSide = 'preserve'
        crossEdgesPreferCurved = $true
      }
    }
    default {
      return [pscustomobject]@{
        defaultRouteType = 'orthogonal'
        loopbackRouteType = 'curved'
        loopbackFromSide = 'auto'
        loopbackToSide = 'auto'
        crossEdgesPreferCurved = $false
      }
    }
  }
}

function Resolve-VisioRoutingIntent {
  param(
    [Parameter(Mandatory = $true)] $GraphModel,
    [ValidateSet('fidelity', 'balanced', 'clean', '')] [string] $OverrideIntent = ''
  )

  if ($OverrideIntent) {
    return [string] $OverrideIntent
  }

  if ($GraphModel.routingIntent) {
    $declared = ([string] $GraphModel.routingIntent).ToLowerInvariant().Trim()
    if ($declared -in @('fidelity', 'balanced', 'clean')) {
      return $declared
    }
  }

  $sourceFormat = if ($GraphModel.sourceFormat) { [string] $GraphModel.sourceFormat } else { '' }
  $diagramType = if ($GraphModel.diagramType) { [string] $GraphModel.diagramType } else { '' }

  if ($sourceFormat -eq 'image-reconstruction') {
    return 'fidelity'
  }

  if ($sourceFormat -in @('drawio', 'mermaid')) {
    return 'balanced'
  }

  if ($diagramType -eq 'image-reconstruction') {
    return 'fidelity'
  }

  if ($diagramType -in @('network', 'uml', 'uml-component', 'architecture')) {
    return 'clean'
  }

  return 'balanced'
}

function Resolve-VisioLoopbackGlue {
  param(
    [Parameter(Mandatory = $true)] $FromNode,
    [Parameter(Mandatory = $true)] $ToNode,
    [Parameter(Mandatory = $true)] [string] $LayoutDirection,
    [string] $LoopbackFromSide = 'auto',
    [string] $LoopbackToSide = 'auto'
  )

  $fromX = if ($null -ne $FromNode.x) { [double] $FromNode.x } else { 0.0 }
  $fromY = if ($null -ne $FromNode.y) { [double] $FromNode.y } else { 0.0 }
  $toX = if ($null -ne $ToNode.x) { [double] $ToNode.x } else { 0.0 }
  $toY = if ($null -ne $ToNode.y) { [double] $ToNode.y } else { 0.0 }

  function Resolve-SideToGlue {
    param([string] $Side, [double] $RefX, [double] $RefY)

    switch ($Side) {
      'top'    { return [pscustomobject]@{ x = 0.5; y = 1.0 } }
      'bottom' { return [pscustomobject]@{ x = 0.5; y = 0.0 } }
      'left'   { return [pscustomobject]@{ x = 0.0; y = 0.5 } }
      'right'  { return [pscustomobject]@{ x = 1.0; y = 0.5 } }
      default  {
        switch ($LayoutDirection) {
          'TB' { return [pscustomobject]@{ x = 0.0; y = 0.5 } }
          'BT' { return [pscustomobject]@{ x = 1.0; y = 0.5 } }
          'RL' { return [pscustomobject]@{ x = 0.5; y = 1.0 } }
          default { return [pscustomobject]@{ x = 0.5; y = 1.0 } }
        }
      }
    }
  }

  $fromGlue = Resolve-SideToGlue -Side $LoopbackFromSide -RefX $fromX -RefY $fromY
  $toGlue = Resolve-SideToGlue -Side $LoopbackToSide -RefX $toX -RefY $toY

  return [pscustomobject]@{
    fromX = [double] $fromGlue.x
    fromY = [double] $fromGlue.y
    toX = [double] $toGlue.x
    toY = [double] $toGlue.y
  }
}

function Get-VisioGraphNodeBounds {
  param(
    [Parameter(Mandatory = $true)] $Node
  )

  $x = if ($null -ne $Node.x) { [double] $Node.x } else { 0.0 }
  $y = if ($null -ne $Node.y) { [double] $Node.y } else { 0.0 }
  $width = if ($null -ne $Node.width) { [double] $Node.width } else { 1.2 }
  $height = if ($null -ne $Node.height) { [double] $Node.height } else { 0.6 }

  return [pscustomobject]@{
    x = $x
    y = $y
    width = $width
    height = $height
    left = $x - ($width / 2.0)
    right = $x + ($width / 2.0)
    bottom = $y - ($height / 2.0)
    top = $y + ($height / 2.0)
  }
}

function Get-VisioRoutePort {
  param(
    [Parameter(Mandatory = $true)] $Bounds,
    [Parameter(Mandatory = $true)] [ValidateSet('top', 'bottom', 'left', 'right')] [string] $Side
  )

  switch ($Side) {
    'top' {
      return [pscustomobject]@{
        side = 'top'
        glueX = 0.5
        glueY = 1.0
        x = [double] $Bounds.x
        y = [double] $Bounds.top
      }
    }
    'bottom' {
      return [pscustomobject]@{
        side = 'bottom'
        glueX = 0.5
        glueY = 0.0
        x = [double] $Bounds.x
        y = [double] $Bounds.bottom
      }
    }
    'left' {
      return [pscustomobject]@{
        side = 'left'
        glueX = 0.0
        glueY = 0.5
        x = [double] $Bounds.left
        y = [double] $Bounds.y
      }
    }
    default {
      return [pscustomobject]@{
        side = 'right'
        glueX = 1.0
        glueY = 0.5
        x = [double] $Bounds.right
        y = [double] $Bounds.y
      }
    }
  }
}

function New-VisioRoutePoint {
  param(
    [Parameter(Mandatory = $true)] [double] $X,
    [Parameter(Mandatory = $true)] [double] $Y
  )

  return [pscustomobject]@{
    x = [double] $X
    y = [double] $Y
  }
}

function Get-VisioRouteSideVector {
  param(
    [Parameter(Mandatory = $true)] [ValidateSet('top', 'bottom', 'left', 'right')] [string] $Side
  )

  switch ($Side) {
    'top' { return [pscustomobject]@{ x = 0.0; y = 1.0 } }
    'bottom' { return [pscustomobject]@{ x = 0.0; y = -1.0 } }
    'left' { return [pscustomobject]@{ x = -1.0; y = 0.0 } }
    default { return [pscustomobject]@{ x = 1.0; y = 0.0 } }
  }
}

function Expand-VisioBounds {
  param(
    [Parameter(Mandatory = $true)] $Bounds,
    [double] $Padding = 0.08
  )

  return [pscustomobject]@{
    x = [double] $Bounds.x
    y = [double] $Bounds.y
    width = [double] $Bounds.width + (2.0 * $Padding)
    height = [double] $Bounds.height + (2.0 * $Padding)
    left = [double] $Bounds.left - $Padding
    right = [double] $Bounds.right + $Padding
    bottom = [double] $Bounds.bottom - $Padding
    top = [double] $Bounds.top + $Padding
  }
}

function Get-VisioGraphTextBounds {
  param(
    [Parameter(Mandatory = $true)] $TextItem
  )

  $boundsSource = if ($TextItem.bounds) { $TextItem.bounds } else { $TextItem }
  $x = if ($null -ne $boundsSource.x) { [double] $boundsSource.x } else { 0.0 }
  $y = if ($null -ne $boundsSource.y) { [double] $boundsSource.y } else { 0.0 }
  $width = if ($null -ne $boundsSource.width) { [double] $boundsSource.width } else { 1.0 }
  $height = if ($null -ne $boundsSource.height) { [double] $boundsSource.height } else { 0.3 }

  return [pscustomobject]@{
    x = $x
    y = $y
    width = $width
    height = $height
    left = $x - ($width / 2.0)
    right = $x + ($width / 2.0)
    bottom = $y - ($height / 2.0)
    top = $y + ($height / 2.0)
  }
}

function Get-VisioGraphRegionBounds {
  param(
    [Parameter(Mandatory = $true)] $Region
  )

  $boundsSource = if ($Region.bounds) { $Region.bounds } else { $Region }
  $x = if ($null -ne $boundsSource.x) { [double] $boundsSource.x } else { 0.0 }
  $y = if ($null -ne $boundsSource.y) { [double] $boundsSource.y } else { 0.0 }
  $width = if ($null -ne $boundsSource.width) { [double] $boundsSource.width } else { 1.0 }
  $height = if ($null -ne $boundsSource.height) { [double] $boundsSource.height } else { 1.0 }

  return [pscustomobject]@{
    x = $x
    y = $y
    width = $width
    height = $height
    left = $x - ($width / 2.0)
    right = $x + ($width / 2.0)
    bottom = $y - ($height / 2.0)
    top = $y + ($height / 2.0)
  }
}

function Test-VisioRangeOverlap {
  param(
    [Parameter(Mandatory = $true)] [double] $A1,
    [Parameter(Mandatory = $true)] [double] $A2,
    [Parameter(Mandatory = $true)] [double] $B1,
    [Parameter(Mandatory = $true)] [double] $B2
  )

  $minA = [Math]::Min($A1, $A2)
  $maxA = [Math]::Max($A1, $A2)
  $minB = [Math]::Min($B1, $B2)
  $maxB = [Math]::Max($B1, $B2)

  return ($maxA -gt $minB -and $maxB -gt $minA)
}

function Test-VisioPointInsideBounds {
  param(
    [Parameter(Mandatory = $true)] $Point,
    [Parameter(Mandatory = $true)] $Bounds
  )

  return (
    [double] $Point.x -gt [double] $Bounds.left -and
    [double] $Point.x -lt [double] $Bounds.right -and
    [double] $Point.y -gt [double] $Bounds.bottom -and
    [double] $Point.y -lt [double] $Bounds.top
  )
}

function Test-VisioRouteSegmentIntersectsBounds {
  param(
    [Parameter(Mandatory = $true)] $Start,
    [Parameter(Mandatory = $true)] $End,
    [Parameter(Mandatory = $true)] $Bounds
  )

  $tolerance = 0.0001
  $isVertical = [Math]::Abs([double] $Start.x - [double] $End.x) -lt $tolerance
  $isHorizontal = [Math]::Abs([double] $Start.y - [double] $End.y) -lt $tolerance

  if ($isVertical -and $isHorizontal) {
    return $false
  }

  if ($isVertical) {
    $x = [double] $Start.x
    if ($x -le [double] $Bounds.left -or $x -ge [double] $Bounds.right) {
      return $false
    }
    return Test-VisioRangeOverlap -A1 ([double] $Start.y) -A2 ([double] $End.y) -B1 ([double] $Bounds.bottom) -B2 ([double] $Bounds.top)
  }

  if ($isHorizontal) {
    $y = [double] $Start.y
    if ($y -le [double] $Bounds.bottom -or $y -ge [double] $Bounds.top) {
      return $false
    }
    return Test-VisioRangeOverlap -A1 ([double] $Start.x) -A2 ([double] $End.x) -B1 ([double] $Bounds.left) -B2 ([double] $Bounds.right)
  }

  # Current route candidates are orthogonal; keep a conservative bounding-box
  # fallback for future non-axis-aligned candidates.
  $segmentLeft = [Math]::Min([double] $Start.x, [double] $End.x)
  $segmentRight = [Math]::Max([double] $Start.x, [double] $End.x)
  $segmentBottom = [Math]::Min([double] $Start.y, [double] $End.y)
  $segmentTop = [Math]::Max([double] $Start.y, [double] $End.y)

  return (
    (Test-VisioRangeOverlap -A1 $segmentLeft -A2 $segmentRight -B1 ([double] $Bounds.left) -B2 ([double] $Bounds.right)) -and
    (Test-VisioRangeOverlap -A1 $segmentBottom -A2 $segmentTop -B1 ([double] $Bounds.bottom) -B2 ([double] $Bounds.top))
  )
}

function Test-VisioRouteSegmentsIntersect {
  param(
    [Parameter(Mandatory = $true)] $AStart,
    [Parameter(Mandatory = $true)] $AEnd,
    [Parameter(Mandatory = $true)] $BStart,
    [Parameter(Mandatory = $true)] $BEnd
  )

  $tolerance = 0.0001
  $aVertical = [Math]::Abs([double] $AStart.x - [double] $AEnd.x) -lt $tolerance
  $aHorizontal = [Math]::Abs([double] $AStart.y - [double] $AEnd.y) -lt $tolerance
  $bVertical = [Math]::Abs([double] $BStart.x - [double] $BEnd.x) -lt $tolerance
  $bHorizontal = [Math]::Abs([double] $BStart.y - [double] $BEnd.y) -lt $tolerance

  if (($aVertical -and $aHorizontal) -or ($bVertical -and $bHorizontal)) {
    return $false
  }

  if ($aVertical -and $bHorizontal) {
    $x = [double] $AStart.x
    $y = [double] $BStart.y
    return (
      (Test-VisioRangeOverlap -A1 $x -A2 $x -B1 ([double] $BStart.x) -B2 ([double] $BEnd.x)) -and
      (Test-VisioRangeOverlap -A1 $y -A2 $y -B1 ([double] $AStart.y) -B2 ([double] $AEnd.y))
    )
  }

  if ($aHorizontal -and $bVertical) {
    $x = [double] $BStart.x
    $y = [double] $AStart.y
    return (
      (Test-VisioRangeOverlap -A1 $x -A2 $x -B1 ([double] $AStart.x) -B2 ([double] $AEnd.x)) -and
      (Test-VisioRangeOverlap -A1 $y -A2 $y -B1 ([double] $BStart.y) -B2 ([double] $BEnd.y))
    )
  }

  if ($aVertical -and $bVertical -and [Math]::Abs([double] $AStart.x - [double] $BStart.x) -lt $tolerance) {
    return Test-VisioRangeOverlap -A1 ([double] $AStart.y) -A2 ([double] $AEnd.y) -B1 ([double] $BStart.y) -B2 ([double] $BEnd.y)
  }

  if ($aHorizontal -and $bHorizontal -and [Math]::Abs([double] $AStart.y - [double] $BStart.y) -lt $tolerance) {
    return Test-VisioRangeOverlap -A1 ([double] $AStart.x) -A2 ([double] $AEnd.x) -B1 ([double] $BStart.x) -B2 ([double] $BEnd.x)
  }

  return $false
}

function Get-VisioRoutePathMetrics {
  param(
    [Parameter(Mandatory = $true)] [array] $Points,
    [array] $ObstacleBounds = @(),
    [array] $LabelBounds = @(),
    [array] $BoundaryBounds = @(),
    [array] $ExistingRoutes = @()
  )

  $simplified = New-Object System.Collections.ArrayList
  foreach ($point in @($Points)) {
    if ($null -eq $point) { continue }
    $count = $simplified.Count
    if ($count -gt 0) {
      $last = $simplified[$count - 1]
      if ([Math]::Abs([double] $last.x - [double] $point.x) -lt 0.0001 -and
          [Math]::Abs([double] $last.y - [double] $point.y) -lt 0.0001) {
        continue
      }
    }
    $null = $simplified.Add((New-VisioRoutePoint -X ([double] $point.x) -Y ([double] $point.y)))
  }

  $pathPoints = @($simplified.ToArray())
  $length = 0.0
  $hitKeys = @{}
  $labelHitKeys = @{}
  $boundaryHitKeys = @{}
  $crossingKeys = @{}

  for ($i = 1; $i -lt $pathPoints.Count; $i++) {
    $start = $pathPoints[$i - 1]
    $end = $pathPoints[$i]
    $length += [Math]::Abs([double] $start.x - [double] $end.x) + [Math]::Abs([double] $start.y - [double] $end.y)

    for ($j = 0; $j -lt @($ObstacleBounds).Count; $j++) {
      $bounds = @($ObstacleBounds)[$j]
      if (Test-VisioRouteSegmentIntersectsBounds -Start $start -End $end -Bounds $bounds) {
        $hitKeys[[string] $j] = $true
      }
    }

    for ($j = 0; $j -lt @($LabelBounds).Count; $j++) {
      $bounds = @($LabelBounds)[$j]
      if (Test-VisioRouteSegmentIntersectsBounds -Start $start -End $end -Bounds $bounds) {
        $labelHitKeys[[string] $j] = $true
      }
    }

    for ($j = 0; $j -lt @($BoundaryBounds).Count; $j++) {
      $bounds = @($BoundaryBounds)[$j]
      if (Test-VisioRouteSegmentIntersectsBounds -Start $start -End $end -Bounds $bounds) {
        $boundaryHitKeys[[string] $j] = $true
      }
    }

    for ($routeIndex = 0; $routeIndex -lt @($ExistingRoutes).Count; $routeIndex++) {
      $existingRoute = @($ExistingRoutes)[$routeIndex]
      if ($null -eq $existingRoute -or $null -eq $existingRoute.points) { continue }
      $existingPoints = @($existingRoute.points)
      for ($existingIndex = 1; $existingIndex -lt $existingPoints.Count; $existingIndex++) {
        if (Test-VisioRouteSegmentsIntersect -AStart $start -AEnd $end -BStart $existingPoints[$existingIndex - 1] -BEnd $existingPoints[$existingIndex]) {
          $crossingKeys["$routeIndex-$existingIndex"] = $true
        }
      }
    }
  }

  return [pscustomobject]@{
    points = $pathPoints
    length = [double] $length
    obstacleHits = [int] $hitKeys.Count
    labelHits = [int] $labelHitKeys.Count
    boundaryHits = [int] $boundaryHitKeys.Count
    crossingHits = [int] $crossingKeys.Count
  }
}

function Resolve-VisioOrthogonalRoutePath {
  param(
    [Parameter(Mandatory = $true)] $FromPort,
    [Parameter(Mandatory = $true)] [ValidateSet('top', 'bottom', 'left', 'right')] [string] $FromSide,
    [Parameter(Mandatory = $true)] $ToPort,
    [Parameter(Mandatory = $true)] [ValidateSet('top', 'bottom', 'left', 'right')] [string] $ToSide,
    [array] $ObstacleBounds = @(),
    [array] $LabelBounds = @(),
    [array] $BoundaryBounds = @(),
    [array] $ExistingRoutes = @(),
    [double] $Clearance = 0.28
  )

  $fromVector = Get-VisioRouteSideVector -Side $FromSide
  $toVector = Get-VisioRouteSideVector -Side $ToSide
  $fromPoint = New-VisioRoutePoint -X ([double] $FromPort.x) -Y ([double] $FromPort.y)
  $toPoint = New-VisioRoutePoint -X ([double] $ToPort.x) -Y ([double] $ToPort.y)
  $fromOut = New-VisioRoutePoint -X ([double] $FromPort.x + ([double] $fromVector.x * $Clearance)) -Y ([double] $FromPort.y + ([double] $fromVector.y * $Clearance))
  $toOut = New-VisioRoutePoint -X ([double] $ToPort.x + ([double] $toVector.x * $Clearance)) -Y ([double] $ToPort.y + ([double] $toVector.y * $Clearance))

  $candidates = @(
    ,@(
      $fromPoint
      $fromOut
      (New-VisioRoutePoint -X ([double] $fromOut.x) -Y ([double] $toOut.y))
      $toOut
      $toPoint
    )
    ,@(
      $fromPoint
      $fromOut
      (New-VisioRoutePoint -X ([double] $toOut.x) -Y ([double] $fromOut.y))
      $toOut
      $toPoint
    )
  )

  $minX = [Math]::Min([Math]::Min([double] $fromOut.x, [double] $toOut.x), [Math]::Min([double] $fromPoint.x, [double] $toPoint.x))
  $maxX = [Math]::Max([Math]::Max([double] $fromOut.x, [double] $toOut.x), [Math]::Max([double] $fromPoint.x, [double] $toPoint.x))
  $minY = [Math]::Min([Math]::Min([double] $fromOut.y, [double] $toOut.y), [Math]::Min([double] $fromPoint.y, [double] $toPoint.y))
  $maxY = [Math]::Max([Math]::Max([double] $fromOut.y, [double] $toOut.y), [Math]::Max([double] $fromPoint.y, [double] $toPoint.y))

  foreach ($bounds in @($ObstacleBounds)) {
    if ($null -eq $bounds) { continue }
    $minX = [Math]::Min($minX, [double] $bounds.left)
    $maxX = [Math]::Max($maxX, [double] $bounds.right)
    $minY = [Math]::Min($minY, [double] $bounds.bottom)
    $maxY = [Math]::Max($maxY, [double] $bounds.top)
  }

  foreach ($bounds in @($LabelBounds)) {
    if ($null -eq $bounds) { continue }
    $minX = [Math]::Min($minX, [double] $bounds.left)
    $maxX = [Math]::Max($maxX, [double] $bounds.right)
    $minY = [Math]::Min($minY, [double] $bounds.bottom)
    $maxY = [Math]::Max($maxY, [double] $bounds.top)
  }

  foreach ($bounds in @($BoundaryBounds)) {
    if ($null -eq $bounds) { continue }
    $minX = [Math]::Min($minX, [double] $bounds.left)
    $maxX = [Math]::Max($maxX, [double] $bounds.right)
    $minY = [Math]::Min($minY, [double] $bounds.bottom)
    $maxY = [Math]::Max($maxY, [double] $bounds.top)
  }

  foreach ($existingRoute in @($ExistingRoutes)) {
    if ($null -eq $existingRoute -or $null -eq $existingRoute.points) { continue }
    foreach ($point in @($existingRoute.points)) {
      if ($null -eq $point) { continue }
      $minX = [Math]::Min($minX, [double] $point.x)
      $maxX = [Math]::Max($maxX, [double] $point.x)
      $minY = [Math]::Min($minY, [double] $point.y)
      $maxY = [Math]::Max($maxY, [double] $point.y)
    }
  }

  $outerPadding = [Math]::Max(($Clearance * 2.0), 0.5)
  $leftDetourX = $minX - $outerPadding
  $rightDetourX = $maxX + $outerPadding
  $bottomDetourY = $minY - $outerPadding
  $topDetourY = $maxY + $outerPadding

  $candidates += ,@(
    $fromPoint
    $fromOut
    (New-VisioRoutePoint -X $leftDetourX -Y ([double] $fromOut.y))
    (New-VisioRoutePoint -X $leftDetourX -Y ([double] $toOut.y))
    $toOut
    $toPoint
  )
  $candidates += ,@(
    $fromPoint
    $fromOut
    (New-VisioRoutePoint -X $rightDetourX -Y ([double] $fromOut.y))
    (New-VisioRoutePoint -X $rightDetourX -Y ([double] $toOut.y))
    $toOut
    $toPoint
  )
  $candidates += ,@(
    $fromPoint
    $fromOut
    (New-VisioRoutePoint -X ([double] $fromOut.x) -Y $bottomDetourY)
    (New-VisioRoutePoint -X ([double] $toOut.x) -Y $bottomDetourY)
    $toOut
    $toPoint
  )
  $candidates += ,@(
    $fromPoint
    $fromOut
    (New-VisioRoutePoint -X ([double] $fromOut.x) -Y $topDetourY)
    (New-VisioRoutePoint -X ([double] $toOut.x) -Y $topDetourY)
    $toOut
    $toPoint
  )

  $best = $null
  foreach ($candidate in $candidates) {
    $metrics = Get-VisioRoutePathMetrics -Points $candidate -ObstacleBounds $ObstacleBounds -LabelBounds $LabelBounds -BoundaryBounds $BoundaryBounds -ExistingRoutes $ExistingRoutes
    if ($null -eq $best -or
        $metrics.obstacleHits -lt $best.obstacleHits -or
        ($metrics.obstacleHits -eq $best.obstacleHits -and $metrics.labelHits -lt $best.labelHits) -or
        ($metrics.obstacleHits -eq $best.obstacleHits -and $metrics.labelHits -eq $best.labelHits -and $metrics.boundaryHits -lt $best.boundaryHits) -or
        ($metrics.obstacleHits -eq $best.obstacleHits -and $metrics.labelHits -eq $best.labelHits -and $metrics.boundaryHits -eq $best.boundaryHits -and $metrics.crossingHits -lt $best.crossingHits) -or
        ($metrics.obstacleHits -eq $best.obstacleHits -and $metrics.labelHits -eq $best.labelHits -and $metrics.boundaryHits -eq $best.boundaryHits -and $metrics.crossingHits -eq $best.crossingHits -and $metrics.length -lt $best.length)) {
      $best = $metrics
    }
  }

  return $best
}

function Resolve-VisioPreferredRouteSides {
  param(
    [Parameter(Mandatory = $true)] $FromNode,
    [Parameter(Mandatory = $true)] $ToNode,
    [Parameter(Mandatory = $true)] [string] $LayoutDirection,
    [bool] $IsLoopback = $false
  )

  $fromBounds = Get-VisioGraphNodeBounds -Node $FromNode
  $toBounds = Get-VisioGraphNodeBounds -Node $ToNode
  $dx = [double] $toBounds.x - [double] $fromBounds.x
  $dy = [double] $toBounds.y - [double] $fromBounds.y

  if ($IsLoopback -and $LayoutDirection -in @('LR', 'RL')) {
    if ([Math]::Abs($dx) -gt 0.15 -and [Math]::Abs($dy) -gt 0.15) {
      $fromSide = if ($dx -ge 0) { 'right' } else { 'left' }
      $toSide = if ($dy -ge 0) { 'bottom' } else { 'top' }
      return [pscustomobject]@{ fromSide = $fromSide; toSide = $toSide }
    }
    if ([Math]::Abs($dy) -gt 0.15) {
      if ($dy -ge 0) {
        return [pscustomobject]@{ fromSide = 'top'; toSide = 'bottom' }
      }
      return [pscustomobject]@{ fromSide = 'bottom'; toSide = 'top' }
    }
    if ($LayoutDirection -eq 'RL') {
      return [pscustomobject]@{ fromSide = 'right'; toSide = 'left' }
    }
    return [pscustomobject]@{ fromSide = 'left'; toSide = 'right' }
  }

  if ($IsLoopback -and $LayoutDirection -in @('TB', 'BT')) {
    if ([Math]::Abs($dx) -gt 0.15) {
      if ($dx -ge 0) {
        return [pscustomobject]@{ fromSide = 'right'; toSide = 'left' }
      }
      return [pscustomobject]@{ fromSide = 'left'; toSide = 'right' }
    }
    if ($LayoutDirection -eq 'BT') {
      return [pscustomobject]@{ fromSide = 'bottom'; toSide = 'top' }
    }
    return [pscustomobject]@{ fromSide = 'top'; toSide = 'bottom' }
  }

  if ([Math]::Abs($dx) -ge [Math]::Abs($dy)) {
    if ($dx -ge 0) {
      return [pscustomobject]@{ fromSide = 'right'; toSide = 'left' }
    }
    return [pscustomobject]@{ fromSide = 'left'; toSide = 'right' }
  }

  if ($dy -ge 0) {
    return [pscustomobject]@{ fromSide = 'top'; toSide = 'bottom' }
  }
  return [pscustomobject]@{ fromSide = 'bottom'; toSide = 'top' }
}

function Resolve-VisioOptimalConnectorGlue {
  param(
    [Parameter(Mandatory = $true)] $FromNode,
    [Parameter(Mandatory = $true)] $ToNode,
    [Parameter(Mandatory = $true)] [string] $LayoutDirection,
    [bool] $IsLoopback = $false,
    [string] $RoutingIntent = 'balanced',
    [array] $ObstacleNodes = @(),
    [array] $LabelBounds = @(),
    [array] $BoundaryBounds = @(),
    [array] $ExistingRoutes = @()
  )

  $fromBounds = Get-VisioGraphNodeBounds -Node $FromNode
  $toBounds = Get-VisioGraphNodeBounds -Node $ToNode
  $dx = [double] $toBounds.x - [double] $fromBounds.x
  $dy = [double] $toBounds.y - [double] $fromBounds.y
  $preferred = Resolve-VisioPreferredRouteSides -FromNode $FromNode -ToNode $ToNode -LayoutDirection $LayoutDirection -IsLoopback:$IsLoopback
  $sides = @('top', 'bottom', 'left', 'right')
  $best = $null
  $obstacleBounds = @()
  foreach ($obstacleNode in @($ObstacleNodes)) {
    if ($null -eq $obstacleNode) { continue }
    $obstacleBounds += @(Expand-VisioBounds -Bounds (Get-VisioGraphNodeBounds -Node $obstacleNode) -Padding 0.08)
  }

  foreach ($fromSide in $sides) {
    $fromPort = Get-VisioRoutePort -Bounds $fromBounds -Side $fromSide
    foreach ($toSide in $sides) {
      $toPort = Get-VisioRoutePort -Bounds $toBounds -Side $toSide
      $routePath = Resolve-VisioOrthogonalRoutePath -FromPort $fromPort -FromSide $fromSide -ToPort $toPort -ToSide $toSide -ObstacleBounds $obstacleBounds -LabelBounds $LabelBounds -BoundaryBounds $BoundaryBounds -ExistingRoutes $ExistingRoutes
      $routeLength = if ($null -ne $routePath) { [double] $routePath.length } else { [Math]::Abs([double] $fromPort.x - [double] $toPort.x) + [Math]::Abs([double] $fromPort.y - [double] $toPort.y) }
      $obstacleHits = if ($null -ne $routePath) { [int] $routePath.obstacleHits } else { 0 }
      $labelHits = if ($null -ne $routePath) { [int] $routePath.labelHits } else { 0 }
      $boundaryHits = if ($null -ne $routePath) { [int] $routePath.boundaryHits } else { 0 }
      $crossingHits = if ($null -ne $routePath) { [int] $routePath.crossingHits } else { 0 }
      $score = $routeLength

      if ($fromSide -ne $preferred.fromSide) { $score += 4.0 }
      if ($toSide -ne $preferred.toSide) { $score += 4.0 }

      $avoidancePenalty = 0.0
      if ($obstacleHits -gt 0) {
        $obstaclePenalty = switch ($RoutingIntent) {
          'fidelity' { 18.0 }
          'clean' { 42.0 }
          default { 30.0 }
        }
        $avoidancePenalty = [double] ($obstacleHits * $obstaclePenalty)
        $score += $avoidancePenalty
      }

      $labelPenalty = 0.0
      if ($labelHits -gt 0) {
        $labelPenaltyUnit = switch ($RoutingIntent) {
          'fidelity' { 10.0 }
          'clean' { 34.0 }
          default { 20.0 }
        }
        $labelPenalty = [double] ($labelHits * $labelPenaltyUnit)
        $score += $labelPenalty
      }

      $boundaryPenalty = 0.0
      if ($boundaryHits -gt 0) {
        $boundaryPenaltyUnit = switch ($RoutingIntent) {
          'fidelity' { 8.0 }
          'clean' { 34.0 }
          default { 18.0 }
        }
        $boundaryPenalty = [double] ($boundaryHits * $boundaryPenaltyUnit)
        $score += $boundaryPenalty
      }

      $crossingPenalty = 0.0
      if ($crossingHits -gt 0) {
        $crossingPenaltyUnit = switch ($RoutingIntent) {
          'fidelity' { 8.0 }
          'clean' { 24.0 }
          default { 14.0 }
        }
        $crossingPenalty = [double] ($crossingHits * $crossingPenaltyUnit)
        $score += $crossingPenalty
      }

      $oppositePenalty = 0.0
      if (($fromSide -eq 'left' -and $dx -gt 0) -or
          ($fromSide -eq 'right' -and $dx -lt 0) -or
          ($fromSide -eq 'bottom' -and $dy -gt 0) -or
          ($fromSide -eq 'top' -and $dy -lt 0)) {
        $oppositePenalty += 1.5
      }
      if (($toSide -eq 'left' -and $dx -lt 0) -or
          ($toSide -eq 'right' -and $dx -gt 0) -or
          ($toSide -eq 'bottom' -and $dy -lt 0) -or
          ($toSide -eq 'top' -and $dy -gt 0)) {
        $oppositePenalty += 1.5
      }
      $score += $oppositePenalty

      if ($IsLoopback -and $fromSide -in @('left', 'right') -and $toSide -in @('left', 'right') -and $LayoutDirection -in @('LR', 'RL')) {
        $score += 3.0
      }
      if ($IsLoopback -and $fromSide -in @('top', 'bottom') -and $toSide -in @('top', 'bottom') -and $LayoutDirection -in @('TB', 'BT')) {
        $score += 3.0
      }

      if ($RoutingIntent -eq 'clean') {
        $score += if ($fromSide -eq $preferred.fromSide -and $toSide -eq $preferred.toSide) { 0.0 } else { 1.0 }
      }

      if ($null -eq $best -or $score -lt $best.score) {
        $best = [pscustomobject]@{
          fromSide = $fromSide
          toSide = $toSide
          fromX = [double] $fromPort.glueX
          fromY = [double] $fromPort.glueY
          toX = [double] $toPort.glueX
          toY = [double] $toPort.glueY
          score = [double] $score
          length = [double] $routeLength
          obstacleHits = [int] $obstacleHits
          avoidancePenalty = [double] $avoidancePenalty
          labelHits = [int] $labelHits
          labelPenalty = [double] $labelPenalty
          boundaryHits = [int] $boundaryHits
          boundaryPenalty = [double] $boundaryPenalty
          crossingHits = [int] $crossingHits
          crossingPenalty = [double] $crossingPenalty
          routePoints = @($routePath.points)
          preferredFromSide = [string] $preferred.fromSide
          preferredToSide = [string] $preferred.toSide
        }
      }
    }
  }

  return $best
}

function Resolve-VisioRouteDecision {
  param(
    [Parameter(Mandatory = $true)] $Edge,
    [Parameter(Mandatory = $true)] $GraphModel,
    [Parameter(Mandatory = $true)] [string] $RoutingIntent,
    [System.Collections.Hashtable] $NodeIndex = @{},
    [array] $ExistingRoutes = @()
  )

  $fromId = [string] $Edge.from
  $toId = [string] $Edge.to
  $edgeRouteType = if ($Edge.routeType) { ([string] $Edge.routeType).ToLowerInvariant() } else { '' }
  $hasExplicitFromX = $null -ne $Edge.fromX
  $hasExplicitToX = $null -ne $Edge.toX
  $hasExplicitFromY = $null -ne $Edge.fromY
  $hasExplicitToY = $null -ne $Edge.toY
  $hasExplicitGlue = $hasExplicitFromX -or $hasExplicitFromY -or $hasExplicitToX -or $hasExplicitToY

  $diagramType = if ($GraphModel.diagramType) { [string] $GraphModel.diagramType } else { 'flowchart' }
  $defaults = Get-VisioRoutingDefault -DiagramType $diagramType

  $fromNode = $null
  $toNode = $null
  if ($NodeIndex.Count -gt 0) {
    if ($NodeIndex.ContainsKey($fromId)) { $fromNode = $NodeIndex[$fromId] }
    if ($NodeIndex.ContainsKey($toId)) { $toNode = $NodeIndex[$toId] }
  } else {
    foreach ($node in @($GraphModel.nodes)) {
      if ([string] $node.id -eq $fromId) { $fromNode = $node }
      if ([string] $node.id -eq $toId) { $toNode = $node }
    }
  }

  $isLoopback = $false
  $isCrossEdge = $false
  if ($null -ne $fromNode -and $null -ne $toNode) {
    $fromDepth = if ($fromNode.PSObject.Properties.Match('x').Count -gt 0 -and $null -ne $fromNode.x) { [double] $fromNode.x } else { 0 }
    $toDepth = if ($toNode.PSObject.Properties.Match('x').Count -gt 0 -and $null -ne $toNode.x) { [double] $toNode.x } else { 0 }
    if ($toDepth -lt $fromDepth) {
      $isLoopback = $true
    }

    $layoutDirection = if ($GraphModel.layout -and $GraphModel.layout.direction) {
      [string] $GraphModel.layout.direction
    } elseif ($GraphModel.layout -and $GraphModel.layout.strategy) {
      'LR'
    } else { 'LR' }

    if ($layoutDirection -in @('TB', 'BT')) {
      $fromRank = if ($null -ne $fromNode.y) { [double] $fromNode.y } else { 0 }
      $toRank = if ($null -ne $toNode.y) { [double] $toNode.y } else { 0 }
      if ($layoutDirection -eq 'TB' -and $toRank -gt $fromRank) { $isLoopback = $true }
      if ($layoutDirection -eq 'BT' -and $toRank -lt $fromRank) { $isLoopback = $true }
    }

    $fromSemantic = if ($fromNode.semanticType) { [string] $fromNode.semanticType } else { '' }
    $toSemantic = if ($toNode.semanticType) { [string] $toNode.semanticType } else { '' }
    if ($fromSemantic -in @('decision', 'gateway') -and $toSemantic -in @('process', 'task') -and $isLoopback) {
      $isCrossEdge = $true
    }
  }

  $resolvedRouteType = $edgeRouteType
  $resolvedFromX = if ($hasExplicitFromX) { [double] $Edge.fromX } else { $null }
  $resolvedFromY = if ($hasExplicitFromY) { [double] $Edge.fromY } else { $null }
  $resolvedToX = if ($hasExplicitToX) { [double] $Edge.toX } else { $null }
  $resolvedToY = if ($hasExplicitToY) { [double] $Edge.toY } else { $null }
  $resolvedFromSide = ''
  $resolvedToSide = ''
  $routeScore = $null
  $routeLength = $null
  $obstacleHits = $null
  $avoidancePenalty = $null
  $labelHits = $null
  $labelPenalty = $null
  $boundaryHits = $null
  $boundaryPenalty = $null
  $crossingHits = $null
  $crossingPenalty = $null
  $routePoints = @()
  $preferredFromSide = ''
  $preferredToSide = ''

  switch ($RoutingIntent) {
    'fidelity' {
      if (-not $edgeRouteType) {
        $resolvedRouteType = if ($defaults.defaultRouteType -eq 'preserve') { 'orthogonal' } else { [string] $defaults.defaultRouteType }
      }
      if (-not $hasExplicitGlue -and $isLoopback -and $null -ne $fromNode -and $null -ne $toNode) {
        $lbSide = if ($defaults.loopbackFromSide -eq 'preserve') { 'auto' } else { [string] $defaults.loopbackFromSide }
        $ltSide = if ($defaults.loopbackToSide -eq 'preserve') { 'auto' } else { [string] $defaults.loopbackToSide }
        $lbGlue = Resolve-VisioLoopbackGlue -FromNode $fromNode -ToNode $toNode -LayoutDirection $(if ($GraphModel.layout -and $GraphModel.layout.direction) { [string] $GraphModel.layout.direction } else { 'LR' }) -LoopbackFromSide $lbSide -LoopbackToSide $ltSide
        $resolvedFromX = [double] $lbGlue.fromX
        $resolvedFromY = [double] $lbGlue.fromY
        $resolvedToX = [double] $lbGlue.toX
        $resolvedToY = [double] $lbGlue.toY
      }
    }

    'clean' {
      if ($isLoopback) {
        if ($edgeRouteType) {
          $resolvedRouteType = $edgeRouteType
        } else {
          $resolvedRouteType = if ($defaults.loopbackRouteType -eq 'preserve') { 'curved' } else { [string] $defaults.loopbackRouteType }
        }
        if (-not $hasExplicitGlue -and $null -ne $fromNode -and $null -ne $toNode) {
          $lbSide = if ($defaults.loopbackFromSide -eq 'preserve') { 'auto' } else { [string] $defaults.loopbackFromSide }
          $ltSide = if ($defaults.loopbackToSide -eq 'preserve') { 'auto' } else { [string] $defaults.loopbackToSide }
          $lbGlue = Resolve-VisioLoopbackGlue -FromNode $fromNode -ToNode $toNode -LayoutDirection $(if ($GraphModel.layout -and $GraphModel.layout.direction) { [string] $GraphModel.layout.direction } else { 'LR' }) -LoopbackFromSide $lbSide -LoopbackToSide $ltSide
          $resolvedFromX = [double] $lbGlue.fromX
          $resolvedFromY = [double] $lbGlue.fromY
          $resolvedToX = [double] $lbGlue.toX
          $resolvedToY = [double] $lbGlue.toY
        }
      } elseif ($isCrossEdge -and $defaults.crossEdgesPreferCurved -and -not $edgeRouteType) {
        $resolvedRouteType = 'curved'
      } elseif ($edgeRouteType) {
        $resolvedRouteType = $edgeRouteType
      } else {
        $resolvedRouteType = [string] $defaults.defaultRouteType
      }
    }

    default {
      if ($edgeRouteType) {
        $resolvedRouteType = $edgeRouteType
      } elseif ($isLoopback) {
        $resolvedRouteType = if ($defaults.loopbackRouteType -eq 'preserve') { 'curved' } else { [string] $defaults.loopbackRouteType }
        if (-not $hasExplicitGlue -and $null -ne $fromNode -and $null -ne $toNode) {
          $lbSide = if ($defaults.loopbackFromSide -eq 'preserve') { 'auto' } else { [string] $defaults.loopbackFromSide }
          $ltSide = if ($defaults.loopbackToSide -eq 'preserve') { 'auto' } else { [string] $defaults.loopbackToSide }
          $lbGlue = Resolve-VisioLoopbackGlue -FromNode $fromNode -ToNode $toNode -LayoutDirection $(if ($GraphModel.layout -and $GraphModel.layout.direction) { [string] $GraphModel.layout.direction } else { 'LR' }) -LoopbackFromSide $lbSide -LoopbackToSide $ltSide
          $resolvedFromX = [double] $lbGlue.fromX
          $resolvedFromY = [double] $lbGlue.fromY
          $resolvedToX = [double] $lbGlue.toX
          $resolvedToY = [double] $lbGlue.toY
        }
      } else {
        $resolvedRouteType = [string] $defaults.defaultRouteType
      }
    }
  }

  if ($resolvedRouteType -notin @('orthogonal', 'straight', 'curved')) {
    $resolvedRouteType = 'orthogonal'
  }

  if (-not $hasExplicitGlue -and $null -ne $fromNode -and $null -ne $toNode) {
    $layoutDirectionForPorts = if ($GraphModel.layout -and $GraphModel.layout.direction) {
      [string] $GraphModel.layout.direction
    } else {
      'LR'
    }
    $obstacleNodes = @()
    foreach ($node in @($GraphModel.nodes)) {
      if ($null -eq $node -or -not $node.id) { continue }
      $nodeId = [string] $node.id
      if ($nodeId -ne $fromId -and $nodeId -ne $toId) {
        $obstacleNodes += @($node)
      }
    }

    $labelBounds = @()
    foreach ($collectionName in @('annotations', 'labels', 'texts')) {
      if ($GraphModel.PSObject.Properties.Match($collectionName).Count -eq 0) { continue }
      foreach ($textItem in @($GraphModel.$collectionName)) {
        if ($null -eq $textItem) { continue }
        $labelBounds += @(Expand-VisioBounds -Bounds (Get-VisioGraphTextBounds -TextItem $textItem) -Padding 0.06)
      }
    }

    $boundaryBounds = @()
    $fromCenter = [pscustomobject]@{ x = [double] $fromNode.x; y = [double] $fromNode.y }
    $toCenter = [pscustomobject]@{ x = [double] $toNode.x; y = [double] $toNode.y }
    foreach ($collectionName in @('containers', 'swimlanes', 'lanes', 'pools', 'groups')) {
      if ($GraphModel.PSObject.Properties.Match($collectionName).Count -eq 0) { continue }
      foreach ($region in @($GraphModel.$collectionName)) {
        if ($null -eq $region) { continue }
        $regionBounds = Get-VisioGraphRegionBounds -Region $region
        if ((Test-VisioPointInsideBounds -Point $fromCenter -Bounds $regionBounds) -or
            (Test-VisioPointInsideBounds -Point $toCenter -Bounds $regionBounds)) {
          continue
        }
        $boundaryBounds += @(Expand-VisioBounds -Bounds $regionBounds -Padding 0.04)
      }
    }

    $optimalGlue = Resolve-VisioOptimalConnectorGlue -FromNode $fromNode -ToNode $toNode -LayoutDirection $layoutDirectionForPorts -IsLoopback:$isLoopback -RoutingIntent $RoutingIntent -ObstacleNodes $obstacleNodes -LabelBounds $labelBounds -BoundaryBounds $boundaryBounds -ExistingRoutes $ExistingRoutes
    if ($null -ne $optimalGlue) {
      $resolvedFromX = [double] $optimalGlue.fromX
      $resolvedFromY = [double] $optimalGlue.fromY
      $resolvedToX = [double] $optimalGlue.toX
      $resolvedToY = [double] $optimalGlue.toY
      $resolvedFromSide = [string] $optimalGlue.fromSide
      $resolvedToSide = [string] $optimalGlue.toSide
      $routeScore = [double] $optimalGlue.score
      $routeLength = [double] $optimalGlue.length
      $obstacleHits = [int] $optimalGlue.obstacleHits
      $avoidancePenalty = [double] $optimalGlue.avoidancePenalty
      $labelHits = [int] $optimalGlue.labelHits
      $labelPenalty = [double] $optimalGlue.labelPenalty
      $boundaryHits = [int] $optimalGlue.boundaryHits
      $boundaryPenalty = [double] $optimalGlue.boundaryPenalty
      $crossingHits = [int] $optimalGlue.crossingHits
      $crossingPenalty = [double] $optimalGlue.crossingPenalty
      $routePoints = @($optimalGlue.routePoints)
      $preferredFromSide = [string] $optimalGlue.preferredFromSide
      $preferredToSide = [string] $optimalGlue.preferredToSide
    }
  }

  return [pscustomobject]@{
    routeType = $resolvedRouteType
    fromX = $resolvedFromX
    fromY = $resolvedFromY
    toX = $resolvedToX
    toY = $resolvedToY
    fromSide = $resolvedFromSide
    toSide = $resolvedToSide
    routeScore = $routeScore
    routeLength = $routeLength
    obstacleHits = $obstacleHits
    avoidancePenalty = $avoidancePenalty
    labelHits = $labelHits
    labelPenalty = $labelPenalty
    boundaryHits = $boundaryHits
    boundaryPenalty = $boundaryPenalty
    crossingHits = $crossingHits
    crossingPenalty = $crossingPenalty
    routePoints = @($routePoints)
    preferredFromSide = $preferredFromSide
    preferredToSide = $preferredToSide
    isLoopback = $isLoopback
    isCrossEdge = $isCrossEdge
    routingIntent = $RoutingIntent
  }
}
