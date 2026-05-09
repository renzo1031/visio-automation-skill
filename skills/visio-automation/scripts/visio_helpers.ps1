$ErrorActionPreference = 'Stop'

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
    $stencil.Close()
  } finally {
    $visio.Quit()
  }
}

function Find-VisioMasters {
  param(
    [Parameter(Mandatory = $true)] [string] $Query,
    [string[]] $StencilPattern = @('*.vssx', '*.vss'),
    [string] $PreferredStencilRegex = '',
    [int] $MaxResults = 50
  )

  $visio = New-InvisibleVisioApplication
  $results = New-Object System.Collections.Generic.List[object]
  $exactTerms = $Query -split '\|' | ForEach-Object { $_.Trim(' ', '^', '$') } | Where-Object { $_ }
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
              Stencil = $file.Name
              FullName = $file.FullName
              Name = $master.Name
              NameU = $master.NameU
            })
          }
        }
        $stencil.Close()
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
    $visio.Quit()
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

  $Shape.CellsU('Char.Color').FormulaU = $Color
  $Shape.CellsU('Char.Size').FormulaU = "$Size pt"
  $Shape.CellsU('Char.Style').FormulaU = "$Bold"
  $Shape.CellsU('Para.HorzAlign').FormulaU = '1'
  $Shape.CellsU('VerticalAlign').FormulaU = '1'
}

function Set-VisioShapeFill {
  param(
    [Parameter(Mandatory = $true)] $Shape,
    [Parameter(Mandatory = $true)] [string] $Fill,
    [string] $Line = $Fill,
    [string] $LineWeight = '1 pt'
  )

  $Shape.CellsU('FillForegnd').FormulaU = $Fill
  $Shape.CellsU('LineColor').FormulaU = $Line
  $Shape.CellsU('LineWeight').FormulaU = $LineWeight
}

function Connect-VisioShapesStraight {
  param(
    [Parameter(Mandatory = $true)] $Page,
    [Parameter(Mandatory = $true)] $ConnectorMaster,
    [Parameter(Mandatory = $true)] $From,
    [Parameter(Mandatory = $true)] $To,
    [double] $FromX = 1.0,
    [double] $FromY = 0.5,
    [double] $ToX = 0.0,
    [double] $ToY = 0.5,
    [int] $EndArrow = 4,
    [string] $LineColor = 'RGB(45,45,45)'
  )

  $connector = $Page.Drop($ConnectorMaster, 0, 0)
  $connector.CellsU('BeginX').GlueToPos($From, $FromX, $FromY)
  $connector.CellsU('EndX').GlueToPos($To, $ToX, $ToY)
  $connector.CellsU('ShapeRouteStyle').FormulaU = '2'
  $connector.CellsU('ConLineRouteExt').FormulaU = '1'
  $connector.CellsU('LineColor').FormulaU = $LineColor
  $connector.CellsU('LineWeight').FormulaU = '1 pt'
  $connector.CellsU('EndArrow').FormulaU = "$EndArrow"
  return $connector
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
  $label.CellsU('LinePattern').FormulaU = '0'
  $label.CellsU('FillPattern').FormulaU = '0'
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
