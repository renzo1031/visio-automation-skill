# Official Visio Automation References

Prefer Microsoft Learn pages for API behavior and ShapeSheet cell values. Use these links when a script needs a method or cell that is not already covered by the helper script.

## Core COM/VBA Object Model

- [Visio VBA reference](https://learn.microsoft.com/en-us/office/vba/api/overview/visio)
- [Application object](https://learn.microsoft.com/en-us/office/vba/api/visio.application)
- [Documents.OpenEx method](https://learn.microsoft.com/en-us/office/vba/api/visio.documents.openex)
- [Page.Drop method](https://learn.microsoft.com/en-us/office/vba/api/visio.page.drop)
- [Shape.AutoConnect method](https://learn.microsoft.com/en-us/office/vba/api/Visio.Shape.AutoConnect)
- [Cell.GlueToPos method](https://learn.microsoft.com/en-us/office/vba/api/visio.cell.gluetopos)
- [Page.Export method](https://learn.microsoft.com/en-us/office/vba/api/visio.page.export)

Useful notes:

- `Documents.OpenEx(file, 66)` opens a stencil hidden and read-only in the tested workflow. It combines `visOpenHidden = 64` and `visOpenRO = 2`.
- `Page.Drop(master, x, y)` places a master with its pin at page coordinates.
- `Cell.GlueToPos(shape, xPercent, yPercent)` glues a connector endpoint to a relative position on a target shape and creates a connection point if needed.
- `Shape.AutoConnect(toShape, direction, connector)` is useful for quick right/left/up/down layout, but manual `GlueToPos` is better when matching a reference image.

## ShapeSheet Cells Used Most Often

- [ShapeSheet reference](https://learn.microsoft.com/en-us/office/client-developer/visio/visio-shapesheet-reference)
- [ShapeRouteStyle cell](https://learn.microsoft.com/en-us/office/client-developer/visio/shaperoutestyle-cell-shape-layout-section)
- [ConLineRouteExt cell](https://learn.microsoft.com/en-us/office/client-developer/visio/conlinerouteext-cell-shape-layout-section)
- [ConFixedCode cell](https://learn.microsoft.com/en-us/office/client-developer/visio/confixedcode-cell-shape-layout-section)
- [EndArrow cell](https://learn.microsoft.com/en-us/office/client-developer/visio/endarrow-cell-line-format-section)
- [Line Format section](https://learn.microsoft.com/en-us/office/client-developer/visio/line-format-section)
- [Shape Layout section](https://learn.microsoft.com/en-us/office/client-developer/visio/shape-layout-section)
- [Shape Transform section](https://learn.microsoft.com/en-us/office/client-developer/visio/shape-transform-section)

Connector values used in the local workflow:

| Cell | Value | Meaning |
| --- | ---: | --- |
| `ShapeRouteStyle` | `2` | Straight route (`visLORouteStraight`) |
| `ConLineRouteExt` | `1` | Straight connector appearance (`visLORouteExtStraight`) |
| `ConLineRouteExt` | `2` | Curved connector appearance (`visLORouteExtNURBS`) |
| `EndArrow` | `0` | No arrowhead |
| `EndArrow` | `1` to `45` | Indexed arrowhead styles |

## External Documentation Fallback

When Microsoft Learn is insufficient:

1. Search official docs first:
   - `site:learn.microsoft.com/en-us/office/vba/api/visio <method or object>`
   - `site:learn.microsoft.com/en-us/office/client-developer/visio <ShapeSheet cell>`
2. If official docs do not cover the behavior, search reputable examples for the exact COM/VBA method name.
3. Prefer validating any third-party claim with a tiny local Visio script before using it in the user-facing diagram.
4. Record newly verified masters or ShapeSheet values in `references/master-catalog.md` or a new reference file.

## Local Inspection Snippets

List all stencils under a typical Office 16 install:

```powershell
Get-ChildItem -Path "$env:ProgramFiles\Microsoft Office\root\Office16\Visio Content" `
  -Recurse -Include '*.vssx','*.vss','*.vstx' |
  Select-Object FullName
```

List masters for a stencil:

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"
Get-VisioMasters -StencilNameOrPath 'DATFLO_M.VSSX' | Format-Table -AutoSize
```

Search all installed stencil masters by keyword:

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"
Find-VisioMasters -Query 'Gateway|网关' -PreferredStencilRegex 'BPMN' | Format-Table -AutoSize
```

By default the helper searches `*.vssx` and `*.vss` stencil files. Include `*.vstx` templates only when template discovery is required.

Check connector route settings in an open document:

```powershell
$doc = Get-OpenVisioDocumentByPath 'E:\path\diagram.vsdx'
$page = $doc.Pages.Item(1)
foreach ($shape in @($page.Shapes)) {
  if ($shape.OneD -ne 0) {
    [pscustomobject]@{
      ID = $shape.ID
      NameU = $shape.NameU
      ShapeRouteStyle = $shape.CellsU('ShapeRouteStyle').FormulaU
      ConLineRouteExt = $shape.CellsU('ConLineRouteExt').FormulaU
    }
  }
}
```
