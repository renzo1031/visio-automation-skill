# Master Discovery

Use this file when the skill does not know the needed shape or the local Visio install differs from the catalog.

## Recovery Path

Do not fall back immediately to drawing raw rectangles or lines. Use this recovery path:

1. Search the local master catalog in `references/master-catalog.md`.
2. If nothing matches, build or reuse a local master index, then search installed Visio stencil files through the index first:

   ```powershell
   . "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"
   $indexPath = Join-Path $env:USERPROFILE '.codex\cache\visio-master-index.json'
   Build-VisioMasterIndex -OutputPath $indexPath | Out-Null
   Find-VisioMasters -Query 'Gateway|网关' -PreferredStencilRegex 'BPMN' -IndexPath $indexPath | Format-Table -AutoSize
   ```

3. Use `PreferredStencilRegex` for domain terms (`BPMN`, `VSM`, `UML`, `AWS`, `Azure`, `BASFLO`) so discovery does not waste time scanning every cloud icon stencil first.
4. If discovery finds a master, open that stencil and use the discovered `NameU`.
5. If discovery does not find a shape, search Microsoft Learn for the API behavior and search installed Visio content by likely domain terms.
6. Create a tiny local proof file with the discovered master or API call.
7. If the proof works and the shape is likely reusable, append it to `references/master-catalog.md`.
8. Only use primitive `BASIC_*` shapes as a final fallback, and state that a direct built-in master was not found.

## Discovery Performance Notes

- Default discovery searches only stencil files (`*.vssx`, `*.vss`). Templates (`*.vstx`) can be slow or inappropriate for master enumeration, so include them only when the task explicitly asks for templates.
- `Build-VisioMasterIndex` records the local Visio version, language, content roots, stencil file timestamps, stencil names, paths, and master `Name`/`NameU`.
- `Find-VisioMasters -IndexPath ...` uses a valid index first and falls back to live scanning when the index is absent, stale, or has no match.
- Rebuild the index with `-Force` after Visio content is installed, removed, repaired, or language packs change.
- If a broad query like `Gateway` returns unrelated cloud/network masters, rerun it with a domain-specific `PreferredStencilRegex`, for example `BPMN`.
- If a timed-out discovery leaves invisible `VISIO.EXE` processes, close only blank-title background Visio processes; do not close a user's visible document window.
