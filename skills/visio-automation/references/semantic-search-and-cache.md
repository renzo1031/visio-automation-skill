# Semantic Search And Session Cache

Use this file when the user describes diagram concepts in natural language, especially Chinese terms, or when editing the same diagram across multiple turns.

## Semantic Master Search

Use `Find-SemanticVisioMaster` to map intent to Visio masters before falling back to raw stencil scanning.

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"

# Natural language to master mapping
Find-SemanticVisioMaster -Query '数据库' | Format-Table -AutoSize
# Returns: BASFLO_M.VSSX / Database

Find-SemanticVisioMaster -Query 'gateway' -Category 'bpmn' | Format-Table -AutoSize
# Returns: BPMN_M.VSSX / Gateway

Find-SemanticVisioMaster -Query '库存' -Category 'vsm' | Format-Table -AutoSize
# Returns: VSM_M.VSSX / Inventory
```

The semantic map lives at `references/semantic-master-map.json` and covers flowchart, BPMN, DFD, basic shapes, and VSM categories with Chinese/English aliases.

Use `-Category` to narrow the search when the diagram family is known, and `-IncludeFallback` to also run `Find-VisioMasters` when the semantic map has no match.

## Session Cache

For连续编辑同一工作区的图时，使用会话缓存避免重复 discovery。缓存存储在工作区的 `.cache/visio-automation/session-cache.json`。

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"

# Save diagram context after rendering
Update-VisioSessionCache -DiagramType 'flowchart' -Stencil 'BASFLO_M.VSSX' `
  -LastOutputPath 'E:\output\diagram.vsdx' -LastLayoutStrategy 'Flowchart' -LastLayoutDirection 'LR'

# Track recently used masters for quick re-selection
Add-VisioSessionRecentMaster -Stencil 'BASFLO_M.VSSX' -MasterNameU 'Process'
Add-VisioSessionRecentMaster -Stencil 'BPMN_M.VSSX' -MasterNameU 'Gateway'

# Load cached state for next operation
$cache = Get-VisioSessionCache
# $cache.diagramType, $cache.stencil, $cache.recentMasters, etc.
```

Session cache fields:

- `diagramType`
- `stencil`
- `masterMapping`
- `lastOutputPath`
- `lastPreviewPath`
- `lastLayoutStrategy`
- `lastLayoutDirection`
- `recentMasters` as `{stencil, master, usedAt}`

Cache files are workspace-local and should be gitignored.
