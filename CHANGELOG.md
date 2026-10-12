# Changelog

All notable changes to **QuotaRecon** are tracked here. Format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and versions follow
[Semantic Versioning](https://semver.org/) (`MAJOR.MINOR.PATCH`).

The current version is also printed at the top of every script run and is
available via:

```powershell
./Invoke-QuotaRecon.ps1 -Version
```

---

## [1.6.0] - 2026-10-12

### Added

- **`-IncludeCognitiveServicesQuota` / `-OnlyCognitiveServicesQuota`
  switches** — the second service onboarded onto the `-Include<Service>` /
  `-Only<Service>` pattern. Adds a `CognitiveServicesQuota` sheet with
  Azure Cognitive Services / Azure OpenAI / AI Foundry quota for every
  target Sub x Region: account limits plus every per-model-family quota
  item (provisioned throughput units, global provisioned throughput,
  fine-tuned deployment counts, etc). Reads the native
  `Microsoft.CognitiveServices/locations/{region}/usages` endpoint, which
  is the same row shape `Microsoft.Compute/usages` uses, so it reuses the
  existing `Convert-QRUsageRow` helper directly rather than needing a
  Microsoft.Quota merge like Container Apps. `-OnlyCognitiveServicesQuota`
  follows the same fast-path contract: implies the Include switch, skips
  every unrelated subsystem, and trims the workbook to
  `CognitiveServicesQuota`, `Inputs`, `Errors`.
- Multiple `-Only<Service>` switches can now be combined in one run (e.g.
  `-OnlyContainerAppsQuota -OnlyCognitiveServicesQuota`) — the workbook is
  trimmed to the union of each service's sheets, and everything unrelated
  to any of them is still skipped.

## [1.5.0] - 2026-10-12

### Added

- **`-OnlyContainerAppsQuota` switch** — fast-path mode that implies
  `-IncludeContainerAppsQuota` and skips everything unrelated: SKU
  resolution/prompt, Microsoft.Compute provider checks, the zone map,
  Group Quotas, VM Inventory, and the Summary/Quota/Restrictions
  collection. Only Subscriptions and Regions are needed.
- **Output workbook trimming for `-Only<Service>` switches** —
  `Write-QRWorkbook` now accepts an `-OnlySheets` list; when non-empty,
  every sheet not in that list (other than the always-present
  `Inputs`/`Errors`) is omitted entirely rather than written out empty.
  `-OnlyContainerAppsQuota` uses this to produce a minimal 4-sheet
  workbook: `ContainerAppsQuota`, `ContainerAppsEnvQuota`, `Inputs`,
  `Errors`. This is the template mechanism for future `-Only<Service>`
  switch additions — see `.github/copilot-instructions.md`.

### Changed

- **`QuotaGroupLimits` and `QuotaGroupAllocations` sheets are now omitted
  entirely** from the output workbook when quota groups weren't collected
  (`-IncludeQuotaGroups` / `-QuotaGroupMgIds` not passed), instead of being
  included empty with no header row. No change when those switches are
  used — both sheets are populated as before.

## [1.4.0] - 2026-10-12

### Changed

- **`-IncludeContainerAppsQuota` now collects BOTH Container Apps quota
  sheets** — subscription/region-scoped quota (`ContainerAppsQuota`) and
  per-managed-environment workload-profile core quota
  (`ContainerAppsEnvQuota`: Consumption, General Purpose/D-series, Memory
  Optimized/E-series, GPU cores), discovered via one Azure Resource Graph
  query (same approach as `-IncludeVmInventory`) plus one usages call per
  environment found. This is what actually answers "how many D-series
  cores are available" — that number lives on a specific managed
  environment, not the subscription or region. One switch now covers the
  whole feature instead of splitting it across two opt-ins, since users
  think of "Container Apps quota" as one thing.
- If no managed environments exist for the specified
  subscriptions/regions, that's now noted as a row on the `Errors` sheet
  (not just a console message) so it's visible to anyone who only opens
  the workbook. `ContainerAppsQuota` is still collected regardless; no
  error is raised.

## [1.3.0] - 2026-10-12

### Added

- **`-IncludeContainerAppsQuota` switch** — adds a `ContainerAppsQuota`
  sheet with Azure Container Apps subscription/region-scoped quota
  (Managed Environment Count, Session Pools, Subscription NCA100 GPUs) for
  every target Sub x Region. Reads the Microsoft.Quota provider under the
  Microsoft.App namespace — no new module or CLI dependency. Does **not**
  cover per-managed-environment workload-profile core quota (Consumption,
  General Purpose/D-series, Memory Optimized/E-series, GPU); those require
  enumerating Container Apps environments and are planned for a future
  release.

## [1.2.0] - 2026-10-01

### Added

- **`-IncludeAllSKUs` switch** — enumerates every virtualMachine family
  present in the union of the target regions' `Microsoft.Compute/skus`
  catalogs, so the Summary sheet covers everything Azure offers in those
  regions without needing a manual SKU list. Catalog-enumerated rows show
  `SampleSizeSource = catalog` on the Inputs sheet. When mixed with
  user-supplied SKUs, the user's entries keep their sample size on
  overlapping families.

## [1.1.2] - 2026-10-01

### Fixed

- **Summary zone cells now distinguish "not offered" from "unrestricted".**
  Previously, logical zones that weren't in a SKU's `locationInfo[].zones`
  were shown as `None` (green), implying availability. They now read `N/A`
  (uncoloured). Example: `Standard_DC2s_v3` in `southeastasia` is only
  offered in zone 1 — AZ1 shows `None`, AZ2 and AZ3 now show `N/A`.

## [1.1.1] - 2026-10-01

### Fixed

- **Null-array crash** on the first `(Sub × Region × Family)` iteration
  caused by the v1.1.0 refactor dropping the `$azStatus = @{}` initializer.

### Changed

- **Family-aware probe for the Summary sheet.** For each
  `(Sub × Region × Family)` the script now enumerates every size in the
  family that's offered in that region's catalog, not just the pre-picked
  sample. Added two new Summary columns:
  - `OfferedSizesInRegion` — count of family sizes in the region's catalog
    (`0` means the family isn't offered there at all).
  - `ProbeSampleUsed` — the actual size used for the restriction probe.
    If the pre-picked sample isn't offered in the region, QuotaRecon falls
    back to the smallest sibling that is offered and notes the swap.

## [1.1.0] - 2026-10-01

### Added

- **Script versioning.** `$script:QRVersion` constant printed at the top of
  every run and exposed via a new `-Version` switch.
- **Invalid-input callouts.** Unresolved subscription GUIDs, regions, and
  SKUs/families are listed to the console as a `WARNING` before collection
  begins, in addition to the `Inputs` sheet. Prevents silent typos.

### Changed

- **Group Quotas are now opt-in** via `-IncludeQuotaGroups`
  (or `-QuotaGroupMgIds` which implies opt-in). Previously enabled by
  default. `-SkipQuotaGroups` is retained as a no-op for backward
  compatibility.
- **Summary restriction cells show `N/A`** (uncoloured) instead of `None`
  (green) when the sample size isn't offered in a specific region — so a
  green cell always reflects real catalog data.

## [1.0.0] - 2026-09

### Added

- Initial public release.
- Cross-subscription Azure compute quota, SKU-restriction, and
  availability-zone-mapping audit, output as a single `.xlsx`.
- Modular script layout under `src/` (Auth, Input, Skus, Quota, Zones,
  QuotaGroups, Inventory, Excel).
- Opt-in VM inventory (`-IncludeVmInventory`) via Azure Resource Graph.
- Opt-in diagnostic transcript (`-DiagnosticLog`).
- Templates under `templates/` and read-only reference catalogs under
  `references/`.

[1.6.0]: https://github.com/MS-mikeo/QuotaRecon/releases/tag/v1.6.0
[1.5.0]: https://github.com/MS-mikeo/QuotaRecon/releases/tag/v1.5.0
[1.4.0]: https://github.com/MS-mikeo/QuotaRecon/releases/tag/v1.4.0
[1.3.0]: https://github.com/MS-mikeo/QuotaRecon/releases/tag/v1.3.0
[1.2.0]: https://github.com/MS-mikeo/QuotaRecon/releases/tag/v1.2.0
[1.1.2]: https://github.com/MS-mikeo/QuotaRecon/releases/tag/v1.1.2
[1.1.1]: https://github.com/MS-mikeo/QuotaRecon/releases/tag/v1.1.1
[1.1.0]: https://github.com/MS-mikeo/QuotaRecon/releases/tag/v1.1.0
[1.0.0]: https://github.com/MS-mikeo/QuotaRecon/releases/tag/v1.0.0
