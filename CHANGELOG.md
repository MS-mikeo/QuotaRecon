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

[1.2.0]: https://github.com/MS-mikeo/QuotaRecon/releases/tag/v1.2.0
[1.1.2]: https://github.com/MS-mikeo/QuotaRecon/releases/tag/v1.1.2
[1.1.1]: https://github.com/MS-mikeo/QuotaRecon/releases/tag/v1.1.1
[1.1.0]: https://github.com/MS-mikeo/QuotaRecon/releases/tag/v1.1.0
[1.0.0]: https://github.com/MS-mikeo/QuotaRecon/releases/tag/v1.0.0
