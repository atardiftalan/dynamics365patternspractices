# Business Process Catalog Azure DevOps setup package - July

This July package combines the Azure DevOps setup scripts, the resumable Business Process Catalog work item importer, and a deterministic HTML run summary report.

Use this package to create or update the Azure DevOps process/project configuration and import the Business Process Catalog source workbooks into Azure DevOps Boards.

## Phases

| Phase | Purpose | Script |
| ---: | --- | --- |
| 1 | Create process, project, work item types, fields, picklists, and Test Case `New` state | `1_ADO_Creation_Script.py` |
| 2 | Configure work item page layouts and DevLabs Multivalue controls | `2_ADO_Page_Layout_Script_Threaded.py` |
| 3 | Create teams, area paths, and team area assignments | `3_ADO_Teams_Areas_Script.py` |
| 4 | Configure backlog levels, iterations, and team settings | `4_ADO_Backlog_Config_Script.py` |
| 5 | Import Business Process Catalog work items | `5_BPC_Catalog_Import.py` |
| 6 | Generate deterministic HTML setup/import summary report | `6_Generate_HTML_Report.py` |
| 7 | Update existing Business Process Catalog work items from source files | `7_BPC_Catalog_Update.py` |

The package also includes `8_Generate_Guideline.py`, a read-only utility that exports an existing Azure DevOps project/process configuration to a July-compatible guideline workbook. It is intentionally separate from the mutating setup phases.

## Install dependencies

Python 3.12 or later is recommended.

```powershell
python -m pip install -r requirements.txt
```

> [!TIP]
> A Python virtual environment is optional. Use one if you want to isolate this package's dependencies from other Python tools on the machine. For Windows guidance, see [Creation of virtual environments](https://docs.python.org/3/library/venv.html#creating-virtual-environments). After creating and activating a virtual environment, use the same commands shown in this README.

## Run setup, import, and report phases

```powershell
python setup_wizard.py `
  --ado-org-url "https://dev.azure.com/<organization>" `
  --ado-project "<project name>" `
  --process-name "<process name>" `
  --excel-file "<path to ADO template guideline workbook>" `
  --catalog-source-dir "<folder containing one or more catalog source files>" `
  --catalog-output ".\out" `
  --catalog-parallel-workers 4
```

The wizard prompts for the PAT if `BPC_ADO_PAT` is not already set. The PAT is kept in memory for the run and is not written to script files.

For large imports, start with 2-8 parallel workers. Higher worker counts may trigger Azure DevOps ATCPU throttling. The importer retries transient HTTP 408, 429, and 5xx responses.

The wizard defaults to phases 1-6. Phase 7 updates existing work items and must be run explicitly.

## Run selected phases

Rerun only phases 1 and 2:

```powershell
python setup_wizard.py --start-at 1 --stop-after 2
```

Run only phase 5 import:

```powershell
python setup_wizard.py --start-at 5 --stop-after 5
```

Regenerate only the HTML summary report:

```powershell
python setup_wizard.py --start-at 6 --stop-after 6
```

Preview catalog updates without changing Azure DevOps:

```powershell
python setup_wizard.py --start-at 7 --stop-after 7
```

Apply catalog updates after reviewing the preview:

```powershell
python setup_wizard.py --start-at 7 --stop-after 7 --catalog-update-apply
```

## Export an existing Azure DevOps configuration

Use the read-only exporter to create a guideline workbook from an existing project. The project normally identifies its process automatically; pass `--process-name` only to override that selection.

```powershell
python 8_Generate_Guideline.py `
  --ado-org-url "https://dev.azure.com/<organization>" `
  --ado-project "<existing project name>" `
  --output ".\out\<project>-ADO-guideline-July.xlsx"
```

The command prompts for the PAT if `BPC_ADO_PAT` is not set. The setup PAT described below can be reused. For a read-only token, grant access to read projects/teams and work-item/process configuration.

The exporter reads work item types, assigned fields, picklists, form layouts, backlog behaviors, areas, iterations, teams, and team settings. It writes the seven configuration sheets consumed by phases 1-4 and preserves the July workbook's formatting. Review the generated workbook before using it as setup input, especially these values that Azure DevOps does not store in the same form as the guideline:

- business-purpose and field-use recommendation text,
- guideline-only rule columns,
- historical `Rename from` intent for custom backlog levels.

For fields already present in the supplied July template, guideline-only metadata is preserved by reference name. Unknown fields are exported with safe defaults and blank guideline-only metadata. Area paths deeper than the four levels supported by the July setup scripts fail with a clear error instead of producing a lossy workbook.

## Phase 5 import behavior

Phase 5 uses the same Azure DevOps organization, project, PAT, and template workbook from the wizard. It calls the July importer with:

- parent-aware parallel creation,
- project-scoped output folders,
- retry handling for transient connection failures,
- idempotent resume through `ado-id-map.csv`,
- continued processing after individual work item failures,
- dynamic work item type reference resolution,
- Test Case create fallback when Azure DevOps rejects custom state values at create time,
- deprecated/deleted source rows skipped by default.

By default, Phase 5 records individual work item failures in `import-failures.json` and continues with other importable rows. To stop after the first individual failure, run the wizard with `--catalog-fail-fast` or set `BPC_ADO_IMPORT_CONTINUE_ON_ERROR=0` before running `5_BPC_Catalog_Import.py` directly.

If the template workbook is in a folder named `Python Scripts`, the catalog source folder defaults to that folder's parent. Otherwise pass `--catalog-source-dir`.

### Output files

Phase 5 writes project-scoped output under:

```text
out\<organization>_<project>\
```

Key output files:

- `ado-id-map.csv` - successful imported or recovered work item IDs. Reruns skip these keys.
- `import-plan.json` - deterministic import plan and field payloads.
- `import-preview.csv` - human-readable preview of planned work items.
- `import-failures.json` - unresolved failure details from the latest failed run. The Phase 6 report reconciles this with `ado-id-map.csv` so resolved prior failures do not keep the report in a failed state.
- `skipped-deprecated-deleted.csv` - source rows skipped by create/import mode.
- `bpc-ado-setup-summary.html` - Phase 6 HTML summary report.

## Phase 7 update behavior

Phase 7 is separate from Phase 5. Phase 5 keeps its create/import and resume behavior. Phase 7 reads the same catalog source files, matches existing work items by `ado-id-map.csv`, and can update existing work item fields.

By default, Phase 7 is a dry run. It reads current Azure DevOps work item values and writes an update plan and results without changing Azure DevOps. Pass `--catalog-update-apply` to the wizard to apply updates.

Phase 7 updates only catalog-owned fields by default:

- `MSBPC.*` fields,
- `System.Title`,
- `System.Description`.

It excludes project-management and workflow fields such as state, reason, assigned-to, area path, iteration path, tags, and test-step fields unless the lower-level importer CLI is run with explicit field options.

Phase 7 writes project-scoped output under the same output folder as Phase 5:

- `update-plan.json` - field-level update plan and desired/current values.
- `update-results.csv` - per-row planned, updated, unchanged, skipped, or failed status.
- `update-failures.json` - update failure details when failures occur.

## Test Case New state

Phase 1 ensures Test Case work item type references have a `New` state in the `Proposed` state category. This supports source files that use `State = New` for standard Test Case work items.

Phase 5 also resolves work item type display names to the current project's Azure DevOps work item type reference names. This avoids hard-coded process/project prefixes and supports projects where the process-specific Test Case reference differs.

## DevLabs multivalue control extension

Phase 2 attempts to install or confirm the DevLabs multivalue control extension for the Azure DevOps organization. If the extension is already installed, the script continues. If the PAT or organization policy doesn't allow extension installation, Phase 2 logs a warning and skips multivalue controls until the extension is installed manually or the script is rerun with an account/PAT that can manage extensions.

## Phase 6 HTML summary report

Phase 6 creates a self-contained HTML report in the project output folder. The report includes:

- planned, imported, skipped, unresolved failure, and resolved prior failure counts,
- worker/thread count, retry settings, and tracked elapsed time when available,
- quick links to the Azure DevOps project, Boards, Work Items, process settings, output folder, import files, and latest phase logs,
- latest phase log API status counts and expandable historical log findings,
- failure reconciliation against `ado-id-map.csv` so stale failure files do not incorrectly show the import as failed after a successful rerun.

## More guidance

- [Set up Azure DevOps with the Business Process Catalog July](https://learn.microsoft.com/dynamics365/guidance/business-processes/about-configure-azure-devops-july)
- [What's new in the Business Process Catalog Azure DevOps setup July](https://learn.microsoft.com/dynamics365/guidance/business-processes/about-configure-azure-devops-july-whats-new)
- [Business Process Catalog Azure DevOps setup FAQ - July](https://learn.microsoft.com/dynamics365/guidance/business-processes/about-configure-azure-devops-july-faq)
- [What's new in July](docs/whats-new-july.md)
- [July user guide](docs/user-guide-july.md)
- [July FAQ](docs/faq-july.md)
