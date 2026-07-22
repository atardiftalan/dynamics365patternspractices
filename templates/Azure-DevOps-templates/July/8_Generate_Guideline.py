"""CLI for exporting an existing Azure DevOps configuration to a July guideline."""

from __future__ import annotations

import argparse
import getpass
import os
import re
from pathlib import Path

from ado_guideline_export import AzureDevOpsReadClient, export_guideline, read_devops_configuration


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = SCRIPT_DIR / "ADO template guideline (July).xlsx"


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Read an existing Azure DevOps project/process and build a July-compatible guideline workbook."
    )
    parser.add_argument("--ado-org-url", help="Organization URL or name. Environment: BPC_ADO_ORG_URL.")
    parser.add_argument("--ado-project", help="Project name. Environment: BPC_ADO_PROJECT.")
    parser.add_argument("--process-name", help="Optional process name override. Normally derived from the project.")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="July guideline workbook used as the format baseline.")
    parser.add_argument("--output", help="Output .xlsx path. Default: out/<project>-ADO-guideline-July.xlsx")
    parser.add_argument("--pat", help="PAT value. Prefer BPC_ADO_PAT or the hidden prompt.")
    parser.add_argument("--non-interactive", action="store_true", help="Fail instead of prompting for missing values.")
    return parser


def _value(argument: str | None, *environment_names: str) -> str:
    if argument and argument.strip():
        return argument.strip()
    for name in environment_names:
        value = os.getenv(name, "").strip()
        if value:
            return value
    return ""


def _required(label: str, value: str, *, secret: bool, non_interactive: bool) -> str:
    if value:
        return value
    if non_interactive:
        raise ValueError(f"Missing required configuration: {label}")
    result = getpass.getpass(f"{label} (hidden): ") if secret else input(f"{label}: ")
    result = result.strip()
    if not result:
        raise ValueError(f"{label} is required.")
    return result


def main() -> int:
    args = _parser().parse_args()
    org_url = _required(
        "Azure DevOps organization URL",
        _value(args.ado_org_url, "BPC_ADO_ORG_URL", "ADO_ORG_URL", "AZURE_DEVOPS_ORG_URL"),
        secret=False,
        non_interactive=args.non_interactive,
    )
    project = _required(
        "Azure DevOps project name",
        _value(args.ado_project, "BPC_ADO_PROJECT", "ADO_PROJECT"),
        secret=False,
        non_interactive=args.non_interactive,
    )
    pat = _required(
        "Azure DevOps PAT",
        _value(args.pat, "BPC_ADO_PAT", "AZURE_DEVOPS_EXT_PAT", "ADO_PAT"),
        secret=True,
        non_interactive=args.non_interactive,
    )
    process_name = _value(args.process_name, "BPC_ADO_PROCESS_NAME", "ADO_PROCESS_NAME")
    safe_project = re.sub(r"[^A-Za-z0-9._-]+", "-", project).strip("-") or "project"
    output = Path(args.output) if args.output else SCRIPT_DIR / "out" / f"{safe_project}-ADO-guideline-July.xlsx"

    print(f"Reading Azure DevOps configuration for '{project}'...")
    client = AzureDevOpsReadClient(org_url, pat)
    snapshot = read_devops_configuration(client, project, process_name=process_name)
    print("Building July-compatible guideline workbook...")
    result = export_guideline(snapshot, args.template, output)
    print(f"Guideline created: {result}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nExport canceled.")
        raise SystemExit(130)
    except Exception as exc:
        print(f"\nERROR: {exc}")
        raise SystemExit(1)

