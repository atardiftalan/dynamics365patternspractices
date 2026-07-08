import os

from ado_setup_config import load_config
from bpc_ado_import.cli import main as importer_main


def _default_catalog_source(excel_file: str) -> str:
    excel_dir = os.path.dirname(os.path.abspath(excel_file))
    if os.path.basename(excel_dir).lower() == "python scripts":
        return os.path.dirname(excel_dir)
    return excel_dir


def _enabled(value: str | None) -> bool:
    return str(value or "").strip().lower() in ("1", "true", "yes")


def main() -> int:
    config = load_config(
        default_log_file="7_BPC_Catalog_Update_Log.txt",
        require_process=True,
        ignore_unknown_args=True,
    )
    source_dir = (
        os.getenv("BPC_ADO_CATALOG_SOURCE_DIR")
        or os.getenv("BPC_ADO_SOURCE_DIR")
        or _default_catalog_source(config.excel_file)
    )
    output_dir = os.getenv("BPC_ADO_IMPORT_OUTPUT") or os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")

    os.environ["BPC_ADO_PAT"] = config.pat
    project_url = f"{config.ado_org_url}/{config.ado_project}/"
    argv = [
        "update",
        "--source",
        source_dir,
        "--template",
        config.excel_file,
        "--project-url",
        project_url,
        "--output",
        output_dir,
        "--pat-env",
        "BPC_ADO_PAT",
    ]
    if not _enabled(os.getenv("BPC_ADO_UPDATE_APPLY")):
        argv.append("--dry-run")
    if os.getenv("BPC_ADO_INCLUDE_DEPRECATED_DELETED", "").strip().lower() in ("1", "true", "yes"):
        argv.append("--include-deprecated-deleted")
    return importer_main(argv)


if __name__ == "__main__":
    raise SystemExit(main())
