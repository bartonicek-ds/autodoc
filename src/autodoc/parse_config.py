import argparse
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import openpyxl


@dataclass
class Config:
    input_dir: Path
    vars: dict[str, Any]
    output_dir: Path


def app_dir() -> Path:
    """The folder with the executable when built with PyInstaller, else the current one."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd()


def find_vars_files(input_dir: Path) -> list[Path]:
    paths = input_dir.glob("*.xlsx")
    vars_paths = sorted(path for path in paths if not is_ignored(path))

    if not vars_paths:
        raise ValueError(f"no .xlsx vars file found in {input_dir}")

    return vars_paths


def load_config(input_dir: Path) -> Config:
    if not input_dir.is_dir():
        raise ValueError(f"input directory not found: {input_dir}")

    input_dir = input_dir.resolve()
    output_dir = input_dir.parent / "output"

    vars_files = find_vars_files(input_dir)
    merged: dict[str, Any] = {}
    sources: dict[str, Path] = {}

    for path in vars_files:
        for key, value in load_vars(path).items():
            if key in merged:
                source_name = sources[key].name
                raise ValueError(
                    f"variable {key} is defined in both {source_name} and {path.name}"
                )
            merged[key] = value
            sources[key] = path

    return Config(input_dir=input_dir, vars=merged, output_dir=output_dir)


def is_ignored(path: Path) -> bool:
    """Hidden files, e.g. the "._smlouva.docx" companions macOS creates on USB and
    network drives, and Office lock files for open documents, e.g. "~$smlouva.docx"."""
    return path.name.startswith((".", "~$"))


def load_vars(path: Path) -> dict[str, Any]:
    """Read vars from the first sheet: variable names in column A, values in column B."""
    try:
        # data_only returns the computed values of formulas instead of the formulas
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:  # openpyxl raises various errors on bad files
        raise ValueError(f"cannot read Excel vars file {path}: {e}") from e

    data: dict[str, Any] = {}
    rows = workbook.worksheets[0].iter_rows(max_col=2, values_only=True)
    for row_number, row in enumerate(rows, start=1):
        key, value = (*row, None)[:2]
        if key is None:
            if value is not None:
                raise ValueError(f"{path}, row {row_number}: value without a name")
            continue

        key = str(key).strip()
        if key in data:
            raise ValueError(f"{path}, row {row_number}: duplicate variable {key}")
        data[key] = format_excel_value(value)

    workbook.close()
    return data


def format_excel_value(value: Any) -> str:
    if value is None:
        return ""

    if isinstance(value, datetime):
        value = value.date()

    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")

    # Excel stores all numbers as floats, so 42 would otherwise render as "42.0"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value)


def validate_config(config: Config) -> None:
    if config.output_dir.exists() and not config.output_dir.is_dir():
        raise ValueError(f"output path is not a directory: {config.output_dir}")


def parse_config(argv: list[str] | None = None) -> Config:
    parser = argparse.ArgumentParser(
        prog="autodoc",
        description=(
            "Fill .docx templates with variables from an Excel file. "
            "Rendered documents are written to an output folder next to the input folder."
        ),
    )

    default_input_dir = app_dir() / "input"
    parser.add_argument(
        "input_dir",
        type=Path,
        nargs="?",
        default=default_input_dir,
        help=(
            "directory with the .docx templates and .xlsx vars files "
            "(variable names in column A, values in column B; all files are merged); "
            f"default: {default_input_dir}"
        ),
    )

    args = parser.parse_args(argv)

    try:
        config = load_config(args.input_dir)
        validate_config(config)
    except ValueError as e:
        parser.error(str(e))

    return config
