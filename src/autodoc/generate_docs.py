import logging
import re
import shutil
import subprocess
import sys
import unicodedata
from pathlib import Path
from typing import Any

import docx2pdf  # type: ignore[import-untyped]
import jinja2
import jinja2.meta
from docxtpl import DocxTemplate

from autodoc.parse_config import Config, is_ignored

logger = logging.getLogger(__name__)


def generate_docs(config: Config) -> None:
    config.output_dir.mkdir(parents=True, exist_ok=True)

    paths = config.input_dir.glob("*.docx")
    template_paths = sorted(path for path in paths if not is_ignored(path))
    if not template_paths:
        logger.warning("no .docx templates found in %s", config.input_dir)

    output_paths: list[Path] = []
    for template_path in template_paths:
        # One broken template shouldn't stop the others from being generated
        try:
            output_path = render_template(template_path, config)
        except Exception as e:  # noqa: BLE001
            logger.error("failed to render %s: %s", template_path.name, e)
            continue

        if output_path in output_paths:
            message = "%s overwrote an earlier output %s"
            logger.warning(message, template_path.name, output_path.name)
        else:
            output_paths.append(output_path)

    if output_paths:
        convert_to_pdf(output_paths, config.output_dir)


def render_template(template_path: Path, config: Config) -> Path:
    """Fill in a template and save it to the output folder; returns the saved path."""
    template = DocxTemplate(template_path)
    output_name, filename_vars = render_filename(template_path.name, config.vars)

    # Jinja renders undefined variables as empty strings, so warn about them explicitly
    missing_vars = template.get_undeclared_template_variables(context=config.vars)
    missing_vars |= filename_vars - config.vars.keys()
    if missing_vars:
        logger.warning(
            "%s uses variables missing from vars: %s",
            template_path.name,
            ", ".join(sorted(missing_vars)),
        )

    template.render(config.vars)
    output_path = config.output_dir / output_name
    template.save(output_path)
    logger.info("rendered %s -> %s", template_path, output_path)
    return output_path


def render_filename(name: str, vars: dict[str, Any]) -> tuple[str, set[str]]:
    """Fill in variables in a template's filename, e.g. "{{ uchazeč }}_smlouva.docx"
    -> "jan_novak_smlouva.docx". Returns the new name and the variables it uses."""
    # macOS can hand back filenames in decomposed Unicode (NFD), where "č" is "c" plus a
    # separate caron, which wouldn't match variable names like "uchazeč"
    name = unicodedata.normalize("NFC", name)
    stem, suffix = Path(name).stem, Path(name).suffix

    env = jinja2.Environment()
    try:
        used_vars = jinja2.meta.find_undeclared_variables(env.parse(stem))
    except jinja2.TemplateSyntaxError as e:
        logger.warning("cannot fill in variables in filename %s: %s", name, e)
        return name, set()

    filename_vars = {key: slugify(value) for key, value in vars.items()}
    template = env.from_string(stem)
    rendered_stem = template.render(filename_vars)
    return rendered_stem + suffix, used_vars


def slugify(value: Any) -> str:
    """Make a value safe for filenames, e.g. "Jan Novák" -> "jan_novak"."""
    # NFKD splits accented letters into base letter + accent, e.g. "á" -> "a" + "´"
    decomposed = unicodedata.normalize("NFKD", str(value))
    # Dropping non-ASCII characters then removes the accents
    ascii_text = decomposed.encode("ascii", "ignore").decode()
    slug = re.sub(r"[^a-z0-9]+", "_", ascii_text.lower())
    return slug.strip("_")


def convert_to_pdf(docx_paths: list[Path], output_dir: Path) -> None:
    # Remove PDFs from earlier runs, so a failed conversion isn't mistaken for success
    pdf_paths = [output_dir / path.with_suffix(".pdf").name for path in docx_paths]
    for pdf_path in pdf_paths:
        pdf_path.unlink(missing_ok=True)

    # docx2pdf drives Microsoft Word, which only exists on macOS and Windows
    has_word = sys.platform in ("win32", "darwin")
    convert_fn = convert_with_word if has_word else convert_with_libreoffice
    details = convert_fn(docx_paths, output_dir)
    if details is None:
        return

    for docx_path, pdf_path in zip(docx_paths, pdf_paths):
        if pdf_path.is_file():
            logger.info("converted %s -> %s", docx_path, pdf_path)
        else:
            reason = details or "unknown error"
            logger.error("failed to convert %s to PDF: %s", docx_path, reason)


# The converters return None if they skipped the conversion, otherwise details to show
# for any file that didn't convert ("" if there are none)


def convert_with_word(docx_paths: list[Path], output_dir: Path) -> str | None:
    """Convert with Microsoft Word, via docx2pdf on Windows and its script on macOS."""
    errors = []
    for docx_path in docx_paths:
        pdf_path = output_dir / docx_path.with_suffix(".pdf").name
        try:
            if sys.platform == "darwin":
                convert_with_word_on_mac(docx_path, pdf_path)
            else:
                # keep_active leaves Word open, so it isn't restarted for every file
                docx2pdf.convert(str(docx_path), str(pdf_path), keep_active=True)
        # docx2pdf raises various error types when Word is missing or fails
        except Exception as e:  # noqa: BLE001
            errors.append(f"{docx_path.name}: {e}")
    return "; ".join(errors)


def convert_with_word_on_mac(docx_path: Path, pdf_path: Path) -> None:
    """Run docx2pdf's Word script directly, because docx2pdf itself hides its errors
    on macOS (it ignores osascript's error output and exit code)."""
    script = Path(docx2pdf.__file__).parent / "convert.jxa"
    # The last argument is keep_active: leave Word open for the next file
    script_args = [str(script), str(docx_path), str(pdf_path), "true"]
    command = ["/usr/bin/osascript", "-l", "JavaScript", *script_args]

    result = subprocess.run(command, capture_output=True, text=True, check=False)
    if result.returncode != 0:
        output = result.stderr.strip() or result.stdout.strip()
        raise RuntimeError(output or f"osascript exited with code {result.returncode}")


def convert_with_libreoffice(docx_paths: list[Path], output_dir: Path) -> str | None:
    """Convert with LibreOffice, all files in one call since it's slow to start."""
    office = shutil.which("soffice") or shutil.which("libreoffice")
    if office is None:
        logger.warning("LibreOffice not found, skipping PDF conversion")
        return None

    paths = [str(path) for path in docx_paths]
    args = [office, "--headless", "--convert-to", "pdf", "--outdir", str(output_dir)]

    command = [*args, *paths]
    result = subprocess.run(command, capture_output=True, text=True, check=False)
    return result.stderr.strip()
