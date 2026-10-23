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
    for docx_path in docx_paths:
        pdf_path_for(docx_path, output_dir).unlink(missing_ok=True)

    # Use Microsoft Word where it exists (macOS, Windows), LibreOffice elsewhere
    if sys.platform == "darwin":
        convert_fn = convert_with_word_on_mac
    elif sys.platform == "win32":
        convert_fn = convert_with_word_on_windows
    else:
        convert_fn = convert_with_libreoffice

    errors = convert_fn(docx_paths, output_dir)
    if errors is None:
        return

    for docx_path in docx_paths:
        pdf_path = pdf_path_for(docx_path, output_dir)
        if pdf_path.is_file():
            logger.info("converted %s -> %s", docx_path, pdf_path)
        else:
            reason = errors.get(docx_path) or "unknown error"
            logger.error("failed to convert %s to PDF: %s", docx_path.name, reason)


def pdf_path_for(docx_path: Path, output_dir: Path) -> Path:
    return output_dir / docx_path.with_suffix(".pdf").name


# Converters return None if skipped, else errors for the files that failed
ConversionErrors = dict[Path, str]


# Saves a .docx as PDF with Word. Unlike docx2pdf's script, needs no System Events access
WORD_TO_PDF_SCRIPT = """
function run(argv) {
  const [docxPath, pdfPath] = argv;
  const word = Application("Microsoft Word");
  word.open(Path(docxPath));
  const doc = word.activeDocument;
  doc.saveAs({ fileName: pdfPath, fileFormat: "format PDF" });
  doc.close({ saving: "no" });
}
"""
OSASCRIPT = ["/usr/bin/osascript", "-l", "JavaScript", "-e"]
# Error -1743: macOS didn't allow controlling another app
MAC_PERMISSION_ERROR = "-1743"
MAC_PERMISSION_HINT = (
    " (macOS blocked access to Word: in System Settings > Privacy & Security > "
    "Automation, allow Terminal to control Microsoft Word)"
)


def convert_with_word_on_mac(
    docx_paths: list[Path], output_dir: Path
) -> ConversionErrors:
    errors = {}
    for docx_path in docx_paths:
        pdf_path = pdf_path_for(docx_path, output_dir)
        error = run_word_script(docx_path, pdf_path)
        if error is not None:
            errors[docx_path] = error
    return errors


def run_word_script(docx_path: Path, pdf_path: Path) -> str | None:
    """Returns an error message, or None if the conversion worked."""
    command = [*OSASCRIPT, WORD_TO_PDF_SCRIPT, str(docx_path), str(pdf_path)]
    try:
        result = subprocess.run(command, capture_output=True, text=True, check=False)
    except OSError as e:
        return str(e)

    if result.returncode == 0:
        return None

    exit_message = f"osascript exited with code {result.returncode}"
    message = result.stderr.strip() or exit_message
    if MAC_PERMISSION_ERROR in message:
        message += MAC_PERMISSION_HINT
    return message


def convert_with_word_on_windows(
    docx_paths: list[Path], output_dir: Path
) -> ConversionErrors:
    errors = {}
    for docx_path in docx_paths:
        pdf_path = pdf_path_for(docx_path, output_dir)
        try:
            # keep_active leaves Word open, so it isn't restarted for every file
            docx2pdf.convert(str(docx_path), str(pdf_path), keep_active=True)
        # docx2pdf raises various error types when Word is missing or fails
        except Exception as e:  # noqa: BLE001
            errors[docx_path] = str(e)
    return errors


def convert_with_libreoffice(
    docx_paths: list[Path], output_dir: Path
) -> ConversionErrors | None:
    """Convert all files in one call, since LibreOffice is slow to start."""
    office = shutil.which("soffice") or shutil.which("libreoffice")
    if office is None:
        logger.warning("LibreOffice not found, skipping PDF conversion")
        return None

    paths = [str(path) for path in docx_paths]
    args = [office, "--headless", "--convert-to", "pdf", "--outdir", str(output_dir)]

    command = [*args, *paths]
    result = subprocess.run(command, capture_output=True, text=True, check=False)
    # LibreOffice doesn't say which file failed, so show its output for each
    return {docx_path: result.stderr.strip() for docx_path in docx_paths}
