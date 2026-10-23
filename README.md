# autodoc

Fills Word templates (`.docx`) with values from an Excel file(s) and exports each
filled-in document as both `.docx` and `.pdf`.

## Input folder

Put the templates and the variables in the `input/` folder next to the program:

```
autodoc/
├── autodoc                        # the program (autodoc.exe on Windows)
├── input/
│   ├── vars.xlsx                  # variables (any name; all .xlsx files are merged)
│   ├── smlouva.docx               # templates
│   ├── email_gratulace.docx
│   └── {{ uchazeč }}_smlouva.docx # variables work in file names too
└── output/                        # created by the program
```

**Variables (`.xlsx`):** first sheet, variable name in column A, value in column B,
no header row. If there are several `.xlsx` files, they're merged; the same variable
in two files is an error. Dates are written as `dd/mm/yyyy`. Format IČO and other
numbers with leading zeros as text, or Excel drops the zeros.

| | A | B |
|---|---|---|
| 1 | uchazeč | Jan Novák |
| 2 | ičo | 08765432 |
| 3 | datum_nástupu | 1. 11. 2026 |

**Templates (`.docx`):** use `{{ variable_name }}` anywhere in the text, with the
exact name from column A. Templates use [Jinja](https://jinja.palletsprojects.com/)
syntax, so conditions like `{% if ... %}` work too (see the
[docxtpl docs](https://docxtpl.readthedocs.io/)).

In **file names**, values are made file-name friendly: `{{ uchazeč }}_smlouva.docx`
becomes `jan_novak_smlouva.docx`.

Variables used in a template but missing from the Excel file are reported as
warnings and render as empty text.

## Output

Documents are written to the `output/` folder next to `input/`, as both `.docx` and
`.pdf`. Files from earlier runs with the same name are overwritten.

`output/logs.txt` contains the log of the latest run, including any warnings and
errors. If something goes wrong, send this file along.

PDFs are exported with **Microsoft Word** on macOS and Windows, and with
**LibreOffice** on Linux. If neither is available, only the `.docx` files are created.

## Using the program

Download the archive for your system from the latest run of the
[Build app](../../actions/workflows/build.yml) workflow and unpack it. You get an
`autodoc/` folder with the program and an example `input/` folder. Then:

1. Replace the files in `input/` with your templates and variables.
2. Double-click the program. A window shows the progress; warnings (e.g. variables
   missing from the Excel file) are shown in red.
3. Press Enter to close the window and open `output/`.

Keep the program and `input/` together in the same folder.

**macOS:** unzip `autodoc-macos.zip`. The program isn't signed by Apple, so the first
launch is blocked: open **System Settings → Privacy & Security** and click **Open
Anyway**. On the first PDF export, allow Terminal to control Microsoft Word.

**Windows:** unzip `autodoc-windows.zip` and run `autodoc.exe`. SmartScreen may warn
about an unknown publisher: click **More info → Run anyway**.

**Linux:** unpack `autodoc-linux.tar.gz` and run `./autodoc` from a terminal.
Requires LibreOffice for PDFs.

## Development

Requires [uv](https://docs.astral.sh/uv/).

```bash
uv run autodoc              # uses ./input, writes to ./output
uv run autodoc some/folder  # any other input folder, writes to some/output
```

### Building the program

```bash
uv run python build.py
```

The result is in `dist/`. PyInstaller only builds for the system it runs on, so the
GitHub Actions workflow in `.github/workflows/build.yml` builds all three and packages
each with the example `input/` folder: run it from the Actions tab, or push a tag
such as `v0.1.0`.
