# Build a single autodoc executable into dist/: `uv run python build.py`
# Double-clicking it fills the templates in the input/ folder next to it and writes
# them to output/. PyInstaller can't cross-compile, so build on each target platform.
import PyInstaller.__main__

PyInstaller.__main__.run(
    [
        "src/autodoc/main.py",
        "--name=autodoc",
        "--onefile",
        "--paths=src",
        "--noconfirm",
        # PyInstaller generates a .spec file; keep it with the other build files
        "--specpath=build",
        # python-docx's default document parts
        "--collect-data=docx",
    ]
)
