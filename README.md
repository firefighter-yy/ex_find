# Excel Information Search

This repository contains the Windows desktop foundation for a local Excel search tool.
The application is designed for Python 3.10+, PySide6, and Microsoft Excel with `pywin32`.

## Development

The project uses the Miniconda environment named `Ex_Find`. Install Miniconda under
`F:\Miniconda3`, then create the environment from the repository root:

```powershell
# Run the official Miniconda installer and choose this install directory:
# F:\Miniconda3
conda env create -f environment.yml
conda activate Ex_Find
python -m pip install -e . --no-deps
```

If `conda` is not available yet, install Miniconda for Windows from the official
Anaconda repository, select `F:\Miniconda3` as the destination, then open a new
Anaconda Prompt before running the commands above. Do not create a second `venv`
inside this project.

The unified entry point is available as either `excel-search` or `python -m ex_transform`:

```powershell
excel-search diagnose
excel-search start
excel-search test
```

When invoking pytest directly, run it against the project path rather than a drive root:

```powershell
cd E:\ex_transform
F:\Miniconda3\envs\Ex_Find\python.exe -m pytest -q E:\ex_transform
```

`diagnose` reports Windows, Python, and private Excel COM availability. The diagnostic probe
creates its own hidden Excel instance and always closes it. The source workbook is never opened
by the foundation layer except through the read-only `ExcelApplication.open_readonly` method.

Set `EXCEL_SEARCH_TEMP_ROOT` to place per-session temporary indexes under a chosen directory,
or `EXCEL_SEARCH_LOG_LEVEL` to change the application log level. Logs redact workbook paths and
do not contain cell contents or search terms.

## Windows packaging

Run `scripts\build-windows.ps1` from the `Ex_Find` environment. The script runs the test
suite and creates `dist\ExcelInformationSearch.exe` with an embedded Python runtime.
Supported platforms, installation steps, and known limitations are documented in
`docs\release-notes-0.1.0.md`.
