# JoeOS Write

**Work in progress.** A local Windows DOCX editor focused on preserving the original document package while editing text. Not a Microsoft Word replacement.

## Run from source

1. Install Python 3.12 or newer on Windows.
2. Run `Setup.cmd` to create a virtual environment and install Python dependencies.
3. Run `Run JoeOS Write.cmd`.

The active entry point is `joeos_app.py`. Older editor prototypes are retained for reference. Renderer assets are included; to regenerate them, install Node.js, run `npm ci`, then `.venv\\Scripts\\python.exe prepare_renderer.py`.

## Design

The desktop shell uses PySide6/Qt WebEngine and a local HTML/JavaScript preview. `preserved_docx.py` retains the original DOCX package and patches supported XML changes rather than rebuilding documents from HTML. Unedited saves are intended to preserve original bytes. Synthetic tests exercise package preservation, text editing and layout helpers.

## Limitations

Complex Word layouts, floating objects, tables, charts and equations may render differently or have limited editing support. Missing local fonts affect preview fidelity. Word-perfect visual equivalence has not been established. Protected or signed documents are read-only; encrypted DOCX, legacy DOC and DOCM are not supported. There is no guaranteed crash recovery: work on copies and save frequently.

## Development

```powershell
.venv\\Scripts\\python.exe -B -m unittest test_preservation test_text_editing test_package_storage test_layout_model
```

Other scripts include desktop/UI tests and local visual QA tools. These may open windows and create local test output. `Build.cmd` builds a portable Windows app using PyInstaller; generated distributions are not part of this source repository.

## Dependencies and assets

Third-party notices are under `THIRD-PARTY-LICENSES/`. The bundled docx-preview renderer is Apache-2.0 licensed; JSZip is dual MIT/GPLv3 licensed. No Microsoft or other proprietary font binaries are included. Optional font installers affect your own machine and are not required for basic startup.

JoeOS branding assets are project artwork. Publication does not grant an additional license to project code or branding; no project-wide license has been selected here. Third-party components retain their own terms.

## Publication hygiene

This is a fresh source snapshot, without local Git history, personal documents, resume-specific tests, audit captures/results, user profiles, local preferences, installed dependencies, executable archives or build outputs. Originals on the development computer were left unchanged. See `VERIFICATION.md` for the publication check scope.
