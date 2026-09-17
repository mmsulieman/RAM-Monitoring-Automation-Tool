# J-MAAP Streamlit Cloud Deployment Checklist

This v1.2 package fixes the `ModuleNotFoundError` seen at `from src.io_utils ...` by removing the fragile `src` package dependency. All local Python modules now sit beside `app.py` at the repository root.

## GitHub repository root must contain

- `app.py`
- `io_utils.py`
- `dq_matching.py`
- `pipeline.py`
- `analytics.py`
- `outputs.py`
- `monitoring_automation_v1.py`
- `requirements.txt`
- `config/`
- `.streamlit/`

Do **not** upload only `app.py`, and do **not** keep the ZIP as a single file in GitHub. Extract the ZIP and upload/commit all contents.

## Streamlit Community Cloud

1. Push all files above to the root of the GitHub repository (for example `j-maap`).
2. In Streamlit Cloud, choose the repository and set **Main file path** to `app.py`.
3. Use Python 3.11 or 3.12 if the deployment settings allow selecting a version.
4. Reboot/redeploy the app after replacing the old files.
5. If an old build remains cached, use **Manage app → Reboot app**.

## Expected startup

The app should load before any MoDa file is uploaded. If it fails before the upload screen, inspect the Streamlit logs; the first unresolved import/package name identifies the deployment problem.
