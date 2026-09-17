# J-MAAP
## Jijiga Monitoring, Analytics & Action Platform

J-MAAP is a Streamlit application for monthly monitoring data intake, data-quality review, cross-file reconciliation, raw-data exploration, findings analysis, action tracking, and generation of management outputs.

### Local run

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

### Streamlit Cloud

Upload/commit **all extracted files and folders** to the repository root. The main file path is:

```text
app.py
```

See `DEPLOYMENT_CHECKLIST.md` for the exact required repository structure and troubleshooting steps.

### Important v1.2 fix

Earlier builds imported local modules through `src.*`. If the `src` folder was not committed to the deployed repository, Streamlit raised `ModuleNotFoundError` before the app could load. v1.2 uses flat root-level local modules to make deployment more robust.
