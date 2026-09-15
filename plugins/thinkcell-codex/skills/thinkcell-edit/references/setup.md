# Setup

Run locally on Windows with desktop PowerPoint and a licensed, connected think-cell installation. Cloud-only sessions cannot reach the local Office runtime. The optional PowerPoint handoff also requires the separately installed, connected PowerPoint plugin with the export/OOXML tools described in its reference.

Use Python 3.10 or later. In Codex, discover its available workspace runtime before assuming the `python` launcher is usable. Otherwise use an installed Python or a project virtual environment. Install these dependencies once into that interpreter:

```powershell
python -m pip install -r "<skill-directory>/scripts/requirements.txt"
python "<skill-directory>/scripts/thinkcell.py" doctor
```

`doctor` reads dependency versions and finds `ppttc.exe`. It does not open PowerPoint, attest a license, or contact a service. Execution checks the connected think-cell add-in during native verification. If discovery fails, pass `--ppttc "C:/path/to/ppttc.exe"` to `update` or set the process environment variable `THINKCELL_PPTTC`. An explicit invalid override is not silently ignored.

PowerShell helpers require Windows PowerShell 5.1, not PowerShell 7. Invoke its absolute path under `%SystemRoot%/System32/WindowsPowerShell/v1.0/powershell.exe`, with `-NoProfile -ExecutionPolicy RemoteSigned`. Respect managed execution policies; do not bypass them or reset credentials. Installing dependencies requires package-index access; chart-processing scripts do not make network requests.

The archive contains the scripts, not Python, PowerPoint, think-cell, proprietary sample decks, or a PowerPoint connector. No OpenAI API key is needed by these scripts. Codex and any Office connector follow the user's account settings and data policies.

If a runtime step fails, preserve its report and working folder. Diagnose the specific error. A timeout may leave an Office operation running; do not restart until its state is known. `doctor` success alone is not a chart-editability certificate.
