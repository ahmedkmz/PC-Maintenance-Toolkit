# PMT

PMT is a single-file Windows maintenance utility for diagnostics, servicing, cleanup, update checks, verification, and report generation.

It is designed to run as a standalone Python script on Windows 10 and Windows 11, with no required third-party dependencies.

## What it does

PMT runs a structured maintenance workflow with reporting:

1. Baseline system diagnostics
2. Windows servicing checks and repair commands
3. Disk and filesystem health checks
4. Safe cleanup and maintenance tasks
5. Windows Update scan and repair flow
6. Driver assistance checks
7. Optional network remediation
8. Post-repair verification

It also generates session artifacts including logs, raw command output, JSON data, a text summary, and a PDF report.

## Requirements

- Windows 10 or Windows 11
- Python 3.11 or newer recommended
- Administrator rights

Optional:

- Internet access for Windows Update checks and NVIDIA App download assistance
- `reportlab` for enhanced PDF output

If `reportlab` is not installed, PMT uses its built-in PDF fallback.

## Script

The GitHub-ready script in this repository is:

```powershell
pc-mt.py
```



## Quick start

Run from an elevated PowerShell window:

```powershell
python .\pc-mt.py
```

Safe diagnostic-only run:

```powershell
python .\pc-mt.py --report-only --quick-mode --skip-driver-stage
```

Non-interactive run:

```powershell
python .\pc-mt.py --non-interactive
```

Allow reboot scheduling when needed:

```powershell
python .\pc-mt.py --allow-reboot
```

## Command-line flags

- `--non-interactive`  
  Run without prompts and use safe defaults automatically.

- `--allow-reboot`  
  Allow reboot scheduling and automatic resume support when required.

- `--skip-driver-stage`  
  Skip the NVIDIA driver assistance stage.

- `--network-reset`  
  Enable optional network remediation when applicable.

- `--report-only`  
  Collect diagnostics and generate reports without maintenance or repair actions.

- `--deep-scan`  
  Expand diagnostic scope, including deeper event log and verification coverage.

- `--quick-mode`  
  Reduce auxiliary work while still running the core maintenance flow.

## Output

Each run creates a session folder under:

```text
PMT\<session_id>\
```

Typical structure:

```text
PMT\<session_id>\
|- logs\
|- raw\
|- reports\
|- downloads\
|- temp\
`- session_state.json
```

Reports are written to:

```text
PMT\<session_id>\reports\
```

Including:

- `*_report.json`
- `*_summary.txt`
- `*_report.pdf`

## Behavior notes

- PMT checks for administrator rights and will try to relaunch elevated through UAC if needed.
- Safety checks can prevent automated repair actions when the environment is not suitable.
- Some operations can take a long time, especially DISM, SFC, CHKDSK, and Windows Update scans.
- The tool stores raw command output for review in the session folder.

## Recommended usage

For first-time use on an unfamiliar machine:

```powershell
python .\pc-mt.py --report-only --quick-mode --skip-driver-stage
```

Then review the generated report before running a full maintenance session.

