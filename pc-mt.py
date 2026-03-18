#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PMT
Version: 1.0
By ahmedkmz

Single-file Windows maintenance utility for diagnosis, maintenance, servicing,
update assistance, verification, and reporting.

Command-line flags:
  --non-interactive   Run without prompts; choose safe defaults automatically.
  --allow-reboot      Allow the tool to schedule reboot-required work and offer
                      or trigger reboot with resume support.
  --skip-driver-stage Skip NVIDIA driver assistance stage.
  --network-reset     Enable controlled networking remediation when applicable.
  --report-only       Collect diagnostics and generate reports without making
                      repair or maintenance changes.
  --deep-scan         Expand event log and disk inspection scope.
  --quick-mode        Reduce auxiliary diagnostics and cleanup scope while still
                      executing core health and repair checks.

Internal flags:
  --resume-state PATH Resume a previously saved session state after reboot.
"""

from __future__ import annotations

import argparse
import base64
import ctypes
import datetime as dt
import getpass
import json
import locale
import logging
import os
import platform
import re
import shutil
import socket
import ssl
import subprocess
import sys
import tempfile
import textwrap
import time
import traceback
import urllib.request
import uuid
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import winreg  # type: ignore
except Exception:  # pragma: no cover
    winreg = None


TOOL_NAME = "PMT"
TOOL_VERSION = "1.0"
COPYRIGHT_TEXT = "By ahmedkmz"
REPORT_TITLE = "PC Maintenance Toolkit Report"
AUTHOR_NAME = "ahmedkmz"
APP_DIR_NAME = "PMT"
CONSOLE_WIDTH = 88
DEFAULT_HTTP_TIMEOUT = 15
SYSTEM_DRIVE_ENV = os.environ.get("SystemDrive", "C:")
NVIDIA_APP_PAGE = "https://www.nvidia.com/en-us/software/nvidia-app/"
PUBLIC_IP_ENDPOINTS = [
    "https://api.ipify.org?format=json",
    "https://api64.ipify.org?format=json",
    "https://checkip.amazonaws.com/",
]
CRITICAL_SERVICES = [
    "BITS",
    "CryptSvc",
    "Dnscache",
    "EventLog",
    "LanmanWorkstation",
    "RpcSs",
    "Schedule",
    "W32Time",
    "WinDefend",
    "Winmgmt",
    "wuauserv",
]
UPDATE_SERVICES = ["BITS", "CryptSvc", "TrustedInstaller", "UsoSvc", "WaaSMedicSvc", "wuauserv"]
SFC_SCAN_TIMEOUT_SECONDS = 60 * 60
SFC_VERIFY_TIMEOUT_SECONDS = 30 * 60
SFC_SCAN_STALL_THRESHOLD_SECONDS = 10 * 60
SFC_VERIFY_STALL_THRESHOLD_SECONDS = 8 * 60
SFC_MIN_RUNTIME_BEFORE_STALL_SECONDS = 15 * 60


@dataclass
class SessionPaths:
    base: Path
    logs: Path
    temp: Path
    downloads: Path
    reports: Path
    raw: Path
    state_file: Path
    json_report: Path
    txt_summary: Path
    pdf_report: Path
    log_file: Path


def enable_virtual_terminal() -> bool:
    if not sys.stdout.isatty():
        return False
    if os.name != "nt":
        return True
    try:
        kernel32 = ctypes.windll.kernel32
        invalid_handle = ctypes.c_void_p(-1).value
        for handle_id in (-11, -12):
            handle = kernel32.GetStdHandle(handle_id)
            if handle in (0, invalid_handle):
                continue
            mode = ctypes.c_uint()
            if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
                kernel32.SetConsoleMode(handle, mode.value | 0x0004 | 0x0001 | 0x0002)
        return True
    except Exception:
        return False


def set_console_title(title: str) -> None:
    try:
        ctypes.windll.kernel32.SetConsoleTitleW(title)
    except Exception:
        pass


class Console:
    RESET = "\033[0m"
    BOLD = "\033[1m"
    DIM = "\033[2m"
    COLORS = {
        "primary": "\033[38;2;97;157;215m",
        "secondary": "\033[38;2;71;98;128m",
        "ice": "\033[38;2;186;207;224m",
        "text": "\033[38;2;240;245;249m",
        "muted": "\033[38;2;126;141;156m",
        "accent": "\033[38;2;208;165;86m",
        "success": "\033[38;2;94;166;118m",
        "warning": "\033[38;2;222;183;97m",
        "danger": "\033[38;2;210;96;96m",
        "signal": "\033[38;2;116;186;201m",
    }
    STATUS_COLORS = {
        "INFO": "ice",
        "OK": "success",
        "WARN": "warning",
        "FAIL": "danger",
        "ACTION": "primary",
        "BOOT": "accent",
    }
    PHASE_COLORS = [
        "primary",
        "secondary",
        "ice",
        "accent",
        "primary",
        "secondary",
        "accent",
        "success",
    ]

    def __init__(self, logger: logging.Logger, animated: bool = True) -> None:
        self.logger = logger
        self.animated = animated
        self.use_color = enable_virtual_terminal()
        self.width = max(78, min(120, shutil.get_terminal_size((CONSOLE_WIDTH, 30)).columns))
        self.use_unicode = self._can_render("┏━┃┗◆")
        self.progress_active = False
        self.progress_last_width = 0

    def _can_render(self, text: str) -> bool:
        encoding = getattr(sys.stdout, "encoding", None) or "utf-8"
        try:
            text.encode(encoding)
            return True
        except Exception:
            return False

    def _fmt(self, text: str, *tokens: str) -> str:
        if not self.use_color:
            return text
        codes = []
        for token in tokens:
            if token == "bold":
                codes.append(self.BOLD)
            elif token == "dim":
                codes.append(self.DIM)
            else:
                codes.append(self.COLORS.get(token, ""))
        return "".join(codes) + text + self.RESET

    def _write(self, text: str, delay: float = 0.0) -> None:
        self.progress_done()
        print(text)
        if self.animated and delay > 0:
            time.sleep(delay)

    def _pulse(self, label: str, message: str, color: str, steps: int = 6) -> None:
        if not (self.animated and self.use_color and sys.stdout.isatty()):
            self._line(label, message)
            return
        frames = ["-", "\\", "|", "/"]
        prefix = self._fmt(f"[{label:<6}]", color, "bold")
        for index in range(steps):
            frame = frames[index % len(frames)]
            sys.stdout.write(f"\r{prefix} {self._fmt(frame, color)} {self._fmt(message, 'text')}")
            sys.stdout.flush()
            time.sleep(0.05)
        sys.stdout.write("\r" + " " * max(0, self.width - 1) + "\r")
        sys.stdout.flush()
        self._line(label, message)

    def _phase_color(self, title: str) -> str:
        match = re.search(r"Stage\s+(\d+)", title)
        if match:
            index = max(1, int(match.group(1))) - 1
            return self.PHASE_COLORS[index % len(self.PHASE_COLORS)]
        return "secondary"

    def _visible_length(self, text: str) -> int:
        return len(re.sub(r"\x1b\[[0-9;]*m", "", text))

    def _clip_plain(self, text: str, width: int) -> str:
        if width <= 0:
            return ""
        plain = str(text)
        if len(plain) <= width:
            return plain
        if width <= 3:
            return plain[:width]
        return plain[: width - 3] + "..."

    def _box_line(self, left: str, fill: str, right: str, color: str) -> str:
        return self._fmt(left + fill * (self.width - 2) + right, color, "bold")

    def _box_row(
        self,
        content: str,
        edge: str,
        edge_color: str,
        text_color: str = "text",
    ) -> str:
        clean = content[: self.width - 2].ljust(self.width - 2)
        return (
            self._fmt(edge, edge_color, "bold")
            + self._fmt(clean, text_color)
            + self._fmt(edge, edge_color, "bold")
        )

    def progress(self, label: str, elapsed_seconds: int, timeout_seconds: int = 0, note: str = "") -> None:
        if not sys.stdout.isatty():
            return
        spinner_frames = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧"] if self.use_unicode else ["-", "\\", "|", "/"]
        spinner = spinner_frames[(elapsed_seconds // 5) % len(spinner_frames)]
        runtime = minutes_to_human(elapsed_seconds / 60.0)
        watch = f"elapsed {runtime} / cap {minutes_to_human(timeout_seconds / 60.0)}" if timeout_seconds else f"elapsed {runtime}"
        prefix = self._fmt("[RUN   ]", "primary", "bold")
        spin = self._fmt(spinner, "accent", "bold")
        activity = self._fmt("active", "signal", "bold")
        label_room = 24
        label_text = self._fmt(self._clip_plain(label, label_room).ljust(label_room), "text", "bold")
        timing = self._fmt(watch, "muted")
        static_len = (
            1
            + self._visible_length(prefix)
            + 1
            + self._visible_length(label_text)
            + 1
            + self._visible_length(spin)
            + 1
            + self._visible_length(activity)
            + 1
            + self._visible_length(timing)
        )
        note_room = max(0, self.width - static_len - 2)
        note_text = ""
        if note_room > 6:
            clipped_note = self._clip_plain(note, note_room)
            if clipped_note:
                note_text = " " + self._fmt(clipped_note, "muted")
        line = f"{prefix} {label_text} {spin} {activity} {timing}{note_text}"
        plain_width = self._visible_length(line)
        padding = max(0, self.progress_last_width - plain_width)
        sys.stdout.write("\r" + line + (" " * padding))
        sys.stdout.flush()
        self.progress_active = True
        self.progress_last_width = max(self.progress_last_width, plain_width)

    def progress_done(self) -> None:
        if not self.progress_active or not sys.stdout.isatty():
            return
        sys.stdout.write("\r" + (" " * self.progress_last_width) + "\r")
        sys.stdout.flush()
        self.progress_active = False
        self.progress_last_width = 0

    def _line(self, label: str, message: str) -> None:
        self.progress_done()
        color = self.STATUS_COLORS.get(label, "text")
        timestamp = dt.datetime.now().strftime("%H:%M:%S")
        prefix = self._fmt(f"[{label:<6}]", color, "bold")
        stamp = self._fmt(timestamp, "dim", "muted")
        body = self._fmt(message, "text")
        print(f"{stamp} {prefix} {body}")
        self.logger.info("%s %s", label, message)

    def section(self, title: str) -> None:
        self.progress_done()
        color = self._phase_color(title)
        fill = "━" if self.use_unicode else "="
        badge = f"{'◆ ' if self.use_unicode else '> '}{title}"
        side = max(0, (self.width - len(badge) - 2) // 2)
        line = (fill * side) + " " + badge + " "
        if len(line) < self.width:
            line += fill * (self.width - len(line))
        border = self._fmt(line[: self.width], color, "bold")
        print()
        print(border)
        self.logger.info("SECTION %s", title)

    def banner(self, session_id: str) -> None:
        set_console_title(f"{TOOL_NAME} | Session {session_id}")
        headline = TOOL_NAME.upper()
        strapline = "WINDOWS MAINTENANCE AND DIAGNOSTICS"
        subtitle = "Single-file maintenance utility"
        box_color = "primary"
        info_color = "muted"
        left_top, right_top, left_bottom, right_bottom, side, fill = (
            ("┏", "┓", "┗", "┛", "┃", "━") if self.use_unicode else ("+", "+", "+", "+", "|", "=")
        )
        divider_fill = "─" if self.use_unicode else "-"
        border = self._box_line(left_top, fill, right_top, box_color)
        footer = self._box_line(left_bottom, fill, right_bottom, box_color)
        now_text = dt.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")
        header_lines = [
            border,
            self._box_row("", side, box_color),
            self._box_row(strapline.center(self.width - 2), side, box_color, "secondary"),
            self._box_row("", side, box_color),
            self._box_row(headline.center(self.width - 2), side, box_color, "text"),
            self._box_row(subtitle.center(self.width - 2), side, box_color, "accent"),
            self._box_row(divider_fill * (self.width - 2), side, box_color, "secondary"),
            self._box_row(f"Version    : {TOOL_VERSION}", side, box_color, info_color),
            self._box_row(COPYRIGHT_TEXT, side, box_color, "accent"),
            self._box_row("", side, box_color),
            self._box_row(f"Session ID : {session_id}", side, box_color, info_color),
            self._box_row(f"Timestamp  : {now_text}", side, box_color, info_color),
            footer,
        ]
        for line in header_lines:
            self._write(line, delay=0.025 if self.animated else 0.0)
        self._pulse("BOOT", "Console ready", "accent")
        self._pulse("BOOT", "Checks initialized", "primary")
        self._pulse("BOOT", "Starting session", "signal")
        self.logger.info("Banner shown for session %s", session_id)

    def info(self, message: str) -> None:
        self._line("INFO", message)

    def ok(self, message: str) -> None:
        self._line("OK", message)

    def warn(self, message: str) -> None:
        self._line("WARN", message)

    def fail(self, message: str) -> None:
        self._line("FAIL", message)

    def action(self, message: str) -> None:
        self._line("ACTION", message)


class ToolkitContext:
    def __init__(self, args: argparse.Namespace, session_paths: SessionPaths, logger: logging.Logger) -> None:
        self.args = args
        self.paths = session_paths
        self.logger = logger
        self.console = Console(logger, animated=not args.non_interactive)
        self.session_id = session_paths.base.name
        self.session_started = dt.datetime.now().astimezone()
        self.raw_counter = 0
        self.data: Dict[str, Any] = {
            "tool": {
                "name": TOOL_NAME,
                "version": TOOL_VERSION,
                "copyright": COPYRIGHT_TEXT,
            },
            "session": {
                "id": self.session_id,
                "started_at": iso_now(),
                "cwd": str(Path.cwd()),
                "command_line": " ".join(sys.argv),
                "flags": vars(args),
            },
            "machine": {},
            "prechecks": {},
            "diagnostics": {},
            "actions": [],
            "stages": [],
            "findings": [],
            "recommendations": [],
            "unresolved_issues": [],
            "hardware_suspicions": [],
            "summary": {},
            "artifacts": {},
            "resume": {},
        }
        self.major_issues: List[str] = []
        self.warnings: List[str] = []
        self.unresolved: List[str] = []
        self.recommendations: List[str] = []
        self.hardware_suspicions: List[str] = []
        self.reboot_required = False
        self.detected_network_issue = False
        self.resume_task_name: Optional[str] = None
        self.resume_registered = False
        self.resume_loaded = False

    def add_finding(self, severity: str, title: str, details: str, category: str = "general") -> None:
        finding = {
            "severity": severity,
            "title": title,
            "details": details,
            "category": category,
            "timestamp": iso_now(),
        }
        self.data["findings"].append(finding)
        if severity.lower() == "major":
            self.major_issues.append(title)
        elif severity.lower() == "warning":
            self.warnings.append(title)
        self.logger.info("Finding added [%s] %s - %s", severity, title, details)

    def add_recommendation(self, text: str) -> None:
        if text and text not in self.recommendations:
            self.recommendations.append(text)
            self.data["recommendations"] = list(self.recommendations)
            self.logger.info("Recommendation added: %s", text)

    def add_unresolved(self, text: str) -> None:
        if text and text not in self.unresolved:
            self.unresolved.append(text)
            self.data["unresolved_issues"] = list(self.unresolved)
            self.logger.info("Unresolved issue: %s", text)

    def add_hardware_suspicion(self, text: str) -> None:
        if text and text not in self.hardware_suspicions:
            self.hardware_suspicions.append(text)
            self.data["hardware_suspicions"] = list(self.hardware_suspicions)
            self.logger.info("Hardware suspicion: %s", text)

    def record_action(
        self,
        stage: str,
        action: str,
        status: str,
        details: str = "",
        command: Optional[str] = None,
        exit_code: Optional[int] = None,
        duration_seconds: Optional[float] = None,
        raw_stdout: Optional[str] = None,
        raw_stderr: Optional[str] = None,
        metadata: Optional[Dict[str, Any]] = None,
    ) -> None:
        item = {
            "stage": stage,
            "action": action,
            "status": status,
            "details": details,
            "command": command,
            "exit_code": exit_code,
            "duration_seconds": round(duration_seconds or 0.0, 2),
            "stdout_path": raw_stdout,
            "stderr_path": raw_stderr,
            "metadata": metadata or {},
            "timestamp": iso_now(),
        }
        self.data["actions"].append(item)

    def record_stage(
        self,
        name: str,
        status: str,
        started_at: dt.datetime,
        ended_at: dt.datetime,
        summary: str,
        notes: Optional[List[str]] = None,
        metrics: Optional[Dict[str, Any]] = None,
    ) -> None:
        duration = (ended_at - started_at).total_seconds()
        stage_record = {
            "name": name,
            "status": status,
            "started_at": started_at.isoformat(),
            "ended_at": ended_at.isoformat(),
            "duration_seconds": round(duration, 2),
            "summary": summary,
            "notes": notes or [],
            "metrics": metrics or {},
        }
        self.data["stages"].append(stage_record)


def iso_now() -> str:
    return dt.datetime.now().astimezone().isoformat()


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return cleaned[:80] or "capture"


def bytes_to_gb(value: Any) -> float:
    try:
        return round(int(value) / (1024 ** 3), 2)
    except Exception:
        return 0.0


def format_bytes(value: Any) -> str:
    try:
        size = float(value)
    except Exception:
        return "0 B"
    units = ["B", "KB", "MB", "GB", "TB", "PB"]
    index = 0
    while size >= 1024 and index < len(units) - 1:
        size /= 1024.0
        index += 1
    return f"{size:.2f} {units[index]}"


def minutes_to_human(total_minutes: float) -> str:
    minutes = int(max(total_minutes, 0))
    days, rem = divmod(minutes, 1440)
    hours, mins = divmod(rem, 60)
    parts = []
    if days:
        parts.append(f"{days}d")
    if hours:
        parts.append(f"{hours}h")
    parts.append(f"{mins}m")
    return " ".join(parts)


def run_timestamp() -> str:
    return dt.datetime.now().astimezone().strftime("%Y%m%d_%H%M%S")


def setup_logging(session_base: Path) -> Tuple[logging.Logger, Path]:
    logs_dir = session_base / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    log_file = logs_dir / f"pmt_{run_timestamp()}.log"
    logger = logging.getLogger(f"pmt_{session_base.name}")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    formatter = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s")
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(formatter)
    logger.addHandler(fh)
    return logger, log_file


def ensure_windows() -> None:
    if os.name != "nt":
        print("This script is intended to run on Windows 10 or Windows 11.")
        sys.exit(1)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=TOOL_NAME)
    parser.add_argument("--non-interactive", action="store_true")
    parser.add_argument("--allow-reboot", action="store_true")
    parser.add_argument("--skip-driver-stage", action="store_true")
    parser.add_argument("--network-reset", action="store_true")
    parser.add_argument("--report-only", action="store_true")
    parser.add_argument("--deep-scan", action="store_true")
    parser.add_argument("--quick-mode", action="store_true")
    parser.add_argument("--resume-state", help=argparse.SUPPRESS)
    return parser


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin() -> None:
    script = str(Path(sys.argv[0]).resolve())
    params = subprocess.list2cmdline([script] + sys.argv[1:])
    rc = ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
    if rc <= 32:
        raise RuntimeError("UAC elevation was denied or failed.")


def determine_paths(session_id: str, existing_base: Optional[Path] = None) -> SessionPaths:
    base = existing_base or (Path(__file__).resolve().parent / APP_DIR_NAME / session_id)
    logs = base / "logs"
    temp = base / "temp"
    downloads = base / "downloads"
    reports = base / "reports"
    raw = base / "raw"
    for folder in (logs, temp, downloads, reports, raw):
        folder.mkdir(parents=True, exist_ok=True)
    return SessionPaths(
        base=base,
        logs=logs,
        temp=temp,
        downloads=downloads,
        reports=reports,
        raw=raw,
        state_file=base / "session_state.json",
        json_report=reports / f"{session_id}_report.json",
        txt_summary=reports / f"{session_id}_summary.txt",
        pdf_report=reports / f"{session_id}_report.pdf",
        log_file=logs / f"pmt_{run_timestamp()}.log",
    )


def save_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def load_json(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def load_resume_state(path: Path) -> Dict[str, Any]:
    return load_json(path)


def save_resume_state(ctx: ToolkitContext, resume_from: str) -> None:
    manual_resume_command = subprocess.list2cmdline(
        [sys.executable, str(Path(__file__).resolve()), "--resume-state", str(ctx.paths.state_file), "--non-interactive"]
    )
    state = {
        "tool": TOOL_NAME,
        "version": TOOL_VERSION,
        "session_id": ctx.session_id,
        "base_dir": str(ctx.paths.base),
        "state_file": str(ctx.paths.state_file),
        "manual_resume_command": manual_resume_command,
        "resume_from": resume_from,
        "saved_at": iso_now(),
        "data": ctx.data,
        "reboot_required": ctx.reboot_required,
        "resume_task_name": ctx.resume_task_name,
    }
    save_json(ctx.paths.state_file, state)
    ctx.data["resume"] = {
        "state_file": str(ctx.paths.state_file),
        "manual_resume_command": manual_resume_command,
        "saved_at": state["saved_at"],
        "resume_from": resume_from,
        "task_name": ctx.resume_task_name,
    }
    ctx.logger.info("Resume state saved at %s", ctx.paths.state_file)


def delete_if_exists(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except Exception:
        pass


def create_resume_task(ctx: ToolkitContext) -> Tuple[bool, str]:
    task_name = f"{TOOL_NAME.replace(' ', '_')}_{ctx.session_id}"
    script_path = str(Path(__file__).resolve())
    command = subprocess.list2cmdline(
        [sys.executable, script_path, "--resume-state", str(ctx.paths.state_file), "--non-interactive"]
    )
    create_cmd = [
        "schtasks.exe",
        "/Create",
        "/TN",
        task_name,
        "/SC",
        "ONLOGON",
        "/RL",
        "HIGHEST",
        "/RU",
        "SYSTEM",
        "/TR",
        command,
        "/F",
    ]
    result = run_command(ctx, "Resume", "Create scheduled resume task", create_cmd, timeout=60)
    if result["ok"]:
        ctx.resume_task_name = task_name
        ctx.resume_registered = True
        return True, f"Scheduled task created: {task_name}"
    return False, "Unable to create scheduled resume task"


def remove_resume_task(ctx: ToolkitContext) -> None:
    if not ctx.resume_task_name:
        return
    run_command(
        ctx,
        "Resume",
        "Remove scheduled resume task",
        ["schtasks.exe", "/Delete", "/TN", ctx.resume_task_name, "/F"],
        timeout=45,
        acceptable_exit_codes=(0, 1),
    )


def reg_key_exists(root: Any, path: str, value_name: Optional[str] = None) -> bool:
    if winreg is None:
        return False
    try:
        with winreg.OpenKey(root, path) as key:
            if value_name:
                winreg.QueryValueEx(key, value_name)
            return True
    except Exception:
        return False


def reg_value(root: Any, path: str, value_name: str) -> Optional[Any]:
    if winreg is None:
        return None
    try:
        with winreg.OpenKey(root, path) as key:
            value, _ = winreg.QueryValueEx(key, value_name)
            return value
    except Exception:
        return None


def pending_reboot_reasons() -> List[str]:
    reasons: List[str] = []
    checks = [
        (
            winreg.HKEY_LOCAL_MACHINE if winreg else None,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
            None,
            "Component Based Servicing pending reboot",
        ),
        (
            winreg.HKEY_LOCAL_MACHINE if winreg else None,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired",
            None,
            "Windows Update pending reboot",
        ),
        (
            winreg.HKEY_LOCAL_MACHINE if winreg else None,
            r"SYSTEM\CurrentControlSet\Control\Session Manager",
            "PendingFileRenameOperations",
            "Pending file rename operations",
        ),
        (
            winreg.HKEY_LOCAL_MACHINE if winreg else None,
            r"SOFTWARE\Microsoft\Updates",
            "UpdateExeVolatile",
            "UpdateExeVolatile indicates pending reboot",
        ),
    ]
    if winreg is None:
        return reasons
    for root, path, value_name, text in checks:
        if root is None:
            continue
        if value_name is None:
            if reg_key_exists(root, path):
                reasons.append(text)
        else:
            value = reg_value(root, path, value_name)
            if value not in (None, "", 0):
                reasons.append(text)
    return reasons


def is_safe_mode() -> bool:
    if os.environ.get("SAFEBOOT_OPTION"):
        return True
    if winreg is None:
        return False
    value = reg_value(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\SafeBoot\Option", "OptionValue")
    return bool(value)


def internet_available() -> Tuple[bool, str]:
    tests = [
        ("Socket 1.1.1.1:53", lambda: socket.create_connection(("1.1.1.1", 53), timeout=4)),
        ("HTTPS Microsoft connectivity test", lambda: urllib.request.urlopen("https://www.msftconnecttest.com/connecttest.txt", timeout=8)),
    ]
    for label, func in tests:
        try:
            handle = func()
            handle.close()
            return True, label
        except Exception:
            continue
    return False, "No outbound connectivity test succeeded"


def get_public_ip() -> Optional[str]:
    for endpoint in PUBLIC_IP_ENDPOINTS:
        try:
            req = urllib.request.Request(endpoint, headers={"User-Agent": f"{TOOL_NAME}/{TOOL_VERSION}"})
            with urllib.request.urlopen(req, timeout=DEFAULT_HTTP_TIMEOUT, context=ssl.create_default_context()) as response:
                payload = response.read().decode("utf-8", errors="replace").strip()
                if payload.startswith("{"):
                    ip = json.loads(payload).get("ip")
                else:
                    ip = payload.strip()
                if ip and len(ip) <= 64:
                    return ip
        except Exception:
            continue
    return None


def detect_local_ips() -> Dict[str, Any]:
    ips: List[str] = []
    try:
        hostname = socket.gethostname()
        for result in socket.getaddrinfo(hostname, None):
            address = result[4][0]
            if address and ":" not in address and not address.startswith("127.") and address not in ips:
                ips.append(address)
    except Exception:
        pass
    lan_ip = ips[0] if ips else None
    return {"lan_ip": lan_ip, "all_ipv4": ips}


def encode_powershell(script: str) -> str:
    return base64.b64encode(script.encode("utf-16le")).decode("ascii")


def powershell_json_wrapper(inner: str, depth: int = 8) -> str:
    return textwrap.dedent(
        f"""
        $ProgressPreference = 'SilentlyContinue'
        $ErrorActionPreference = 'Stop'
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $result = & {{
        {textwrap.indent(inner.strip(), '    ')}
        }}
        $result | ConvertTo-Json -Depth {depth} -Compress
        """
    ).strip()


def run_powershell_json_quiet(script: str, timeout: int = 30, depth: int = 6) -> Optional[Any]:
    wrapped = powershell_json_wrapper(script, depth=depth)
    cmd = [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-EncodedCommand",
        encode_powershell(wrapped),
    ]
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
    except Exception:
        return None
    if result.returncode != 0 or not result.stdout.strip():
        return None
    try:
        return json.loads(result.stdout.strip())
    except Exception:
        return None


def run_command(
    ctx: ToolkitContext,
    stage: str,
    label: str,
    command: Sequence[str] | str,
    timeout: int = 1800,
    acceptable_exit_codes: Tuple[int, ...] = (0,),
    shell: bool = False,
    cwd: Optional[Path] = None,
    heartbeat_interval: int = 15,
    progress_probe: Optional[Any] = None,
    progress_probe_interval: int = 60,
) -> Dict[str, Any]:
    ctx.raw_counter += 1
    base_name = f"{ctx.raw_counter:03d}_{sanitize_filename(label)}"
    stdout_path = ctx.paths.raw / f"{base_name}_stdout.txt"
    stderr_path = ctx.paths.raw / f"{base_name}_stderr.txt"
    start = dt.datetime.now().astimezone()
    start_monotonic = time.monotonic()
    cmd_text = command if isinstance(command, str) else subprocess.list2cmdline(list(command))
    ctx.console.action(f"{stage}: {label}")
    stdout = ""
    stderr = ""
    stdout_bytes = b""
    stderr_bytes = b""
    exit_code = None
    timed_out = False
    stalled = False
    error_details = ""
    try:
        proc = subprocess.Popen(
            command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=shell,
            cwd=str(cwd) if cwd else None,
        )
        heartbeat = 0
        last_probe = 0
        last_probe_note = ""
        last_progress_update = 0
        while True:
            try:
                out, err = proc.communicate(timeout=5)
                stdout_bytes += out or b""
                stderr_bytes += err or b""
                break
            except subprocess.TimeoutExpired:
                elapsed = int(time.monotonic() - start_monotonic)
                if elapsed > timeout:
                    proc.kill()
                    out, err = proc.communicate()
                    stdout_bytes += out or b""
                    stderr_bytes += err or b""
                    timed_out = True
                    break
                probe_note = ""
                if progress_probe and elapsed - last_probe >= progress_probe_interval:
                    try:
                        probe_result = progress_probe(elapsed)
                    except Exception:
                        probe_result = None
                    last_probe = elapsed
                    if isinstance(probe_result, dict):
                        probe_note = str(probe_result.get("note") or "").strip()
                        last_probe_note = probe_note or last_probe_note
                        if probe_result.get("stalled"):
                            stalled = True
                            error_details += "\nCommand terminated after live probe marked it as stalled."
                            if probe_note:
                                error_details += f"\nLast probe note: {probe_note}"
                            proc.kill()
                            out, err = proc.communicate()
                            stdout_bytes += out or b""
                            stderr_bytes += err or b""
                            timed_out = True
                            break
                if elapsed - last_progress_update >= heartbeat_interval:
                    last_progress_update = elapsed
                    note = probe_note or last_probe_note
                    if sys.stdout.isatty():
                        ctx.console.progress(
                            label=label,
                            elapsed_seconds=elapsed,
                            timeout_seconds=timeout,
                            note=note,
                        )
                    else:
                        note_suffix = f" | {note}" if note else ""
                        ctx.logger.info("ACTION %s still running (%ss)%s", label, elapsed, note_suffix)
        exit_code = proc.returncode
    except FileNotFoundError:
        error_details = f"Command not found: {command}"
        exit_code = 9009
    except Exception:
        error_details = traceback.format_exc()
        exit_code = 1
    stdout = decode_command_output(stdout_bytes)
    stderr = decode_command_output(stderr_bytes)
    if error_details:
        stderr = f"{stderr.rstrip()}\n{error_details}".strip() if stderr else error_details
    duration = time.monotonic() - start_monotonic
    ctx.console.progress_done()
    write_text(stdout_path, stdout)
    write_text(stderr_path, stderr)
    ok = (exit_code in acceptable_exit_codes) and not timed_out
    status = "OK" if ok else ("WARN" if timed_out else "FAIL")
    details = f"Exit code {exit_code}, duration {duration:.1f}s"
    if timed_out:
        if stalled:
            details += ", command stopped after stall detection"
            ctx.console.warn(f"{label} was stopped after appearing stalled for too long")
        else:
            details += ", command timed out"
            ctx.console.warn(f"{label} timed out after {duration:.1f}s")
    elif ok:
        ctx.console.ok(f"{label} completed in {duration:.1f}s")
    else:
        ctx.console.warn(f"{label} completed with exit code {exit_code}")
    ctx.record_action(
        stage=stage,
        action=label,
        status=status,
        details=details,
        command=cmd_text,
        exit_code=exit_code,
        duration_seconds=duration,
        raw_stdout=str(stdout_path),
        raw_stderr=str(stderr_path),
        metadata={
            "started_at": start.isoformat(),
            "ended_at": dt.datetime.now().astimezone().isoformat(),
            "timed_out": timed_out,
            "stalled": stalled,
        },
    )
    return {
        "ok": ok,
        "stdout": stdout,
        "stderr": stderr,
        "exit_code": exit_code,
        "duration_seconds": duration,
        "stdout_path": str(stdout_path),
        "stderr_path": str(stderr_path),
        "command": cmd_text,
        "timed_out": timed_out,
        "stalled": stalled,
    }


def run_powershell_json(
    ctx: ToolkitContext,
    stage: str,
    label: str,
    script: str,
    timeout: int = 300,
    depth: int = 8,
) -> Tuple[Optional[Any], Dict[str, Any]]:
    wrapped = powershell_json_wrapper(script, depth=depth)
    cmd = [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-EncodedCommand",
        encode_powershell(wrapped),
    ]
    result = run_command(ctx, stage, label, cmd, timeout=timeout)
    if not result["stdout"].strip():
        return None, result
    try:
        return json.loads(result["stdout"].strip()), result
    except Exception:
        ctx.console.warn(f"JSON parse failed for {label}; raw output kept for review")
        return None, result


def parse_update_result_code(value: Any) -> str:
    mapping = {
        0: "Not started",
        1: "In progress",
        2: "Succeeded",
        3: "Succeeded with errors",
        4: "Failed",
        5: "Aborted",
    }
    try:
        return mapping.get(int(value), f"Unknown ({value})")
    except Exception:
        return "Unknown"


def parse_hresults(text: str) -> List[str]:
    return sorted(set(re.findall(r"0x[0-9A-Fa-f]{6,8}", text or "")))


def detect_command_output_encoding(payload: bytes) -> Optional[str]:
    if not payload:
        return None
    if payload.startswith(b"\xff\xfe") or payload.startswith(b"\xfe\xff"):
        return "utf-16"
    sample = payload[: min(len(payload), 512)]
    zero_count = sample.count(0)
    if zero_count < max(2, len(sample) // 8):
        return None
    even_bytes = sample[::2]
    odd_bytes = sample[1::2]
    even_ratio = even_bytes.count(0) / max(1, len(even_bytes))
    odd_ratio = odd_bytes.count(0) / max(1, len(odd_bytes))
    if odd_ratio >= 0.35 and even_ratio <= 0.10:
        return "utf-16le"
    if even_ratio >= 0.35 and odd_ratio <= 0.10:
        return "utf-16be"
    return None


def decode_command_output(payload: bytes) -> str:
    if not payload:
        return ""
    hinted_encoding = detect_command_output_encoding(payload)
    if hinted_encoding:
        try:
            return payload.decode(hinted_encoding, errors="replace")
        except Exception:
            pass
    candidates = [
        "utf-8-sig",
        "utf-8",
        locale.getpreferredencoding(False) or "cp1252",
        "cp1252",
        "cp850",
        "cp437",
    ]
    for encoding in candidates:
        try:
            return payload.decode(encoding)
        except (LookupError, UnicodeDecodeError):
            continue
    return payload.decode("utf-8", errors="replace")


def flatten_to_list(value: Any) -> List[Any]:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


def is_portable_machine(machine: Dict[str, Any]) -> bool:
    computer = machine.get("computer", {})
    enclosure_items = flatten_to_list(machine.get("enclosure"))
    portable_system_types = {2, 8}
    portable_chassis_types = {8, 9, 10, 11, 12, 14, 30, 31, 32}
    for key in ("PCSystemTypeEx", "PCSystemType"):
        try:
            if int(computer.get(key, 0) or 0) in portable_system_types:
                return True
        except Exception:
            continue
    for item in enclosure_items:
        if not isinstance(item, dict):
            continue
        for chassis_type in flatten_to_list(item.get("ChassisTypes")):
            try:
                if int(chassis_type) in portable_chassis_types:
                    return True
            except Exception:
                continue
    return False


def active_safety_blockers(ctx: ToolkitContext) -> List[str]:
    blockers = flatten_to_list(ctx.data.get("prechecks", {}).get("major_blockers"))
    return [str(item).strip() for item in blockers if str(item).strip()]


def repairs_allowed(ctx: ToolkitContext) -> bool:
    return (not ctx.args.report_only) and not active_safety_blockers(ctx)


def stage_skip_reason(ctx: ToolkitContext, report_only_message: str, blocked_message: str) -> Optional[str]:
    if ctx.args.report_only:
        return report_only_message
    blockers = active_safety_blockers(ctx)
    if blockers:
        return f"{blocked_message}: {', '.join(blockers)}."
    return None


def should_run_post_sfc_verify(ctx: ToolkitContext) -> Tuple[bool, str]:
    if ctx.args.quick_mode:
        return False, "quick mode skips a second full SFC verification pass"
    servicing = ctx.data.get("diagnostics", {}).get("servicing", {})
    sfc_result = servicing.get("SFC Scannow", {})
    if not sfc_result:
        return False, "the primary SFC scan did not run"
    if sfc_result.get("timed_out"):
        return False, "the primary SFC scan timed out or was stopped"
    summary = str(sfc_result.get("summary", "")).lower()
    exit_code = sfc_result.get("exit_code")
    if "did not complete cleanly" in summary:
        return False, "the primary SFC scan did not complete cleanly"
    if ctx.args.deep_scan:
        return True, "deep scan mode requested extended post-repair verification"
    if "repaired corrupt files" in summary or "could not repair all files" in summary:
        return True, "the primary SFC scan reported repair activity that warrants verification"
    if exit_code not in (0, 1):
        return False, f"the primary SFC exit code was {exit_code}"
    return False, "the primary SFC scan did not report repair activity that requires another full pass"


def parse_uptime_minutes(last_boot: str) -> int:
    if not last_boot:
        return 0
    try:
        cleaned = re.sub(r"[+-]\d{2}:\d{2}$", "", last_boot)
        parsed = dt.datetime.fromisoformat(cleaned)
        delta = dt.datetime.now() - parsed
        return int(delta.total_seconds() // 60)
    except Exception:
        return 0


def command_exists(name: str) -> bool:
    return shutil.which(name) is not None


def summarize_machine_overview(ctx: ToolkitContext) -> None:
    machine = ctx.data.get("machine", {})
    os_info = machine.get("os", {})
    comp = machine.get("computer", {})
    local_ip = machine.get("networking", {}).get("lan_ip") or "Unavailable"
    public_ip = machine.get("networking", {}).get("public_ip") or "Unavailable"
    uptime_minutes = machine.get("uptime_minutes") or 0
    ctx.console.info(
        f"{comp.get('Hostname', platform.node())} | {os_info.get('Caption', 'Windows')} "
        f"build {os_info.get('BuildNumber', '?')} | User: {ctx.data['session'].get('current_user', getpass.getuser())}"
    )
    ctx.console.info(
        f"Admin: {'Yes' if ctx.data['session'].get('is_admin') else 'No'} | "
        f"Uptime: {minutes_to_human(uptime_minutes)} | Local IP: {local_ip} | Public IP: {public_ip}"
    )
    ctx.console.info(f"Artifacts: {ctx.paths.base}")


def enrich_machine_context(ctx: ToolkitContext, baseline: Dict[str, Any]) -> None:
    local_ips = detect_local_ips()
    online, online_detail = internet_available()
    public_ip = get_public_ip() if online else None
    baseline["networking"] = {
        "internet_available": online,
        "internet_detail": online_detail,
        "lan_ip": local_ips.get("lan_ip"),
        "local_ipv4": local_ips.get("all_ipv4", []),
        "public_ip": public_ip,
    }
    last_boot = baseline.get("os", {}).get("LastBootUpTime", "")
    baseline["uptime_minutes"] = parse_uptime_minutes(last_boot)
    ctx.data["machine"] = baseline
    ctx.data["session"]["current_user"] = getpass.getuser()
    ctx.data["session"]["is_admin"] = is_admin()
    ctx.data["session"]["hostname"] = platform.node()
    ctx.data["session"]["timezone"] = baseline.get("timezone", {})
    ctx.data["artifacts"] = {
        "base": str(ctx.paths.base),
        "logs": str(ctx.paths.logs),
        "temp": str(ctx.paths.temp),
        "downloads": str(ctx.paths.downloads),
        "reports": str(ctx.paths.reports),
        "raw": str(ctx.paths.raw),
    }
    summarize_machine_overview(ctx)


def get_baseline_snapshot(ctx: ToolkitContext) -> Dict[str, Any]:
    days = 14 if ctx.args.deep_scan else (3 if ctx.args.quick_mode else 7)
    update_services_literal = ",".join([f"'{x}'" for x in UPDATE_SERVICES])
    script = f"""
    function Get-StartupItems {{
        $results = @()
        $registryPaths = @(
            'HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Run',
            'HKLM:\\Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Run',
            'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Run'
        )
        foreach ($path in $registryPaths) {{
            if (Test-Path $path) {{
                $props = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
                if ($props) {{
                    foreach ($prop in $props.PSObject.Properties) {{
                        if ($prop.Name -notmatch '^PS') {{
                            $results += [pscustomobject]@{{
                                Location = $path
                                Name = $prop.Name
                                Command = [string]$prop.Value
                            }}
                        }}
                    }}
                }}
            }}
        }}
        $folders = @(
            [Environment]::GetFolderPath('Startup'),
            "$env:ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Startup"
        )
        foreach ($folder in $folders) {{
            if (Test-Path $folder) {{
                Get-ChildItem -Path $folder -Force -ErrorAction SilentlyContinue | ForEach-Object {{
                    $results += [pscustomobject]@{{
                        Location = $folder
                        Name = $_.Name
                        Command = $_.FullName
                    }}
                }}
            }}
        }}
        return $results
    }}
    function Get-DefenderState {{
        if (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue) {{
            Get-MpComputerStatus | Select-Object AMServiceEnabled, AntispywareEnabled, AntivirusEnabled, BehaviorMonitorEnabled, IoavProtectionEnabled, NISEnabled, RealTimeProtectionEnabled, QuickScanAge, FullScanAge
        }}
    }}
    $os = Get-CimInstance Win32_OperatingSystem
    $cs = Get-CimInstance Win32_ComputerSystem
    $bios = Get-CimInstance Win32_BIOS
    $csProduct = Get-CimInstance Win32_ComputerSystemProduct
    $enclosure = Get-CimInstance Win32_SystemEnclosure | Select-Object -First 1 Manufacturer, ChassisTypes
    $cpu = Get-CimInstance Win32_Processor | Select-Object Name, Manufacturer, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed
    $gpu = Get-CimInstance Win32_VideoController | Select-Object Name, AdapterCompatibility, DriverVersion, Status, AdapterRAM
    $disks = Get-CimInstance Win32_DiskDrive | Select-Object Model, DeviceID, InterfaceType, MediaType, SerialNumber, Size, Status
    $volumes = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | Select-Object DeviceID, VolumeName, FileSystem, Size, FreeSpace
    $partitions = @()
    if (Get-Command Get-Partition -ErrorAction SilentlyContinue) {{
        $partitions = Get-Partition | Select-Object DiskNumber, PartitionNumber, DriveLetter, Type, Size, GptType, IsBoot, IsSystem
    }}
    $net = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object {{$_.IPEnabled}} | Select-Object Description, MACAddress, IPAddress, IPSubnet, DefaultIPGateway, DHCPEnabled, DHCPServer, DNSDomain
    $autoServicesNotRunning = Get-CimInstance Win32_Service | Where-Object {{$_.StartMode -eq 'Auto' -and $_.State -ne 'Running'}} | Select-Object -First 30 Name, DisplayName, State, Status, StartMode
    $pnpIssues = Get-CimInstance Win32_PnPEntity | Where-Object {{$_.ConfigManagerErrorCode -ne 0}} | Select-Object -First 20 Name, PNPClass, Manufacturer, ConfigManagerErrorCode
    $storageHealth = @()
    if (Get-Command Get-PhysicalDisk -ErrorAction SilentlyContinue) {{
        $storageHealth = Get-PhysicalDisk | Select-Object FriendlyName, HealthStatus, OperationalStatus, MediaType, Size, SerialNumber
    }}
    $volHealth = @()
    if (Get-Command Get-Volume -ErrorAction SilentlyContinue) {{
        $volHealth = Get-Volume | Where-Object {{$_.DriveType -eq 'Fixed'}} | Select-Object DriveLetter, FileSystemType, HealthStatus, OperationalStatus, Size, SizeRemaining
    }}
    $battery = @()
    try {{
        $battery = Get-CimInstance Win32_Battery | Select-Object BatteryStatus, EstimatedChargeRemaining, Name
    }} catch {{
        $battery = @()
    }}
    $timezone = Get-TimeZone | Select-Object Id, DisplayName
    $topProcesses = Get-Process | Sort-Object CPU -Descending | Select-Object -First 10 ProcessName, Id, CPU, WorkingSet64, PrivateMemorySize64
    $hotfixes = Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 12 HotFixID, Description, InstalledOn
    $updateServices = Get-Service -Name {update_services_literal} -ErrorAction SilentlyContinue | Select-Object Name, Status, StartType
    $history = @()
    $historyCount = 0
    try {{
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $historyCount = $searcher.GetTotalHistoryCount()
        if ($historyCount -gt 0) {{
            $history = $searcher.QueryHistory(0, [Math]::Min(12, $historyCount)) | Select-Object Date, Title, ResultCode, HResult, Operation
        }}
    }} catch {{
        $history = @()
    }}
    $since = (Get-Date).AddDays(-{days})
    $systemEvents = Get-WinEvent -FilterHashtable @{{LogName='System'; StartTime=$since; Level=1,2}} -ErrorAction SilentlyContinue
    $applicationEvents = Get-WinEvent -FilterHashtable @{{LogName='Application'; StartTime=$since; Level=1,2}} -ErrorAction SilentlyContinue
    $serviceEvents = Get-WinEvent -FilterHashtable @{{LogName='System'; ProviderName='Service Control Manager'; StartTime=$since}} -ErrorAction SilentlyContinue | Where-Object {{$_.Id -ge 7000 -and $_.Id -le 7099}}
    $crashIndicators = Get-WinEvent -FilterHashtable @{{LogName='System'; StartTime=$since; Id=41,6008,1001}} -ErrorAction SilentlyContinue
    $appCrashes = Get-WinEvent -FilterHashtable @{{LogName='Application'; StartTime=$since; Id=1000,1001}} -ErrorAction SilentlyContinue
    $whea = Get-WinEvent -FilterHashtable @{{LogName='System'; ProviderName='Microsoft-Windows-WHEA-Logger'; StartTime=$since}} -ErrorAction SilentlyContinue
    $systemSummary = $systemEvents | Group-Object ProviderName | Sort-Object Count -Descending | Select-Object -First 10 Name, Count
    $applicationSummary = $applicationEvents | Group-Object ProviderName | Sort-Object Count -Descending | Select-Object -First 10 Name, Count
    $bitlocker = @()
    if (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue) {{
        $bitlocker = Get-BitLockerVolume | Select-Object MountPoint, VolumeStatus, ProtectionStatus, EncryptionPercentage
    }}
    [pscustomobject]@{{
        os = [pscustomobject]@{{
            Caption = $os.Caption
            Version = $os.Version
            BuildNumber = $os.BuildNumber
            Architecture = $os.OSArchitecture
            LastBootUpTime = $os.LastBootUpTime
            InstallDate = $os.InstallDate
            Locale = $os.Locale
        }}
        computer = [pscustomobject]@{{
            Hostname = $env:COMPUTERNAME
            Manufacturer = $cs.Manufacturer
            Model = $cs.Model
            Domain = $cs.Domain
            PartOfDomain = $cs.PartOfDomain
            UserName = $cs.UserName
            TotalPhysicalMemory = $cs.TotalPhysicalMemory
            SerialNumber = $bios.SerialNumber
            UUID = $csProduct.UUID
            PCSystemType = $cs.PCSystemType
            PCSystemTypeEx = $cs.PCSystemTypeEx
        }}
        enclosure = if ($enclosure) {{ [pscustomobject]@{{ Manufacturer = $enclosure.Manufacturer; ChassisTypes = $enclosure.ChassisTypes }} }} else {{ [pscustomobject]@{{ Manufacturer = ''; ChassisTypes = @() }} }}
        timezone = $timezone
        cpu = $cpu
        gpu = $gpu
        disks = $disks
        partitions = $partitions
        logicalDisks = $volumes
        volumeHealth = $volHealth
        network = $net
        autoServicesNotRunning = $autoServicesNotRunning
        pnpIssues = $pnpIssues
        storageHealth = $storageHealth
        battery = $battery
        topProcesses = $topProcesses
        startupItems = Get-StartupItems
        hotfixes = $hotfixes
        updateServices = $updateServices
        updateHistory = $history
        updateHistoryCount = $historyCount
        eventSummary = [pscustomobject]@{{
            SystemErrorCount = @($systemEvents).Count
            ApplicationErrorCount = @($applicationEvents).Count
            ServiceFailureCount = @($serviceEvents).Count
            CrashIndicatorCount = @($crashIndicators).Count
            AppCrashCount = @($appCrashes).Count
            WHEACount = @($whea).Count
            SystemTopProviders = $systemSummary
            ApplicationTopProviders = $applicationSummary
        }}
        defender = Get-DefenderState
        bitlocker = $bitlocker
        safeMode = [bool]$env:SAFEBOOT_OPTION
    }}
    """
    data, _ = run_powershell_json(ctx, "Stage 1", "Collect baseline system snapshot", script, timeout=420, depth=10)
    return data or {}


def detect_enterprise_management(ctx: ToolkitContext) -> Dict[str, Any]:
    management = {
        "domain_joined": bool(ctx.data.get("machine", {}).get("computer", {}).get("PartOfDomain")),
        "domain_name": ctx.data.get("machine", {}).get("computer", {}).get("Domain"),
        "mdm_enrolled": False,
        "sccm_present": False,
        "enrollment_key_count": 0,
        "enrollment_signals": [],
        "enterprise_managed": False,
    }
    if winreg:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Enrollments") as key:
                subkey_count = winreg.QueryInfoKey(key)[0]
                management["enrollment_key_count"] = subkey_count
                signals: List[str] = []
                for index in range(subkey_count):
                    subkey_name = winreg.EnumKey(key, index)
                    if not re.fullmatch(r"[0-9A-Fa-f-]{36}", subkey_name):
                        continue
                    subkey_path = rf"SOFTWARE\Microsoft\Enrollments\{subkey_name}"
                    for value_name in ("ProviderID", "UPN", "DiscoveryServiceFullURL", "EnrollmentServiceUrl"):
                        value = reg_value(winreg.HKEY_LOCAL_MACHINE, subkey_path, value_name)
                        if value and str(value).strip():
                            signals.append(f"{subkey_name}:{value_name}")
                            break
                management["enrollment_signals"] = signals[:8]
                management["mdm_enrolled"] = bool(signals)
        except Exception:
            management["mdm_enrolled"] = False
    service_script = """
    $svc = Get-Service -Name CcmExec -ErrorAction SilentlyContinue
    if ($svc) { [pscustomobject]@{ Present = $true; Status = $svc.Status.ToString() } } else { [pscustomobject]@{ Present = $false; Status = '' } }
    """
    service_data, _ = run_powershell_json(ctx, "Prechecks", "Detect enterprise management services", service_script, timeout=45)
    management["sccm_present"] = bool((service_data or {}).get("Present"))
    management["enterprise_managed"] = any([management["domain_joined"], management["mdm_enrolled"], management["sccm_present"]])
    return management


def precheck_restore_point(ctx: ToolkitContext) -> Tuple[bool, str]:
    if ctx.args.report_only:
        return False, "Skipped in report-only mode"
    script = f"""
    if (Get-Command Checkpoint-Computer -ErrorAction SilentlyContinue) {{
        Checkpoint-Computer -Description '{TOOL_NAME} {ctx.session_id}' -RestorePointType 'MODIFY_SETTINGS'
        [pscustomobject]@{{ Success = $true; Message = 'Restore point created' }}
    }} else {{
        [pscustomobject]@{{ Success = $false; Message = 'Checkpoint-Computer unavailable' }}
    }}
    """
    data, result = run_powershell_json(ctx, "Prechecks", "Create system restore point", script, timeout=240)
    if data and data.get("Success"):
        return True, data.get("Message", "Restore point created")
    if result["ok"]:
        return False, data.get("Message", "Restore point not created") if isinstance(data, dict) else "Restore point not created"
    return False, result["stderr"].strip() or "Restore point creation failed"


def analyze_baseline(ctx: ToolkitContext) -> None:
    machine = ctx.data.get("machine", {})
    logical_disks = flatten_to_list(machine.get("logicalDisks"))
    startup_items = flatten_to_list(machine.get("startupItems"))
    event_summary = machine.get("eventSummary", {})
    pnp_issues = flatten_to_list(machine.get("pnpIssues"))
    auto_services = flatten_to_list(machine.get("autoServicesNotRunning"))
    volume_health = flatten_to_list(machine.get("volumeHealth"))
    storage_health = flatten_to_list(machine.get("storageHealth"))
    update_history = flatten_to_list(machine.get("updateHistory"))
    if len(startup_items) >= 20:
        ctx.add_finding("warning", "High startup load", f"{len(startup_items)} startup items were detected.", "performance")
        ctx.add_recommendation("Review unnecessary startup applications for performance improvement.")
    if pnp_issues:
        ctx.add_finding("warning", "Device driver issues detected", f"{len(pnp_issues)} device(s) reported a PnP error.", "drivers")
        ctx.add_recommendation("Review devices reporting ConfigManager errors and update or reseat hardware if needed.")
    critical_down = [svc for svc in auto_services if svc.get("Name") in CRITICAL_SERVICES]
    if critical_down:
        ctx.add_finding(
            "warning",
            "Critical automatic services not running",
            ", ".join([svc.get("Name", "") for svc in critical_down]),
            "services",
        )
    if int(event_summary.get("CrashIndicatorCount", 0) or 0) > 0:
        ctx.add_finding(
            "major",
            "System crash indicators found",
            f"{event_summary.get('CrashIndicatorCount')} recent crash-related system events were detected.",
            "stability",
        )
        ctx.add_recommendation("Review recent crash events and confirm hardware stability after servicing.")
    if int(event_summary.get("WHEACount", 0) or 0) > 0:
        ctx.add_hardware_suspicion("WHEA hardware error events were detected.")
        ctx.add_recommendation("Investigate CPU, RAM, motherboard, and storage health due to WHEA events.")
    if int(event_summary.get("AppCrashCount", 0) or 0) >= 3:
        ctx.add_finding(
            "warning",
            "Repeated application crashes detected",
            f"{event_summary.get('AppCrashCount')} application crash indicators were detected.",
            "stability",
        )
    if not machine.get("networking", {}).get("internet_available", False):
        ctx.detected_network_issue = True
        ctx.add_finding("warning", "Internet connectivity unavailable", "Outbound connectivity tests failed.", "network")
        ctx.add_recommendation("Verify upstream internet connectivity if update repair is required.")
    for volume in volume_health:
        if str(volume.get("HealthStatus", "")).lower() not in ("", "healthy"):
            drive = volume.get("DriveLetter")
            ctx.add_finding(
                "warning",
                "Volume health warning",
                f"Volume {drive}: health status = {volume.get('HealthStatus')}, operational status = {volume.get('OperationalStatus')}",
                "storage",
            )
    for disk in storage_health:
        if str(disk.get("HealthStatus", "")).lower() not in ("", "healthy"):
            name = disk.get("FriendlyName") or disk.get("Model") or "Unknown disk"
            ctx.add_hardware_suspicion(f"{name} reported health status {disk.get('HealthStatus')}.")
            ctx.add_recommendation("Storage diagnostics and possible hardware replacement should be considered.")
    failed_updates = [u for u in update_history if int(u.get("ResultCode", 0) or 0) >= 4]
    if failed_updates:
        ctx.add_finding(
            "warning",
            "Recent Windows Update failures detected",
            f"{len(failed_updates)} recent update history entries show failure or abort states.",
            "updates",
        )
    for volume in logical_disks:
        drive = volume.get("DeviceID")
        if not drive:
            continue
        free = int(volume.get("FreeSpace", 0) or 0)
        size = int(volume.get("Size", 0) or 0)
        percent = (free / size * 100) if size else 0
        if percent < 15:
            ctx.add_finding(
                "warning",
                "Low free space on fixed drive",
                f"{drive} has {format_bytes(free)} free ({percent:.2f}%).",
                "storage",
            )


def run_stage_baseline(ctx: ToolkitContext) -> None:
    stage_name = "Stage 1 - Baseline System Diagnostics"
    start = dt.datetime.now().astimezone()
    ctx.console.section(stage_name)
    baseline = get_baseline_snapshot(ctx)
    enrich_machine_context(ctx, baseline)
    analyze_baseline(ctx)
    wmi = run_command(ctx, stage_name, "Verify WMI repository", ["winmgmt.exe", "/verifyrepository"], timeout=120, acceptable_exit_codes=(0, 1))
    wmi_text = (wmi["stdout"] + "\n" + wmi["stderr"]).lower()
    ctx.data["diagnostics"]["wmi_repository"] = {
        "exit_code": wmi["exit_code"],
        "stdout_path": wmi["stdout_path"],
        "stderr_path": wmi["stderr_path"],
        "status": "inconsistent" if "inconsistent" in wmi_text else ("consistent" if "consistent" in wmi_text else "unknown"),
    }
    if "inconsistent" in wmi_text:
        ctx.add_finding("warning", "WMI repository consistency issue indicated", "winmgmt /verifyrepository reported an inconsistent repository.", "wmi")
        ctx.add_recommendation("Review WMI repository health if management or hardware queries remain unreliable.")
    notes = [
        f"Detected {len(flatten_to_list(baseline.get('logicalDisks')))} fixed volumes",
        f"Detected {len(flatten_to_list(baseline.get('partitions')))} partition entries",
        f"Detected {len(flatten_to_list(baseline.get('gpu')))} GPU(s)",
        f"Detected {len(flatten_to_list(baseline.get('network')))} active IP-enabled network adapter(s)",
    ]
    end = dt.datetime.now().astimezone()
    ctx.record_stage(stage_name, "OK", start, end, "Baseline diagnostic collection completed.", notes=notes)


def perform_safety_prechecks(ctx: ToolkitContext) -> Dict[str, Any]:
    ctx.console.section("Safety Pre-Checks")
    machine = ctx.data.get("machine", {})
    logical_disks = flatten_to_list(machine.get("logicalDisks"))
    system_drive = SYSTEM_DRIVE_ENV.rstrip("\\")
    system_volume = next((d for d in logical_disks if str(d.get("DeviceID", "")).upper() == system_drive.upper()), {})
    total = int(system_volume.get("Size", 0) or 0)
    free = int(system_volume.get("FreeSpace", 0) or 0)
    free_percent = round((free / total) * 100, 2) if total else 0.0
    portable_machine = is_portable_machine(machine)
    raw_battery = flatten_to_list(machine.get("battery"))
    battery = [item for item in raw_battery if isinstance(item, dict)] if portable_machine else []
    battery_present = bool(battery)
    battery_status = battery[0].get("BatteryStatus") if battery else None
    ac_power = battery_status in (2, 6, 7, 8, 9, 11)
    pending = pending_reboot_reasons()
    management = detect_enterprise_management(ctx)
    bitlocker = flatten_to_list(machine.get("bitlocker"))
    critically_low_space = free < 5 * 1024 ** 3 or free_percent < 10.0
    major_blockers = []
    if critically_low_space:
        major_blockers.append("System drive free space is critically low")
    if is_safe_mode() or bool(machine.get("safeMode")):
        major_blockers.append("Safe Mode is active")
    if major_blockers and not ctx.args.report_only:
        restore_success, restore_message = False, "Skipped because safety blockers are present"
    else:
        restore_success, restore_message = precheck_restore_point(ctx)
    prechecks = {
        "system_drive": system_drive,
        "system_drive_total_bytes": total,
        "system_drive_free_bytes": free,
        "system_drive_free_percent": free_percent,
        "critically_low_space": critically_low_space,
        "portable_machine": portable_machine,
        "battery_present": battery_present,
        "battery_status": battery_status,
        "on_ac_power": ac_power if battery_present else True,
        "pending_reboot": bool(pending),
        "pending_reboot_reasons": pending,
        "safe_mode": is_safe_mode() or bool(machine.get("safeMode")),
        "management": management,
        "bitlocker": bitlocker,
        "restore_point_created": restore_success,
        "restore_point_message": restore_message,
        "major_blockers": major_blockers,
        "repair_operations_allowed": not major_blockers and not ctx.args.report_only,
    }
    ctx.data["prechecks"] = prechecks
    if critically_low_space:
        ctx.console.warn(f"System drive free space is low: {format_bytes(free)} ({free_percent:.2f}%)")
        ctx.add_finding("major", "Low system drive space", f"System drive {system_drive} has only {format_bytes(free)} free.", "storage")
        ctx.add_recommendation("Free additional space on the system drive or upgrade storage capacity.")
    else:
        ctx.console.ok(f"System drive free space: {format_bytes(free)} ({free_percent:.2f}%)")
    if pending:
        ctx.console.warn(f"Pending reboot detected: {', '.join(pending)}")
        ctx.reboot_required = True
        ctx.add_finding("warning", "Pending reboot already present", "; ".join(pending), "servicing")
    else:
        ctx.console.ok("No existing pending reboot state detected")
    if raw_battery and not portable_machine:
        ctx.logger.info("Battery-like telemetry was reported on a non-portable chassis; battery AC heuristics were ignored.")
    if battery_present:
        if ac_power:
            ctx.console.ok("Laptop battery detected; AC power appears connected")
        else:
            ctx.console.warn("Laptop battery detected; AC power not detected")
            ctx.add_recommendation("Connect AC power before long servicing or update operations.")
    if management["enterprise_managed"]:
        ctx.console.warn("Enterprise or domain management detected; policy-sensitive actions will be constrained")
        ctx.add_finding(
            "warning",
            "Enterprise management detected",
            "Domain, MDM, or SCCM management is present. Policy-sensitive resets and changes are constrained.",
            "policy",
        )
    if bitlocker:
        protected = [vol.get("MountPoint") for vol in bitlocker if str(vol.get("ProtectionStatus", "")).strip() not in ("0", "Off", "")]
        if protected:
            ctx.console.info(f"BitLocker protection present on: {', '.join([p for p in protected if p])}")
    if restore_success:
        ctx.console.ok("Restore point created successfully")
    else:
        ctx.console.warn(f"Restore point creation was not completed: {restore_message}")
    if major_blockers:
        blocker_text = "; ".join(major_blockers)
        ctx.console.fail(f"Safety blockers detected: {blocker_text}")
        ctx.add_finding("major", "Safety blockers prevent automated repair", blocker_text, "safety")
        ctx.add_unresolved("Automated repair actions were skipped because safety blockers were detected.")
        ctx.add_recommendation("Resolve the safety blockers before rerunning automated repair.")
    return prechecks


def summarize_dism_result(command_name: str, result: Dict[str, Any]) -> str:
    if not result["ok"]:
        codes = parse_hresults(result["stdout"] + "\n" + result["stderr"])
        code_text = f" Error codes: {', '.join(codes)}." if codes else ""
        return f"{command_name} did not complete cleanly.{code_text}"
    stdout = result["stdout"].replace("\x00", "").lower()
    if "no component store corruption detected" in stdout:
        return f"{command_name} reported no component store corruption."
    if "component store corruption repaired" in stdout or "restore operation completed successfully" in stdout:
        return f"{command_name} completed and reported repair activity."
    if "windows resource protection did not find any integrity violations" in stdout:
        return "SFC found no integrity violations."
    if "windows resource protection found corrupt files and successfully repaired them" in stdout:
        return "SFC repaired corrupt files."
    if "windows resource protection found corrupt files but was unable to fix some of them" in stdout:
        return "SFC found corruption but could not repair all files."
    return f"{command_name} completed with exit code {result['exit_code']}."


def make_sfc_progress_probe(mode: str = "scan") -> Any:
    state = {"last_cpu": None, "last_cbs_write": None, "last_cbs_size": None, "stagnant_seconds": 0}
    probe_interval = 30
    stall_threshold_seconds = (
        SFC_VERIFY_STALL_THRESHOLD_SECONDS if mode == "verify" else SFC_SCAN_STALL_THRESHOLD_SECONDS
    )

    def _probe(elapsed_seconds: int) -> Dict[str, Any]:
        script = r"""
        $proc = Get-Process -Name sfc -ErrorAction SilentlyContinue | Select-Object -First 1 Id, CPU, WorkingSet64
        $cbs = Get-Item "$env:windir\Logs\CBS\CBS.log" -ErrorAction SilentlyContinue
        $trusted = Get-Service -Name TrustedInstaller -ErrorAction SilentlyContinue
        [pscustomobject]@{
            Running = [bool]$proc
            CPU = if ($proc) { [math]::Round($proc.CPU, 1) } else { $null }
            WorkingSetMB = if ($proc) { [math]::Round($proc.WorkingSet64 / 1MB, 1) } else { $null }
            CBSLastWrite = if ($cbs) { $cbs.LastWriteTime.ToUniversalTime().ToString('o') } else { $null }
            CBSSizeMB = if ($cbs) { [math]::Round($cbs.Length / 1MB, 1) } else { $null }
            TrustedInstaller = if ($trusted) { $trusted.Status.ToString() } else { 'Unknown' }
        }
        """
        data = run_powershell_json_quiet(script, timeout=20, depth=4)
        if not isinstance(data, dict):
            return {"note": "live probe unavailable", "stalled": False}

        running = bool(data.get("Running"))
        cpu = float(data.get("CPU") or 0.0)
        cbs_write = data.get("CBSLastWrite")
        cbs_size = float(data.get("CBSSizeMB") or 0.0)
        progressed = False
        if running and (state["last_cpu"] is None or cpu > float(state["last_cpu"] or 0.0) + 0.05):
            progressed = True
        if cbs_write and cbs_write != state["last_cbs_write"]:
            progressed = True
        if state["last_cbs_size"] is None or cbs_size > float(state["last_cbs_size"] or 0.0) + 0.2:
            progressed = True
        if progressed:
            state["stagnant_seconds"] = 0
        else:
            state["stagnant_seconds"] += probe_interval

        state["last_cpu"] = cpu
        if cbs_write:
            state["last_cbs_write"] = cbs_write
        state["last_cbs_size"] = cbs_size

        cbs_age = "unknown"
        if cbs_write:
            try:
                cbs_dt = dt.datetime.fromisoformat(str(cbs_write).replace("Z", "+00:00"))
                age_seconds = max(0, int((dt.datetime.now(dt.timezone.utc) - cbs_dt).total_seconds()))
                cbs_age = minutes_to_human(age_seconds / 60.0) + " ago"
            except Exception:
                cbs_age = "unknown"

        note = (
            f"SFC {'running' if running else 'not visible'}"
            f", CPU {cpu:.1f}s"
            f", RAM {float(data.get('WorkingSetMB') or 0):.0f}MB"
            f", CBS {cbs_age}"
            f", log {cbs_size:.1f}MB"
            f", TrustedInstaller {data.get('TrustedInstaller', 'Unknown')}"
        )
        stalled = (
            elapsed_seconds >= SFC_MIN_RUNTIME_BEFORE_STALL_SECONDS
            and state["stagnant_seconds"] >= stall_threshold_seconds
        )
        return {"note": note, "stalled": stalled}

    return _probe


def run_stage_servicing(ctx: ToolkitContext) -> Dict[str, Any]:
    stage_name = "Stage 2 - Windows Servicing and Repair"
    start = dt.datetime.now().astimezone()
    ctx.console.section(stage_name)
    skip_reason = stage_skip_reason(
        ctx,
        "Report-only mode active; servicing repair commands are skipped",
        "Servicing repair commands were skipped because safety blockers are present",
    )
    if skip_reason:
        ctx.console.warn(skip_reason)
        summary = {"skipped": True, "reason": skip_reason}
        ctx.data["diagnostics"]["servicing"] = summary
        end = dt.datetime.now().astimezone()
        ctx.record_stage(stage_name, "SKIPPED", start, end, skip_reason)
        return summary
    commands = [
        ("DISM CheckHealth", ["dism.exe", "/Online", "/Cleanup-Image", "/CheckHealth"]),
        ("DISM ScanHealth", ["dism.exe", "/Online", "/Cleanup-Image", "/ScanHealth"]),
        ("DISM RestoreHealth", ["dism.exe", "/Online", "/Cleanup-Image", "/RestoreHealth"]),
        ("DISM StartComponentCleanup", ["dism.exe", "/Online", "/Cleanup-Image", "/StartComponentCleanup"]),
        ("SFC Scannow", ["sfc.exe", "/scannow"]),
    ]
    results: Dict[str, Any] = {}
    notes: List[str] = []
    stage_status = "OK"
    for label, cmd in commands:
        timeout = (
            5400
            if label in ("DISM ScanHealth", "DISM RestoreHealth")
            else (SFC_SCAN_TIMEOUT_SECONDS if label == "SFC Scannow" else 1800)
        )
        heartbeat_interval = 60 if label == "SFC Scannow" else 15
        progress_probe = make_sfc_progress_probe("scan") if label == "SFC Scannow" else None
        acceptable_exit_codes = (0, 1) if label == "SFC Scannow" else (0,)
        if label == "SFC Scannow":
            ctx.console.info(
                "SFC has a 60-minute safety cap and live stall detection so it does not sit silently for hours."
            )
        result = run_command(
            ctx,
            stage_name,
            label,
            cmd,
            timeout=timeout,
            acceptable_exit_codes=acceptable_exit_codes,
            heartbeat_interval=heartbeat_interval,
            progress_probe=progress_probe,
            progress_probe_interval=30,
        )
        summary = summarize_dism_result(label, result)
        results[label] = {
            "summary": summary,
            "exit_code": result["exit_code"],
            "stdout_path": result["stdout_path"],
            "stderr_path": result["stderr_path"],
            "duration_seconds": round(result["duration_seconds"], 2),
            "timed_out": result["timed_out"],
            "stalled": result["stalled"],
        }
        notes.append(summary)
        if not result["ok"]:
            stage_status = "WARN"
            if label == "SFC Scannow":
                if result["timed_out"]:
                    ctx.add_finding(
                        "warning",
                        "SFC runtime exceeded the safety limit",
                        "SFC was stopped after exceeding the configured runtime or appearing stalled.",
                        "servicing",
                    )
                    ctx.add_recommendation("Review CBS.log and rerun SFC after reboot only if servicing issues remain.")
                ctx.add_unresolved("SFC did not complete cleanly; review CBS and raw outputs.")
            else:
                ctx.add_unresolved(f"{label} did not complete cleanly; review DISM outputs.")
    ctx.data["diagnostics"]["servicing"] = results
    restore_summary = results.get("DISM RestoreHealth", {}).get("summary", "")
    sfc_summary = results.get("SFC Scannow", {}).get("summary", "")
    if "could not repair all files" in sfc_summary.lower():
        ctx.add_finding("major", "SFC could not repair all files", sfc_summary, "servicing")
        ctx.add_unresolved("System file corruption remains after SFC.")
    elif "repaired corrupt files" in sfc_summary.lower():
        ctx.add_finding("warning", "SFC repaired system files", sfc_summary, "servicing")
        ctx.reboot_required = True
    if "did not complete cleanly" in restore_summary.lower():
        ctx.add_finding("major", "DISM RestoreHealth did not complete cleanly", restore_summary, "servicing")
        ctx.add_unresolved("Windows component servicing health remains uncertain.")
    end = dt.datetime.now().astimezone()
    ctx.record_stage(stage_name, stage_status, start, end, "Windows servicing workflow completed.", notes=notes)
    return results


def query_dirty_bit(ctx: ToolkitContext, drive: str) -> Dict[str, Any]:
    return run_command(ctx, "Stage 3", f"Query dirty bit on {drive}", ["fsutil.exe", "dirty", "query", drive], timeout=60)


def run_chkdsk_scan(ctx: ToolkitContext, drive: str) -> Dict[str, Any]:
    return run_command(ctx, "Stage 3", f"CHKDSK scan {drive}", ["chkdsk.exe", drive, "/scan"], timeout=5400)


def schedule_chkdsk_if_needed(ctx: ToolkitContext, drive: str) -> Tuple[bool, str]:
    if ctx.args.report_only:
        return False, "Skipped in report-only mode"
    blockers = active_safety_blockers(ctx)
    if blockers:
        return False, f"Repair scheduling blocked by safety checks: {', '.join(blockers)}"
    if not ctx.args.allow_reboot:
        return False, "Repair scheduling not allowed without --allow-reboot"
    schedule_cmd = f'cmd.exe /c "echo Y|chkdsk {drive} /F"'
    result = run_command(ctx, "Stage 3", f"Schedule CHKDSK repair {drive}", schedule_cmd, timeout=120, shell=True, acceptable_exit_codes=(0, 1))
    if result["ok"]:
        ctx.reboot_required = True
        return True, "CHKDSK repair scheduled for next reboot"
    return False, "Unable to schedule CHKDSK repair"


def run_stage_disk_health(ctx: ToolkitContext) -> Dict[str, Any]:
    stage_name = "Stage 3 - Disk and Filesystem Health"
    start = dt.datetime.now().astimezone()
    ctx.console.section(stage_name)
    logical_disks = flatten_to_list(ctx.data.get("machine", {}).get("logicalDisks"))
    disk_results: Dict[str, Any] = {"volumes": []}
    stage_status = "OK"
    blockers = active_safety_blockers(ctx)
    for volume in logical_disks:
        drive = volume.get("DeviceID")
        if not drive:
            continue
        dirty = query_dirty_bit(ctx, drive)
        scan = run_chkdsk_scan(ctx, drive)
        dirty_text = (dirty["stdout"] + "\n" + dirty["stderr"]).lower()
        scan_text = (scan["stdout"] + "\n" + scan["stderr"]).lower()
        needs_repair = ("is dirty" in dirty_text) or ("found problems" in scan_text) or ("run chkdsk /f" in scan_text)
        scheduled = False
        schedule_message = ""
        if needs_repair:
            stage_status = "WARN"
            ctx.add_finding("warning", f"Filesystem issues indicated on {drive}", "CHKDSK scan suggests repair may be required.", "storage")
            if drive.upper() == SYSTEM_DRIVE_ENV.upper():
                scheduled, schedule_message = schedule_chkdsk_if_needed(ctx, drive)
            elif blockers:
                schedule_message = f"Repair blocked by safety checks: {', '.join(blockers)}"
            else:
                repair_result = run_command(
                    ctx,
                    stage_name,
                    f"Attempt CHKDSK repair {drive}",
                    ["chkdsk.exe", drive, "/F"],
                    timeout=3600,
                    acceptable_exit_codes=(0, 1),
                )
                scheduled = repair_result["ok"]
                schedule_message = "Offline CHKDSK repair attempted" if repair_result["ok"] else "Repair attempt failed"
        disk_results["volumes"].append(
            {
                "drive": drive,
                "dirty_query_exit_code": dirty["exit_code"],
                "dirty_stdout_path": dirty["stdout_path"],
                "scan_exit_code": scan["exit_code"],
                "scan_stdout_path": scan["stdout_path"],
                "needs_repair": needs_repair,
                "repair_scheduled": scheduled,
                "repair_message": schedule_message,
            }
        )
    ctx.data["diagnostics"]["disk"] = disk_results
    if any(item.get("needs_repair") for item in disk_results["volumes"]):
        if blockers:
            ctx.add_recommendation("Review CHKDSK findings and clear safety blockers before attempting automated filesystem repair.")
        else:
            ctx.add_recommendation("Review CHKDSK findings and allow reboot if filesystem repair was scheduled.")
    end = dt.datetime.now().astimezone()
    ctx.record_stage(stage_name, stage_status, start, end, "Disk and filesystem checks completed.")
    return disk_results


def safe_remove_contents(path: Path, logger: logging.Logger, older_than_hours: int = 6) -> Tuple[int, int]:
    removed_files = 0
    removed_bytes = 0
    threshold = time.time() - older_than_hours * 3600
    if not path.exists():
        return removed_files, removed_bytes
    for root, dirs, files in os.walk(path, topdown=False):
        for name in files:
            file_path = Path(root) / name
            try:
                stat = file_path.stat()
                if stat.st_mtime > threshold:
                    continue
                removed_bytes += stat.st_size
                file_path.unlink()
                removed_files += 1
            except Exception:
                continue
        for name in dirs:
            dir_path = Path(root) / name
            try:
                if not any(dir_path.iterdir()):
                    dir_path.rmdir()
            except Exception:
                continue
    logger.info("Removed %s files (%s) from %s", removed_files, format_bytes(removed_bytes), path)
    return removed_files, removed_bytes


def service_status(ctx: ToolkitContext, service_name: str) -> Dict[str, Any]:
    script = f"""
    $svc = Get-Service -Name '{service_name}' -ErrorAction SilentlyContinue
    if ($svc) {{
        [pscustomobject]@{{ Exists = $true; Name = $svc.Name; Status = $svc.Status.ToString(); StartType = $svc.StartType.ToString() }}
    }} else {{
        [pscustomobject]@{{ Exists = $false; Name = '{service_name}'; Status = 'Missing'; StartType = '' }}
    }}
    """
    data, _ = run_powershell_json(ctx, "Stage 4", f"Query service {service_name}", script, timeout=30)
    return data or {"Exists": False, "Name": service_name, "Status": "Unknown", "StartType": ""}


def restart_service_if_needed(ctx: ToolkitContext, service_name: str, reason: str) -> str:
    svc = service_status(ctx, service_name)
    if not svc.get("Exists"):
        return "Service not present"
    if svc.get("StartType") == "Disabled":
        return "Service disabled; left unchanged"
    action = "Restart" if svc.get("Status") == "Running" else "Start"
    cmd = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", f"{action}-Service -Name '{service_name}' -ErrorAction Stop"]
    result = run_command(ctx, "Stage 4", f"{action} service {service_name}", cmd, timeout=120)
    if result["ok"]:
        return f"{action}ed for {reason}"
    return f"{action} failed for {reason}"


def run_stage_cleanup(ctx: ToolkitContext) -> Dict[str, Any]:
    stage_name = "Stage 4 - Safe Cleanup and Maintenance"
    start = dt.datetime.now().astimezone()
    ctx.console.section(stage_name)
    skip_reason = stage_skip_reason(
        ctx,
        "Report-only mode active; maintenance actions are skipped",
        "Maintenance actions were skipped because safety blockers are present",
    )
    if skip_reason:
        ctx.console.warn(skip_reason)
        summary = {"skipped": True, "reason": skip_reason}
        ctx.data["diagnostics"]["cleanup"] = summary
        end = dt.datetime.now().astimezone()
        ctx.record_stage(stage_name, "SKIPPED", start, end, skip_reason)
        return summary
    cleanup_summary: Dict[str, Any] = {"temp_cleanup": [], "service_actions": []}
    targets = [Path(tempfile.gettempdir()), Path(os.environ.get("TEMP", tempfile.gettempdir())), Path(os.environ.get("WINDIR", r"C:\Windows")) / "Temp"]
    seen = set()
    for target in targets:
        normalized = str(target).lower()
        if normalized in seen:
            continue
        seen.add(normalized)
        files_removed, bytes_removed = safe_remove_contents(target, ctx.logger, older_than_hours=(3 if ctx.args.quick_mode else 6))
        cleanup_summary["temp_cleanup"].append({"path": str(target), "files_removed": files_removed, "bytes_removed": bytes_removed})
        ctx.console.ok(f"Cleaned {files_removed} temp files from {target} ({format_bytes(bytes_removed)})")
    dns_result = run_command(ctx, stage_name, "Flush DNS cache", ["ipconfig.exe", "/flushdns"], timeout=45)
    cleanup_summary["dns_flush"] = {"ok": dns_result["ok"], "stdout_path": dns_result["stdout_path"]}
    for service_name, reason in [("W32Time", "time sync verification"), ("wuauserv", "Windows Update verification"), ("BITS", "Windows Update transfer verification")]:
        outcome = restart_service_if_needed(ctx, service_name, reason)
        cleanup_summary["service_actions"].append({"service": service_name, "result": outcome})
    w32tm = run_command(ctx, stage_name, "Query Windows Time status", ["w32tm.exe", "/query", "/status"], timeout=45, acceptable_exit_codes=(0, 1))
    cleanup_summary["time_status"] = {"exit_code": w32tm["exit_code"], "stdout_path": w32tm["stdout_path"]}
    ctx.data["diagnostics"]["cleanup"] = cleanup_summary
    end = dt.datetime.now().astimezone()
    ctx.record_stage(stage_name, "OK", start, end, "Safe cleanup and maintenance actions completed.")
    return cleanup_summary


def update_search_script(install_updates: bool) -> str:
    install_block = """
    $downloadCollection = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($item in $searchResult.Updates) {
        if (-not $item.EulaAccepted) { $item.AcceptEula() }
        [void]$downloadCollection.Add($item)
    }
    if ($downloadCollection.Count -gt 0) {
        $downloader = $session.CreateUpdateDownloader()
        $downloader.Updates = $downloadCollection
        $downloadResult = $downloader.Download()
        $output.Download = [pscustomobject]@{
            ResultCode = [int]$downloadResult.ResultCode
            ResultText = [string]$downloadResult.ResultCode
            HResult = ('0x{0:X8}' -f ($downloadResult.HResult -band 0xffffffff))
        }
        $installCollection = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($item in $downloadCollection) {
            if ($item.IsDownloaded) { [void]$installCollection.Add($item) }
        }
        if ($installCollection.Count -gt 0) {
            $installer = $session.CreateUpdateInstaller()
            $installer.Updates = $installCollection
            $installResult = $installer.Install()
            $perUpdate = @()
            for ($i = 0; $i -lt $installCollection.Count; $i++) {
                $update = $installCollection.Item($i)
                $result = $installResult.GetUpdateResult($i)
                $perUpdate += [pscustomobject]@{
                    Title = $update.Title
                    KB = ($update.KBArticleIDs -join ',')
                    ResultCode = [int]$result.ResultCode
                    HResult = ('0x{0:X8}' -f ($result.HResult -band 0xffffffff))
                    RebootRequired = $update.RebootRequired
                }
            }
            $output.Install = [pscustomobject]@{
                ResultCode = [int]$installResult.ResultCode
                ResultText = [string]$installResult.ResultCode
                RebootRequired = $installResult.RebootRequired
                Updates = $perUpdate
            }
            $output.RebootRequired = $installResult.RebootRequired
        }
    }
    """ if install_updates else """
    $output.Download = [pscustomobject]@{
        ResultCode = -1
        ResultText = 'Install skipped'
        HResult = ''
    }
    """
    return f"""
    $session = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $criteria = "IsInstalled=0 and IsHidden=0 and Type='Software'"
    $searchResult = $searcher.Search($criteria)
    $updates = @()
    foreach ($item in $searchResult.Updates) {{
        $updates += [pscustomobject]@{{
            Title = $item.Title
            KB = ($item.KBArticleIDs -join ',')
            IsDownloaded = $item.IsDownloaded
            RebootRequired = $item.RebootRequired
            Categories = ($item.Categories | ForEach-Object {{$_.Name}})
        }}
    }}
    $output = [ordered]@{{
        Criteria = $criteria
        FoundCount = $searchResult.Updates.Count
        SearchResultCode = [int]$searchResult.ResultCode
        Updates = $updates
        Download = $null
        Install = $null
        RebootRequired = $false
    }}
    {install_block}
    [pscustomobject]$output
    """


def repair_update_components(ctx: ToolkitContext) -> Tuple[bool, str]:
    management = ctx.data.get("prechecks", {}).get("management", {})
    if management.get("enterprise_managed"):
        return False, "Skipped update component reset because enterprise management is detected"
    if ctx.args.report_only:
        return False, "Skipped in report-only mode"
    suffix = run_timestamp()
    script = f"""
    $services = 'BITS','wuauserv','CryptSvc','msiserver'
    foreach ($svc in $services) {{
        Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
    }}
    Start-Sleep -Seconds 2
    $sw = Join-Path $env:SystemRoot 'SoftwareDistribution'
    $cr = Join-Path $env:SystemRoot 'System32\\catroot2'
    if (Test-Path $sw) {{ Rename-Item -Path $sw -NewName ('SoftwareDistribution.bak.{suffix}') -ErrorAction Stop }}
    if (Test-Path $cr) {{ Rename-Item -Path $cr -NewName ('catroot2.bak.{suffix}') -ErrorAction Stop }}
    foreach ($svc in $services) {{
        Start-Service -Name $svc -ErrorAction SilentlyContinue
    }}
    [pscustomobject]@{{ Success = $true; Message = 'Windows Update components were reset safely.' }}
    """
    data, result = run_powershell_json(ctx, "Stage 5", "Repair Windows Update components", script, timeout=300)
    if result["ok"] and data and data.get("Success"):
        return True, data.get("Message", "Update components reset")
    return False, result["stderr"].strip() or "Update component repair failed"


def run_stage_windows_update(ctx: ToolkitContext) -> Dict[str, Any]:
    stage_name = "Stage 5 - Windows Update Repair and Installation"
    start = dt.datetime.now().astimezone()
    ctx.console.section(stage_name)
    summary: Dict[str, Any] = {"skipped": False}
    skip_reason = stage_skip_reason(
        ctx,
        "Report-only mode active; Windows Update changes are skipped",
        "Windows Update actions were skipped because safety blockers are present",
    )
    if skip_reason:
        ctx.console.warn(skip_reason)
        summary["skipped"] = True
        summary["reason"] = skip_reason
        ctx.data["diagnostics"]["windows_update"] = summary
        end = dt.datetime.now().astimezone()
        ctx.record_stage(stage_name, "SKIPPED", start, end, skip_reason)
        return summary
    if not ctx.data.get("machine", {}).get("networking", {}).get("internet_available"):
        ctx.console.warn("Internet connectivity unavailable; Windows Update stage is skipped")
        summary["skipped"] = True
        summary["reason"] = "No internet connectivity"
        end = dt.datetime.now().astimezone()
        ctx.record_stage(stage_name, "SKIPPED", start, end, "Windows Update stage skipped because internet is unavailable.")
        return summary
    install_updates = not ctx.args.quick_mode
    data, result = run_powershell_json(
        ctx,
        stage_name,
        "Search and process Windows Updates",
        update_search_script(install_updates=install_updates),
        timeout=7200,
        depth=12,
    )
    if not result["ok"] or data is None:
        ctx.console.warn("Initial Windows Update workflow failed; attempting controlled component repair")
        repaired, message = repair_update_components(ctx)
        summary["component_repair"] = {"performed": repaired, "message": message}
        if repaired:
            data, result = run_powershell_json(
                ctx,
                stage_name,
                "Retry Windows Updates after repair",
                update_search_script(install_updates=install_updates),
                timeout=7200,
                depth=12,
            )
    if data is None:
        ctx.add_finding("major", "Windows Update workflow failed", "Unable to query or process Windows Updates safely.", "updates")
        ctx.add_unresolved("Windows Update servicing remains unresolved.")
        end = dt.datetime.now().astimezone()
        ctx.record_stage(stage_name, "WARN", start, end, "Windows Update workflow could not be completed.")
        return summary
    updates = flatten_to_list(data.get("Updates"))
    summary.update(data)
    summary["found_titles"] = [item.get("Title") for item in updates[:20]]
    if int(data.get("FoundCount", 0) or 0) == 0:
        ctx.console.ok("No pending software updates were detected")
    else:
        ctx.console.info(f"{data.get('FoundCount')} software update(s) detected")
    install_info = data.get("Install") or {}
    if install_info:
        install_code = parse_update_result_code(install_info.get("ResultCode"))
        ctx.console.info(f"Windows Update install result: {install_code}")
        if int(install_info.get("ResultCode", -1) or -1) in (2, 3) and install_info.get("RebootRequired"):
            ctx.reboot_required = True
        elif int(install_info.get("ResultCode", -1) or -1) >= 4:
            ctx.add_finding("major", "Windows Update installation failures occurred", f"Install result: {install_code}", "updates")
            ctx.add_unresolved("One or more Windows Updates failed to install.")
    elif not install_updates and updates:
        ctx.console.info("Quick mode active; update installation was skipped after detection")
    per_update = flatten_to_list((install_info or {}).get("Updates"))
    failed = [item for item in per_update if int(item.get("ResultCode", 0) or 0) >= 4]
    if failed:
        ctx.add_finding("warning", "Specific Windows updates failed", f"{len(failed)} update(s) returned failure states.", "updates")
        ctx.add_recommendation("Review Windows Update raw captures for specific KB failures and HRESULT codes.")
    if data.get("RebootRequired"):
        ctx.reboot_required = True
    ctx.data["diagnostics"]["windows_update"] = summary
    end = dt.datetime.now().astimezone()
    ctx.record_stage(stage_name, "OK" if not failed else "WARN", start, end, "Windows Update workflow completed.")
    return summary


def registry_app_search_script(pattern: str) -> str:
    safe_pattern = pattern.replace("'", "''")
    return f"""
    $paths = @(
        'HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*',
        'HKLM:\\Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*',
        'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*'
    )
    Get-ItemProperty -Path $paths -ErrorAction SilentlyContinue | Where-Object {{ $_.DisplayName -match '{safe_pattern}' }} |
        Select-Object DisplayName, DisplayVersion, Publisher, InstallLocation, InstallDate
    """


def locate_nvidia_app_executable() -> Optional[Path]:
    candidates = [
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "NVIDIA Corporation" / "NVIDIA app" / "NVIDIAApp.exe",
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "NVIDIA Corporation" / "NVIDIA app" / "CEF" / "NVIDIAApp.exe",
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "NVIDIA Corporation" / "NVIDIA app" / "NVIDIA App.exe",
    ]
    for path in candidates:
        if path.exists():
            return path
    return None


def download_nvidia_app(ctx: ToolkitContext) -> Tuple[bool, str, Optional[str]]:
    page_path = ctx.paths.downloads / "nvidia_app_page.html"
    installer_path = ctx.paths.downloads / "NVIDIA_App_Installer.exe"
    try:
        req = urllib.request.Request(NVIDIA_APP_PAGE, headers={"User-Agent": f"{TOOL_NAME}/{TOOL_VERSION}"})
        with urllib.request.urlopen(req, timeout=DEFAULT_HTTP_TIMEOUT) as response:
            html = response.read().decode("utf-8", errors="replace")
        write_text(page_path, html)
        links = re.findall(r"https://[^\"' ]+\.exe", html, flags=re.IGNORECASE)
        links = [link for link in links if "nvidia.com" in link.lower() or "download.nvidia.com" in link.lower()]
        direct_link = links[0] if links else None
        if direct_link:
            req = urllib.request.Request(direct_link, headers={"User-Agent": f"{TOOL_NAME}/{TOOL_VERSION}"})
            with urllib.request.urlopen(req, timeout=DEFAULT_HTTP_TIMEOUT) as response:
                installer_path.write_bytes(response.read())
            return True, f"NVIDIA App installer downloaded to {installer_path}", str(installer_path)
        webbrowser.open(NVIDIA_APP_PAGE)
        return False, "Official NVIDIA App page opened for review; direct installer link was not discovered in page source", None
    except Exception as exc:
        return False, f"NVIDIA App retrieval failed: {exc}", None


def run_stage_driver_assistance(ctx: ToolkitContext) -> Dict[str, Any]:
    stage_name = "Stage 6 - Driver Assistance"
    start = dt.datetime.now().astimezone()
    ctx.console.section(stage_name)
    summary: Dict[str, Any] = {"skipped": False}
    if ctx.args.skip_driver_stage:
        ctx.console.warn("Driver assistance skipped by flag")
        summary["skipped"] = True
        end = dt.datetime.now().astimezone()
        ctx.record_stage(stage_name, "SKIPPED", start, end, "Driver assistance skipped by flag.")
        return summary
    skip_reason = stage_skip_reason(
        ctx,
        "Report-only mode active; driver assistance is skipped",
        "Driver assistance was skipped because safety blockers are present",
    )
    if skip_reason:
        ctx.console.warn(skip_reason)
        summary["skipped"] = True
        summary["reason"] = skip_reason
        ctx.data["diagnostics"]["driver_assistance"] = summary
        end = dt.datetime.now().astimezone()
        ctx.record_stage(stage_name, "SKIPPED", start, end, skip_reason)
        return summary
    gpus = flatten_to_list(ctx.data.get("machine", {}).get("gpu"))
    nvidia = [gpu for gpu in gpus if "nvidia" in str(gpu.get("Name", "")).lower() or "nvidia" in str(gpu.get("AdapterCompatibility", "")).lower()]
    summary["nvidia_detected"] = bool(nvidia)
    if not nvidia:
        ctx.console.ok("No NVIDIA GPU detected; NVIDIA App assistance skipped")
        end = dt.datetime.now().astimezone()
        ctx.record_stage(stage_name, "SKIPPED", start, end, "No NVIDIA GPU detected.")
        return summary
    install_data, _ = run_powershell_json(ctx, stage_name, "Detect NVIDIA App installation", registry_app_search_script("NVIDIA App"), timeout=60)
    installed_entries = flatten_to_list(install_data)
    summary["installed_entries"] = installed_entries
    executable = locate_nvidia_app_executable()
    if installed_entries:
        ctx.console.ok("NVIDIA App is already installed")
        summary["nvidia_app_installed"] = True
        summary["executable"] = str(executable) if executable else None
        if executable and not ctx.args.non_interactive:
            try:
                subprocess.Popen([str(executable)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                summary["launched"] = True
                ctx.console.ok("NVIDIA App launched")
            except Exception as exc:
                summary["launched"] = False
                summary["launch_error"] = str(exc)
                ctx.console.warn(f"NVIDIA App launch failed: {exc}")
        end = dt.datetime.now().astimezone()
        ctx.record_stage(stage_name, "OK", start, end, "NVIDIA driver assistance completed.")
        ctx.data["diagnostics"]["driver_assistance"] = summary
        return summary
    if not ctx.data.get("machine", {}).get("networking", {}).get("internet_available"):
        ctx.console.warn("NVIDIA GPU detected but internet is unavailable; NVIDIA App download skipped")
        ctx.add_recommendation("Install NVIDIA App from the official NVIDIA source when internet access is available.")
        summary["nvidia_app_installed"] = False
        end = dt.datetime.now().astimezone()
        ctx.record_stage(stage_name, "WARN", start, end, "NVIDIA App could not be downloaded because internet is unavailable.")
        ctx.data["diagnostics"]["driver_assistance"] = summary
        return summary
    downloaded, message, installer_path = download_nvidia_app(ctx)
    summary["downloaded"] = downloaded
    summary["message"] = message
    summary["installer_path"] = installer_path
    if downloaded and installer_path and not ctx.args.non_interactive:
        try:
            subprocess.Popen([installer_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            summary["launched"] = True
            ctx.console.ok("NVIDIA App installer launched")
        except Exception as exc:
            summary["launched"] = False
            summary["launch_error"] = str(exc)
            ctx.console.warn(f"NVIDIA App installer launch failed: {exc}")
    elif not downloaded:
        ctx.console.warn(message)
        ctx.add_recommendation("Review the official NVIDIA App page and perform manual installation if required.")
    else:
        ctx.console.ok(message)
    ctx.data["diagnostics"]["driver_assistance"] = summary
    end = dt.datetime.now().astimezone()
    ctx.record_stage(stage_name, "OK" if downloaded else "WARN", start, end, "NVIDIA driver assistance stage completed.")
    return summary


def run_stage_network_remediation(ctx: ToolkitContext) -> Dict[str, Any]:
    stage_name = "Stage 7 - Optional Networking Remediation"
    start = dt.datetime.now().astimezone()
    ctx.console.section(stage_name)
    summary: Dict[str, Any] = {"performed": False}
    if not (ctx.args.network_reset or ctx.detected_network_issue):
        ctx.console.info("Networking remediation not requested and no clear network fault was detected")
        end = dt.datetime.now().astimezone()
        ctx.record_stage(stage_name, "SKIPPED", start, end, "Network remediation was not required.")
        return summary
    skip_reason = stage_skip_reason(
        ctx,
        "Report-only mode active; networking remediation skipped",
        "Networking remediation was skipped because safety blockers are present",
    )
    if skip_reason:
        ctx.console.warn(skip_reason)
        summary["reason"] = skip_reason
        ctx.data["diagnostics"]["network_remediation"] = summary
        end = dt.datetime.now().astimezone()
        ctx.record_stage(stage_name, "SKIPPED", start, end, skip_reason)
        return summary
    management = ctx.data.get("prechecks", {}).get("management", {})
    summary["performed"] = True
    actions = []
    flush = run_command(ctx, stage_name, "Flush DNS cache", ["ipconfig.exe", "/flushdns"], timeout=45)
    actions.append({"action": "flushdns", "ok": flush["ok"]})
    if management.get("enterprise_managed"):
        ctx.console.warn("Enterprise-managed device detected; Winsock and TCP/IP reset skipped")
        actions.append({"action": "winsock_reset", "ok": False, "skipped": True, "reason": "enterprise-managed"})
        actions.append({"action": "tcpip_reset", "ok": False, "skipped": True, "reason": "enterprise-managed"})
    else:
        winsock = run_command(ctx, stage_name, "Reset Winsock", ["netsh.exe", "winsock", "reset"], timeout=120)
        tcpip = run_command(ctx, stage_name, "Reset TCP/IP stack", ["netsh.exe", "int", "ip", "reset"], timeout=180)
        actions.append({"action": "winsock_reset", "ok": winsock["ok"]})
        actions.append({"action": "tcpip_reset", "ok": tcpip["ok"]})
        if winsock["ok"] or tcpip["ok"]:
            ctx.reboot_required = True
    for service_name in ("Dnscache", "NlaSvc"):
        outcome = restart_service_if_needed(ctx, service_name, "network remediation")
        actions.append({"action": f"service_{service_name}", "result": outcome})
    summary["actions"] = actions
    ctx.data["diagnostics"]["network_remediation"] = summary
    end = dt.datetime.now().astimezone()
    ctx.record_stage(stage_name, "OK", start, end, "Networking remediation stage completed.")
    return summary


def run_stage_post_verification(ctx: ToolkitContext) -> Dict[str, Any]:
    stage_name = "Stage 8 - Post-Repair Verification"
    start = dt.datetime.now().astimezone()
    ctx.console.section(stage_name)
    verification: Dict[str, Any] = {}
    notes: List[str] = []
    stage_status = "OK"
    baseline = get_baseline_snapshot(ctx)
    enrich_machine_context(ctx, baseline)
    verification["pending_reboot_reasons"] = pending_reboot_reasons()
    verification["internet_available"] = ctx.data.get("machine", {}).get("networking", {}).get("internet_available")
    verification["system_drive_free_bytes"] = 0
    verification["system_drive_free_percent"] = 0.0
    system_drive = SYSTEM_DRIVE_ENV.rstrip("\\")
    for volume in flatten_to_list(ctx.data.get("machine", {}).get("logicalDisks")):
        if str(volume.get("DeviceID", "")).upper() == system_drive.upper():
            free = int(volume.get("FreeSpace", 0) or 0)
            size = int(volume.get("Size", 0) or 0)
            verification["system_drive_free_bytes"] = free
            verification["system_drive_free_percent"] = round((free / size) * 100, 2) if size else 0.0
            break
    skip_reason = stage_skip_reason(
        ctx,
        "Report-only mode active; command-based post-verification is skipped",
        "Command-based post-verification was skipped because safety blockers are present",
    )
    if skip_reason:
        verification["command_skip_reason"] = skip_reason
        notes.append(skip_reason)
    else:
        check_health = run_command(
            ctx,
            stage_name,
            "DISM CheckHealth post-verification",
            ["dism.exe", "/Online", "/Cleanup-Image", "/CheckHealth"],
            timeout=1800,
        )
        verification["dism_checkhealth"] = summarize_dism_result("DISM CheckHealth", check_health)
        notes.append(verification["dism_checkhealth"])
        if not check_health["ok"]:
            stage_status = "WARN"
            ctx.add_unresolved("DISM post-verification did not complete cleanly.")
        run_verify, verify_reason = should_run_post_sfc_verify(ctx)
        if run_verify:
            ctx.console.info(f"Post-repair SFC verify-only is running because {verify_reason}.")
            sfc_verify = run_command(
                ctx,
                stage_name,
                "SFC verify-only",
                ["sfc.exe", "/verifyonly"],
                timeout=SFC_VERIFY_TIMEOUT_SECONDS,
                acceptable_exit_codes=(0, 1),
                heartbeat_interval=60,
                progress_probe=make_sfc_progress_probe("verify"),
                progress_probe_interval=30,
            )
            verification["sfc_verify"] = summarize_dism_result("SFC VerifyOnly", sfc_verify)
            verification["sfc_verify_timed_out"] = sfc_verify["timed_out"]
            notes.append(verification["sfc_verify"])
            if not sfc_verify["ok"]:
                stage_status = "WARN"
                ctx.add_unresolved("SFC verify-only did not complete cleanly after repair.")
        else:
            verification["sfc_verify"] = f"Skipped: {verify_reason}"
            notes.append(f"SFC verify-only skipped: {verify_reason}")
    ctx.data["diagnostics"]["post_verification"] = verification
    end = dt.datetime.now().astimezone()
    ctx.record_stage(stage_name, stage_status, start, end, "Post-repair verification completed.", notes=notes)
    return verification


def determine_final_status(ctx: ToolkitContext) -> str:
    if ctx.hardware_suspicions:
        return "Possible hardware issue suspected"
    if ctx.unresolved and ctx.major_issues and not ctx.reboot_required:
        return "Manual follow-up required"
    if ctx.unresolved and ctx.reboot_required:
        return "Partially repaired"
    if ctx.reboot_required and not ctx.unresolved:
        return "Repaired with reboot required"
    if ctx.major_issues and not ctx.unresolved:
        return "Improved"
    if ctx.warnings and not ctx.major_issues and not ctx.unresolved:
        return "Improved"
    if ctx.unresolved:
        return "Unresolved issues remain"
    return "Healthy"


def build_results_summary(ctx: ToolkitContext) -> Dict[str, Any]:
    servicing = ctx.data.get("diagnostics", {}).get("servicing", {})
    updates = ctx.data.get("diagnostics", {}).get("windows_update", {})
    disk = ctx.data.get("diagnostics", {}).get("disk", {})
    network = ctx.data.get("diagnostics", {}).get("network_remediation", {})
    drivers = ctx.data.get("diagnostics", {}).get("driver_assistance", {})
    status = determine_final_status(ctx)
    unresolved = list(dict.fromkeys(ctx.unresolved + ctx.hardware_suspicions))
    summary = {
        "final_status": status,
        "dism_result": servicing.get("DISM RestoreHealth", {}).get("summary", "Not run"),
        "sfc_result": servicing.get("SFC Scannow", {}).get("summary", "Not run"),
        "update_result": parse_update_result_code(((updates.get("Install") or {}).get("ResultCode"))) if updates else "Not run",
        "disk_result": "Issues indicated" if any(item.get("needs_repair") for item in disk.get("volumes", [])) else "No repair indication from CHKDSK scans",
        "network_result": "Remediation performed" if network.get("performed") else "No networking remediation performed",
        "driver_result": "NVIDIA assistance performed" if drivers.get("nvidia_detected") else "No NVIDIA-specific action required",
        "reboot_required": ctx.reboot_required,
        "unresolved_issues": unresolved,
    }
    ctx.data["summary"] = summary
    return summary


def create_text_summary(ctx: ToolkitContext) -> str:
    machine = ctx.data.get("machine", {})
    comp = machine.get("computer", {})
    os_info = machine.get("os", {})
    summary = build_results_summary(ctx)
    lines = [
        TOOL_NAME,
        REPORT_TITLE,
        COPYRIGHT_TEXT,
        "",
        f"Session ID: {ctx.session_id}",
        f"Generated: {iso_now()}",
        f"Hostname: {comp.get('Hostname', '')}",
        f"User: {ctx.data['session'].get('current_user', '')}",
        f"Windows: {os_info.get('Caption', '')} {os_info.get('Version', '')} build {os_info.get('BuildNumber', '')}",
        f"Architecture: {os_info.get('Architecture', '')}",
        f"Local IP: {machine.get('networking', {}).get('lan_ip') or 'Unavailable'}",
        f"Public IP: {machine.get('networking', {}).get('public_ip') or 'Unavailable'}",
        "",
        f"Final Status: {summary['final_status']}",
        f"DISM: {summary['dism_result']}",
        f"SFC: {summary['sfc_result']}",
        f"Updates: {summary['update_result']}",
        f"Disk: {summary['disk_result']}",
        f"Network: {summary['network_result']}",
        f"Driver Assistance: {summary['driver_result']}",
        f"Reboot Required: {'Yes' if summary['reboot_required'] else 'No'}",
        "",
        "Findings:",
    ]
    findings = ctx.data.get("findings", [])
    if findings:
        for finding in findings:
            lines.append(f"- [{finding['severity']}] {finding['title']}: {finding['details']}")
    else:
        lines.append("- No major findings recorded.")
    lines.append("")
    lines.append("Recommendations:")
    if ctx.recommendations:
        for rec in ctx.recommendations:
            lines.append(f"- {rec}")
    else:
        lines.append("- No additional recommendations.")
    lines.append("")
    lines.append(f"Generated by {TOOL_NAME}")
    lines.append(f"Session ID: {ctx.session_id}")
    lines.append(COPYRIGHT_TEXT)
    return "\n".join(lines) + "\n"


def ensure_reportlab(ctx: ToolkitContext) -> bool:
    try:
        import reportlab  # type: ignore # noqa: F401

        return True
    except Exception as exc:
        ctx.logger.info("ReportLab unavailable; using built-in PDF fallback: %s", exc)
        return False


def report_rows_from_actions(ctx: ToolkitContext) -> List[List[str]]:
    rows = [["Stage", "Action", "Status", "Details"]]
    for action in ctx.data.get("actions", []):
        rows.append(
            [
                str(action.get("stage", ""))[:24],
                str(action.get("action", ""))[:38],
                str(action.get("status", "")),
                str(action.get("details", ""))[:64],
            ]
        )
    return rows[:70]


def findings_as_paragraphs(ctx: ToolkitContext) -> List[str]:
    items = []
    for finding in ctx.data.get("findings", []):
        items.append(f"[{finding['severity']}] {finding['title']} - {finding['details']}")
    return items or ["No significant findings were recorded."]


def hardware_lines(ctx: ToolkitContext) -> List[str]:
    machine = ctx.data.get("machine", {})
    lines = []
    for cpu in flatten_to_list(machine.get("cpu")):
        lines.append(f"CPU: {cpu.get('Name')} | Cores: {cpu.get('NumberOfCores')} | Logical: {cpu.get('NumberOfLogicalProcessors')}")
    ram = machine.get("computer", {}).get("TotalPhysicalMemory")
    if ram:
        lines.append(f"RAM: {format_bytes(ram)}")
    for gpu in flatten_to_list(machine.get("gpu")):
        lines.append(f"GPU: {gpu.get('Name')} | Driver: {gpu.get('DriverVersion')} | Status: {gpu.get('Status')}")
    for disk in flatten_to_list(machine.get("disks")):
        lines.append(f"Disk: {disk.get('Model')} | {format_bytes(disk.get('Size'))} | {disk.get('InterfaceType')} | {disk.get('Status')}")
    return lines


def machine_identity_lines(ctx: ToolkitContext) -> List[str]:
    machine = ctx.data.get("machine", {})
    comp = machine.get("computer", {})
    os_info = machine.get("os", {})
    networking = machine.get("networking", {})
    return [
        f"PC Name: {comp.get('Hostname', '')}",
        f"Current User: {ctx.data['session'].get('current_user', '')}",
        f"Windows: {os_info.get('Caption', '')}",
        f"Version / Build: {os_info.get('Version', '')} / {os_info.get('BuildNumber', '')}",
        f"Architecture: {os_info.get('Architecture', '')}",
        f"Uptime: {minutes_to_human(machine.get('uptime_minutes', 0))}",
        f"Local IP: {networking.get('lan_ip') or 'Unavailable'}",
        f"Other Local IPs: {', '.join(networking.get('local_ipv4', []) or []) or 'Unavailable'}",
        f"Public IP: {networking.get('public_ip') or 'Unavailable'}",
        f"Domain / Workgroup: {comp.get('Domain', '')}",
        f"Manufacturer / Model: {comp.get('Manufacturer', '')} / {comp.get('Model', '')}",
        f"Serial: {comp.get('SerialNumber', '')}",
    ]


def draw_minimal_pdf(ctx: ToolkitContext) -> None:
    summary = build_results_summary(ctx)
    sections = [
        ("Machine Identity", machine_identity_lines(ctx)),
        ("Hardware Summary", hardware_lines(ctx)),
        ("Diagnostic Findings", findings_as_paragraphs(ctx)),
        ("Recommendations", ctx.recommendations or ["No additional recommendations."]),
    ]
    pages: List[List[str]] = []
    current_page: List[str] = [
        TOOL_NAME,
        REPORT_TITLE,
        f"Session ID: {ctx.session_id}",
        f"Generated: {iso_now()}",
        COPYRIGHT_TEXT,
        "",
        f"Final Status: {summary['final_status']}",
        "",
    ]
    max_lines = 46
    for section_title, lines in sections:
        block = [section_title, "-" * len(section_title)]
        for line in lines:
            block.extend(textwrap.wrap(str(line), width=92) or [""])
        block.append("")
        if len(current_page) + len(block) > max_lines:
            pages.append(current_page)
            current_page = []
        current_page.extend(block)
    actions_section = ["Actions Performed", "-" * len("Actions Performed")]
    for row in report_rows_from_actions(ctx)[1:40]:
        actions_section.extend(textwrap.wrap(" | ".join(row), width=92) or [""])
    actions_section.append("")
    if len(current_page) + len(actions_section) > max_lines:
        pages.append(current_page)
        current_page = []
    current_page.extend(actions_section)
    if current_page:
        pages.append(current_page)

    def pdf_escape(text: str) -> str:
        return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")

    objects: List[bytes] = []

    def add_object(body: bytes) -> int:
        objects.append(body)
        return len(objects)

    font_id = add_object(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    font_bold_id = add_object(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>")
    page_ids: List[int] = []
    for page_number, page_lines in enumerate(pages, start=1):
        title = page_lines[0] if page_lines else TOOL_NAME
        stream_lines = ["BT", "/F2 18 Tf", "50 780 Td", f"({pdf_escape(title)}) Tj", "0 -24 Td", "/F1 10 Tf"]
        for line in page_lines[1:]:
            stream_lines.append(f"({pdf_escape(line)}) Tj")
            stream_lines.append("0 -14 Td")
        footer = f"Generated by {TOOL_NAME} | Session {ctx.session_id} | Page {page_number} | {COPYRIGHT_TEXT}"
        stream_lines.extend(["ET", "BT", "/F1 9 Tf", "50 30 Td", f"({pdf_escape(footer)}) Tj", "ET"])
        stream = "\n".join(stream_lines).encode("latin-1", errors="replace")
        content_id = add_object(f"<< /Length {len(stream)} >>\nstream\n".encode("ascii") + stream + b"\nendstream")
        page_id = add_object(
            f"<< /Type /Page /Parent 0 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 {font_id} 0 R /F2 {font_bold_id} 0 R >> >> /Contents {content_id} 0 R >>".encode("ascii")
        )
        page_ids.append(page_id)
    kids = " ".join(f"{page_id} 0 R" for page_id in page_ids)
    pages_id = add_object(f"<< /Type /Pages /Kids [{kids}] /Count {len(page_ids)} >>".encode("ascii"))
    catalog_id = add_object(f"<< /Type /Catalog /Pages {pages_id} 0 R >>".encode("ascii"))
    info_id = add_object(
        f"<< /Title ({TOOL_NAME} Report) /Author ({AUTHOR_NAME}) /Subject ({REPORT_TITLE}) /Creator ({TOOL_NAME}) /Producer ({TOOL_NAME}) >>".encode("latin-1", errors="replace")
    )
    fixed_objects: List[bytes] = []
    for index, body in enumerate(objects, start=1):
        if index in page_ids:
            fixed_objects.append(body.replace(b"/Parent 0 0 R", f"/Parent {pages_id} 0 R".encode("ascii")))
        else:
            fixed_objects.append(body)
    pdf = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]
    for index, body in enumerate(fixed_objects, start=1):
        offsets.append(len(pdf))
        pdf.extend(f"{index} 0 obj\n".encode("ascii"))
        pdf.extend(body)
        pdf.extend(b"\nendobj\n")
    xref_offset = len(pdf)
    pdf.extend(f"xref\n0 {len(fixed_objects) + 1}\n".encode("ascii"))
    pdf.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        pdf.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    pdf.extend(f"trailer\n<< /Size {len(fixed_objects) + 1} /Root {catalog_id} 0 R /Info {info_id} 0 R >>\n".encode("ascii"))
    pdf.extend(f"startxref\n{xref_offset}\n%%EOF".encode("ascii"))
    ctx.paths.pdf_report.write_bytes(pdf)


def build_reportlab_pdf(ctx: ToolkitContext) -> None:
    from reportlab.lib import colors  # type: ignore
    from reportlab.lib.pagesizes import A4  # type: ignore
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet  # type: ignore
    from reportlab.lib.units import mm  # type: ignore
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle  # type: ignore

    summary = build_results_summary(ctx)
    doc = SimpleDocTemplate(str(ctx.paths.pdf_report), pagesize=A4, rightMargin=16 * mm, leftMargin=16 * mm, topMargin=18 * mm, bottomMargin=18 * mm)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="SectionTitle", fontName="Helvetica-Bold", fontSize=12, leading=14, textColor=colors.HexColor("#17365D"), spaceAfter=6, spaceBefore=8))
    styles.add(ParagraphStyle(name="BodySmall", fontName="Helvetica", fontSize=9, leading=12))
    story: List[Any] = []
    story.append(Paragraph(f"<b>{TOOL_NAME}</b>", styles["Title"]))
    story.append(Paragraph(REPORT_TITLE, styles["Heading2"]))
    story.append(Paragraph(COPYRIGHT_TEXT, styles["BodySmall"]))
    story.append(Spacer(1, 6))
    status_color = {
        "Healthy": "#2E7D32",
        "Improved": "#1565C0",
        "Repaired with reboot required": "#EF6C00",
        "Partially repaired": "#EF6C00",
        "Unresolved issues remain": "#C62828",
        "Manual follow-up required": "#C62828",
        "Possible hardware issue suspected": "#8E0000",
    }.get(summary["final_status"], "#17365D")
    status_table = Table([[Paragraph(f"<font color='white'><b>Final Machine Status: {summary['final_status']}</b></font>", styles["BodyText"])]], colWidths=[175 * mm])
    status_table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(status_color)), ("BOX", (0, 0), (-1, -1), 0.5, colors.black), ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8), ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    story.append(status_table)
    story.append(Spacer(1, 8))
    metadata_rows = [
        ["Session ID", ctx.session_id, "Date/Time", iso_now()],
        ["Tool", TOOL_NAME, "Version", TOOL_VERSION],
        ["Generated By", TOOL_NAME, "Copyright", COPYRIGHT_TEXT],
    ]
    meta = Table(metadata_rows, colWidths=[28 * mm, 60 * mm, 28 * mm, 59 * mm])
    meta.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.whitesmoke), ("GRID", (0, 0), (-1, -1), 0.4, colors.grey), ("FONTNAME", (0, 0), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 0), (-1, -1), 8), ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"), ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold")]))
    story.append(meta)

    def add_list_section(title: str, lines: Iterable[str]) -> None:
        story.append(Spacer(1, 8))
        story.append(Paragraph(title, styles["SectionTitle"]))
        line_items = list(lines) or ["No data recorded."]
        for item in line_items:
            story.append(Paragraph(textwrap.shorten(str(item), width=220, placeholder="..."), styles["BodySmall"]))

    add_list_section("Machine Identity", machine_identity_lines(ctx))
    add_list_section("Hardware Summary", hardware_lines(ctx))
    add_list_section("Diagnostic Findings", findings_as_paragraphs(ctx))
    add_list_section("Recommendations", ctx.recommendations or ["No additional recommendations."])
    story.append(PageBreak())
    story.append(Paragraph("Actions Performed", styles["SectionTitle"]))
    action_rows = report_rows_from_actions(ctx)
    action_table = Table(action_rows, repeatRows=1, colWidths=[28 * mm, 57 * mm, 18 * mm, 72 * mm])
    action_table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#17365D")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 7), ("GRID", (0, 0), (-1, -1), 0.3, colors.grey), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey])]))
    story.append(action_table)
    story.append(Spacer(1, 10))
    story.append(Paragraph("Results Summary", styles["SectionTitle"]))
    results_rows = [
        ["DISM", summary["dism_result"]],
        ["SFC", summary["sfc_result"]],
        ["Windows Update", summary["update_result"]],
        ["Disk / File System", summary["disk_result"]],
        ["Network", summary["network_result"]],
        ["Driver Assistance", summary["driver_result"]],
        ["Reboot Required", "Yes" if summary["reboot_required"] else "No"],
        ["Unresolved Issues", "; ".join(summary["unresolved_issues"]) if summary["unresolved_issues"] else "None recorded"],
    ]
    results_table = Table(results_rows, colWidths=[38 * mm, 137 * mm])
    results_table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.4, colors.grey), ("BACKGROUND", (0, 0), (0, -1), colors.whitesmoke), ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 8), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(results_table)

    def on_page(canvas, document) -> None:  # type: ignore
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.grey)
        canvas.drawString(16 * mm, 10 * mm, f"Generated by {TOOL_NAME}")
        canvas.drawRightString(195 * mm, 10 * mm, f"Session {ctx.session_id} | Page {document.page}")
        canvas.drawCentredString(105 * mm, 10 * mm, COPYRIGHT_TEXT)
        canvas.restoreState()

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)


def generate_reports(ctx: ToolkitContext) -> None:
    ctx.console.section("Reporting")
    write_text(ctx.paths.txt_summary, create_text_summary(ctx))
    save_json(ctx.paths.json_report, ctx.data)
    if ensure_reportlab(ctx):
        try:
            build_reportlab_pdf(ctx)
            ctx.console.ok(f"PDF report generated: {ctx.paths.pdf_report}")
            return
        except Exception as exc:
            ctx.console.warn(f"ReportLab PDF generation failed: {exc}; using built-in PDF fallback")
            ctx.logger.exception("ReportLab PDF generation failed")
    else:
        ctx.console.info("ReportLab is not installed in this Python environment; using the built-in PDF fallback.")
    draw_minimal_pdf(ctx)
    ctx.console.ok(f"PDF report generated with fallback renderer: {ctx.paths.pdf_report}")


def maybe_reboot(ctx: ToolkitContext) -> None:
    if not ctx.reboot_required:
        return
    if not repairs_allowed(ctx):
        return
    ctx.add_recommendation("Reboot the machine to complete pending repairs or update finalization.")
    if not ctx.args.allow_reboot:
        return
    save_resume_state(ctx, resume_from="post_reboot_verification")
    created, message = create_resume_task(ctx)
    save_resume_state(ctx, resume_from="post_reboot_verification")
    manual_resume = ctx.data.get("resume", {}).get("manual_resume_command")
    if manual_resume:
        ctx.add_recommendation(f"If automatic resume fails, run: {manual_resume}")
    if created:
        ctx.console.ok(message)
    else:
        ctx.console.warn(message)
        if manual_resume:
            ctx.console.warn(f"Manual resume command: {manual_resume}")
            ctx.add_recommendation(f"If automatic resume fails, run: {manual_resume}")
    ctx.console.warn("Reboot is required to complete scheduled or pending repair work")
    if not created and ctx.args.non_interactive:
        ctx.console.warn("Automatic reboot skipped because resume task creation failed in non-interactive mode")
        return
    if ctx.args.non_interactive:
        run_command(ctx, "Resume", "Initiate system reboot", ["shutdown.exe", "/r", "/t", "20", "/c", f"{TOOL_NAME} scheduled resume"], timeout=20)
        sys.exit(194)
    response = input("Reboot now and resume automatically? [Y/N]: ").strip().lower()
    if response in ("y", "yes"):
        run_command(ctx, "Resume", "Initiate system reboot", ["shutdown.exe", "/r", "/t", "20", "/c", f"{TOOL_NAME} scheduled resume"], timeout=20)
        sys.exit(194)


def load_context_from_state(args: argparse.Namespace, state_path: Path) -> ToolkitContext:
    state = load_resume_state(state_path)
    session_id = state["session_id"]
    session_paths = determine_paths(session_id, existing_base=Path(state["base_dir"]))
    logger, log_file = setup_logging(session_paths.base)
    session_paths = SessionPaths(
        base=session_paths.base,
        logs=session_paths.logs,
        temp=session_paths.temp,
        downloads=session_paths.downloads,
        reports=session_paths.reports,
        raw=session_paths.raw,
        state_file=session_paths.state_file,
        json_report=session_paths.json_report,
        txt_summary=session_paths.txt_summary,
        pdf_report=session_paths.pdf_report,
        log_file=log_file,
    )
    ctx = ToolkitContext(args, session_paths, logger)
    ctx.data = state.get("data", ctx.data)
    ctx.resume_loaded = True
    ctx.reboot_required = bool(state.get("reboot_required"))
    ctx.resume_task_name = state.get("resume_task_name")
    ctx.recommendations = list(ctx.data.get("recommendations", []))
    ctx.unresolved = list(ctx.data.get("unresolved_issues", []))
    ctx.hardware_suspicions = list(ctx.data.get("hardware_suspicions", []))
    for finding in ctx.data.get("findings", []):
        severity = str(finding.get("severity", "")).lower()
        title = str(finding.get("title", ""))
        if severity == "major" and title not in ctx.major_issues:
            ctx.major_issues.append(title)
        elif severity == "warning" and title not in ctx.warnings:
            ctx.warnings.append(title)
    ctx.data["session"]["resumed_at"] = iso_now()
    ctx.console.banner(ctx.session_id)
    ctx.console.info("Resumed from saved session state")
    return ctx


def initialize_context(args: argparse.Namespace) -> ToolkitContext:
    session_id = f"{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8].upper()}"
    session_paths = determine_paths(session_id)
    logger, log_file = setup_logging(session_paths.base)
    session_paths = SessionPaths(
        base=session_paths.base,
        logs=session_paths.logs,
        temp=session_paths.temp,
        downloads=session_paths.downloads,
        reports=session_paths.reports,
        raw=session_paths.raw,
        state_file=session_paths.state_file,
        json_report=session_paths.json_report,
        txt_summary=session_paths.txt_summary,
        pdf_report=session_paths.pdf_report,
        log_file=log_file,
    )
    ctx = ToolkitContext(args, session_paths, logger)
    ctx.console.banner(ctx.session_id)
    return ctx


def print_final_panel(ctx: ToolkitContext) -> None:
    summary = build_results_summary(ctx)
    ctx.console.section("Run Summary")
    ctx.console.info(f"Final machine status: {summary['final_status']}")
    ctx.console.info(f"Reboot required: {'Yes' if summary['reboot_required'] else 'No'}")
    if summary["unresolved_issues"]:
        ctx.console.warn("Unresolved issues:")
        for item in summary["unresolved_issues"]:
            ctx.console.warn(f" - {item}")
    else:
        ctx.console.ok("No unresolved issues were recorded")
    ctx.console.info(f"JSON report: {ctx.paths.json_report}")
    ctx.console.info(f"TXT summary: {ctx.paths.txt_summary}")
    ctx.console.info(f"PDF report: {ctx.paths.pdf_report}")


def run_full_workflow(ctx: ToolkitContext) -> None:
    run_stage_baseline(ctx)
    perform_safety_prechecks(ctx)
    run_stage_servicing(ctx)
    run_stage_disk_health(ctx)
    run_stage_cleanup(ctx)
    run_stage_windows_update(ctx)
    run_stage_driver_assistance(ctx)
    run_stage_network_remediation(ctx)
    if ctx.reboot_required and ctx.args.allow_reboot and not ctx.resume_loaded and repairs_allowed(ctx):
        generate_reports(ctx)
        print_final_panel(ctx)
        maybe_reboot(ctx)
    run_stage_post_verification(ctx)
    generate_reports(ctx)
    print_final_panel(ctx)


def run_resumed_workflow(ctx: ToolkitContext) -> None:
    run_stage_post_verification(ctx)
    generate_reports(ctx)
    print_final_panel(ctx)
    remove_resume_task(ctx)
    delete_if_exists(ctx.paths.state_file)


def main() -> int:
    ensure_windows()
    parser = build_arg_parser()
    args = parser.parse_args()
    if args.resume_state:
        ctx = load_context_from_state(args, Path(args.resume_state))
        run_resumed_workflow(ctx)
        return 0
    if not is_admin():
        try:
            relaunch_as_admin()
            return 0
        except Exception as exc:
            print(f"[FAIL  ] Administrator rights are required. {exc}")
            return 1
    ctx = initialize_context(args)
    try:
        run_full_workflow(ctx)
        return 0
    except KeyboardInterrupt:
        ctx.console.warn("Execution interrupted by user")
        try:
            generate_reports(ctx)
            print_final_panel(ctx)
        except Exception:
            pass
        return 2
    except Exception:
        ctx.logger.exception("Fatal unhandled exception")
        ctx.console.fail("A fatal error occurred. Capturing reports with available data.")
        ctx.add_unresolved("The tool encountered an unexpected exception; review the session log.")
        try:
            generate_reports(ctx)
            print_final_panel(ctx)
        except Exception:
            ctx.logger.exception("Failed to generate fallback reports after fatal error")
        return 3


if __name__ == "__main__":
    sys.exit(main())
