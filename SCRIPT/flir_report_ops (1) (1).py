"""
FLIR Word report orchestration.

This module keeps manifest validation in Python and delegates the actual FLIR
report creation to a 32-bit PowerShell helper that automates Word through
FLIR's WordApp integration.
"""

from __future__ import annotations

import json
import os
import subprocess
import tempfile
from dataclasses import dataclass, field
from pathlib import Path


POWERSHELL_32_PATH = Path(r"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe")
HELPER_SCRIPT_PATH = Path(__file__).resolve().parent.parent / "scripts" / "flir_word_report_helper.ps1"


@dataclass(frozen=True)
class FlirReportPage:
    """One output report page."""

    template_page_no: int
    ir_path: Path
    visual_path: Path
    field_overrides: dict[str, str] = field(default_factory=dict)


@dataclass(frozen=True)
class FlirReportManifest:
    """Manifest used to generate a FLIR DOCX report."""

    template_path: Path
    output_path: Path
    pages: list[FlirReportPage]


def load_flir_report_manifest(manifest_path: str | os.PathLike[str]) -> FlirReportManifest:
    """Load a JSON manifest for FLIR report generation."""

    manifest_file = Path(manifest_path)
    payload = json.loads(manifest_file.read_text(encoding="utf-8"))

    manifest = FlirReportManifest(
        template_path=Path(payload["template_path"]).expanduser(),
        output_path=Path(payload["output_path"]).expanduser(),
        pages=[
            FlirReportPage(
                template_page_no=int(page["template_page_no"]),
                ir_path=Path(page["ir_path"]).expanduser(),
                visual_path=Path(page["visual_path"]).expanduser(),
                field_overrides={
                    str(key): str(value)
                    for key, value in page.get("field_overrides", {}).items()
                },
            )
            for page in payload.get("pages", [])
        ],
    )
    validate_flir_report_manifest(manifest)
    return manifest


def validate_flir_report_manifest(manifest: FlirReportManifest) -> None:
    """Validate manifest paths and page structure."""

    if not manifest.template_path.exists():
        raise FileNotFoundError(f"Template file not found: {manifest.template_path}")

    if manifest.template_path.suffix.lower() != ".docx":
        raise ValueError(f"Template must be a .docx file: {manifest.template_path}")

    if not manifest.pages:
        raise ValueError("Manifest must include at least one page.")

    for idx, page in enumerate(manifest.pages, start=1):
        if page.template_page_no < 1:
            raise ValueError(f"Page {idx} has invalid template_page_no={page.template_page_no}")
        if not page.ir_path.exists():
            raise FileNotFoundError(f"Page {idx} IR file not found: {page.ir_path}")
        if not page.visual_path.exists():
            raise FileNotFoundError(f"Page {idx} visual file not found: {page.visual_path}")
        if page.visual_path.suffix.lower() not in {".jpg", ".jpeg"}:
            raise ValueError(f"Page {idx} visual image must be a JPEG: {page.visual_path}")
        if not is_flir_radiometric_jpeg(page.ir_path):
            raise ValueError(
                f"Page {idx} IR file is not a detectable FLIR radiometric JPEG: {page.ir_path}"
            )


def is_flir_radiometric_jpeg(path: str | os.PathLike[str]) -> bool:
    """Detect FLIR radiometric JPEGs using FLIR and FFF markers."""

    file_path = Path(path)
    data = file_path.read_bytes()
    if not data.startswith(b"\xff\xd8\xff"):
        return False
    return b"FLIR" in data[:1024] and b"FFF" in data


def generate_flir_report_from_manifest(manifest_path: str | os.PathLike[str]) -> Path:
    """Generate a FLIR DOCX report from a JSON manifest."""

    manifest = load_flir_report_manifest(manifest_path)
    return generate_flir_report(manifest)


def generate_flir_report(manifest: FlirReportManifest) -> Path:
    """Generate a FLIR report through the PowerShell/Word/FLIR helper."""

    validate_flir_report_manifest(manifest)
    _ensure_windows_prerequisites()

    output_path = manifest.output_path.expanduser().resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    normalized_payload = _build_helper_payload(manifest)
    helper_manifest_path = _write_helper_manifest(normalized_payload)
    try:
        _run_flir_word_helper(helper_manifest_path)
    finally:
        helper_manifest_path.unlink(missing_ok=True)

    if not output_path.exists():
        raise RuntimeError(
            "FLIR helper reported success, but the output document was not created: "
            f"{output_path}"
        )
    return output_path


def _build_helper_payload(manifest: FlirReportManifest) -> dict[str, object]:
    return {
        "template_path": str(manifest.template_path.expanduser().resolve()),
        "output_path": str(manifest.output_path.expanduser().resolve()),
        "pages": [
            {
                "template_page_no": page.template_page_no,
                "ir_path": str(page.ir_path.expanduser().resolve()),
                "visual_path": str(page.visual_path.expanduser().resolve()),
                "field_overrides": dict(page.field_overrides),
            }
            for page in manifest.pages
        ],
    }


def _write_helper_manifest(payload: dict[str, object]) -> Path:
    fd, raw_path = tempfile.mkstemp(prefix="tnb-flir-helper-", suffix=".json")
    helper_manifest_path = Path(raw_path)
    os.close(fd)
    helper_manifest_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return helper_manifest_path


def _ensure_windows_prerequisites() -> None:
    if os.name != "nt":
        raise OSError("FLIR interactive report generation is only supported on Windows.")
    if not HELPER_SCRIPT_PATH.exists():
        raise FileNotFoundError(f"FLIR helper script not found: {HELPER_SCRIPT_PATH}")
    if not POWERSHELL_32_PATH.exists():
        raise FileNotFoundError(
            "32-bit PowerShell was not found at the expected path: "
            f"{POWERSHELL_32_PATH}"
        )


def _run_flir_word_helper(helper_manifest_path: Path) -> str:
    completed = subprocess.run(
        [
            str(POWERSHELL_32_PATH),
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(HELPER_SCRIPT_PATH),
            "-ManifestPath",
            str(helper_manifest_path),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )

    stdout = completed.stdout.strip()
    stderr = completed.stderr.strip()
    if completed.returncode != 0:
        message_lines = [
            "FLIR Word helper failed.",
            f"Helper manifest: {helper_manifest_path}",
        ]
        if stdout:
            message_lines.append("stdout:")
            message_lines.append(stdout)
        if stderr:
            message_lines.append("stderr:")
            message_lines.append(stderr)
        raise RuntimeError("\n".join(message_lines))

    return stdout
