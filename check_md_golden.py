#!/usr/bin/env python3
"""Check accepted Markdown outputs against the Markdown golden directory."""

from __future__ import annotations

import hashlib
import os
import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
DOC_DIR = ROOT / ".codex" / "design-doc-analysis"
OUTPUT_ROOT = DOC_DIR / "all-design-md-v3"
GOLDEN_ROOT = DOC_DIR / "md-golden"
MANIFEST = GOLDEN_ROOT / "manifest.txt"
MD_LIST = DOC_DIR / "md-list.md"
GENERATOR = DOC_DIR / "generate_design_md.py"
OUTPUT_SUMMARY = OUTPUT_ROOT / "summary.md"


def configure_stdio() -> None:
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")


def file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def relative_files(root: Path) -> list[Path]:
    return sorted(path.relative_to(root) for path in root.rglob("*") if path.is_file())


def safe_dir_name(name: str) -> str:
    stem = Path(name).stem
    return re.sub(r'[<>:"/\\|?*【】]+', "_", stem).strip(" _")


def md_list_section(text: str, name: str) -> str:
    match = re.search(rf"^## {re.escape(name)}\s*$(.*?)(?=^## |\Z)", text, re.M | re.S)
    return match.group(1) if match else ""


def md_list_items(text: str, section_name: str) -> list[str]:
    section = md_list_section(text, section_name)
    return re.findall(r"^- `([^`]+)`", section, re.M)


def check_md_list_consistency(manifest_items: list[str]) -> list[str]:
    if not MD_LIST.exists():
        return [f"missing md-list: {MD_LIST}"]
    text = MD_LIST.read_text(encoding="utf-8")
    protected = [safe_dir_name(item) for item in md_list_items(text, "已验收需保护")]
    unprotected = [safe_dir_name(item) for item in md_list_items(text, "未纳入保护")]
    pending = [safe_dir_name(item) for item in md_list_items(text, "待继续验证")]
    focus_paths = md_list_items(text, "当前重点样例")

    failures: list[str] = []
    if sorted(protected) != sorted(manifest_items):
        missing_in_list = sorted(set(manifest_items) - set(protected))
        missing_in_manifest = sorted(set(protected) - set(manifest_items))
        if missing_in_list:
            failures.append(f"md-list missing protected items {', '.join(missing_in_list)}")
        if missing_in_manifest:
            failures.append(f"manifest missing md-list protected items {', '.join(missing_in_manifest)}")

    protected_set = set(manifest_items)
    protected_unprotected = sorted(protected_set & set(unprotected))
    if protected_unprotected:
        failures.append(f"protected items also listed as unprotected {', '.join(protected_unprotected)}")
    protected_pending = sorted(protected_set & set(pending))
    if protected_pending:
        failures.append(f"protected items also listed as pending {', '.join(protected_pending)}")

    for focus_path in focus_paths:
        path = Path(focus_path)
        if not (DOC_DIR / path).exists():
            failures.append(f"focus sample missing {focus_path}")
        if len(path.parts) >= 2 and path.parts[1] in protected_set:
            failures.append(f"focus sample points to protected item {focus_path}")
    return failures


def check_output_is_regenerated() -> list[str]:
    failures: list[str] = []
    if not OUTPUT_SUMMARY.exists():
        return [f"missing output summary: {OUTPUT_SUMMARY}"]

    output_time = OUTPUT_SUMMARY.stat().st_mtime
    dependencies = [GENERATOR, MD_LIST, MANIFEST]
    for dependency in dependencies:
        if not dependency.exists():
            failures.append(f"missing generation dependency: {dependency}")
            continue
        if output_time < dependency.stat().st_mtime:
            failures.append(
                f"all-design-md-v3 is older than {dependency.relative_to(DOC_DIR)}; "
                "rerun python .codex/design-doc-analysis/generate_design_md.py"
            )
    return failures


def main() -> int:
    if not MANIFEST.exists():
        print(f"missing manifest: {MANIFEST}", file=sys.stderr)
        return 1
    items = [line.strip() for line in MANIFEST.read_text(encoding="utf-8").splitlines() if line.strip()]
    failures: list[str] = check_md_list_consistency(items)
    failures.extend(check_output_is_regenerated())
    for item in items:
        golden_dir = GOLDEN_ROOT / item
        output_dir = OUTPUT_ROOT / item
        if not golden_dir.exists():
            failures.append(f"{item}: missing golden directory")
            continue
        if not output_dir.exists():
            failures.append(f"{item}: missing output directory")
            continue
        golden_files = relative_files(golden_dir)
        output_files = relative_files(output_dir)
        if golden_files != output_files:
            missing = sorted(set(golden_files) - set(output_files))
            extra = sorted(set(output_files) - set(golden_files))
            if missing:
                failures.append(f"{item}: missing files {', '.join(str(path) for path in missing)}")
            if extra:
                failures.append(f"{item}: extra files {', '.join(str(path) for path in extra)}")
            continue
        for rel_path in golden_files:
            golden_hash = file_hash(golden_dir / rel_path)
            output_hash = file_hash(output_dir / rel_path)
            if golden_hash != output_hash:
                failures.append(f"{item}: changed {rel_path}")
    if failures:
        print("Markdown golden check failed:")
        for failure in failures:
            print(f"- {failure}")
        return 1
    print(f"Markdown golden check passed: {len(items)} items")
    return 0


if __name__ == "__main__":
    configure_stdio()
    raise SystemExit(main())
