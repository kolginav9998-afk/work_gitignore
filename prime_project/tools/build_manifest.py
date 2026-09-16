#!/usr/bin/env python3
"""
Генерирует build_manifest.json для собранного ПОКАТАК_PRIME.ods по схеме, требуемой
github_engineering-заданием: version, git_commit_sha, git_branch, build_timestamp,
source_modules, source_hashes, template_hash, test_summary, libreoffice_test_version.

Не запускает LibreOffice сам - берёт версию LibreOffice из `soffice --version`, если доступна,
иначе оставляет поле "unknown" (честно, не выдумывая значение).
"""
import argparse
import hashlib
import json
import re
import subprocess
from datetime import datetime, timezone
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_BASIC = REPO_ROOT / "src" / "basic"
TEMPLATE = REPO_ROOT / "src" / "templates" / "POKATAK_WMS_1.4.1_template.ods"


def sha256_of(path: Path) -> str:
    h = hashlib.sha256()
    h.update(path.read_bytes())
    return h.hexdigest()


def git(*args: str) -> str:
    try:
        return subprocess.check_output(["git", *args], cwd=REPO_ROOT, text=True).strip()
    except Exception:
        return "unknown"


def schema_version() -> str:
    config = (SRC_BASIC / "PRIME_00_Config.bas").read_text(encoding="utf-8")
    m = re.search(r'PRIME_SCHEMA_VERSION\s+As\s+String\s*=\s*"([^"]+)"', config)
    return m.group(1) if m else "unknown"


def libreoffice_version() -> str:
    try:
        out = subprocess.check_output(["soffice", "--version"], text=True, timeout=30)
        return out.strip()
    except Exception:
        return "unknown"


def main():
    parser = argparse.ArgumentParser(description="Generate build_manifest.json for a built PRIME .ods")
    parser.add_argument("--ods", type=Path, required=True, help="path to the built .ods")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--test-summary", type=Path, default=None,
                         help="optional path to a model_tests.py text report to embed a pass/fail count")
    args = parser.parse_args()

    modules = sorted(p.name for p in SRC_BASIC.glob("*.bas"))
    source_hashes = {name: sha256_of(SRC_BASIC / name) for name in modules}

    test_summary = {"status": "not_provided"}
    if args.test_summary and args.test_summary.exists():
        text = args.test_summary.read_text(encoding="utf-8")
        m = re.search(r"PASSED:\s*(\d+)/(\d+)", text)
        if m:
            test_summary = {"passed": int(m.group(1)), "total": int(m.group(2))}

    manifest = {
        "version": schema_version(),
        "git_commit_sha": git("rev-parse", "HEAD"),
        "git_branch": git("rev-parse", "--abbrev-ref", "HEAD"),
        "build_timestamp": datetime.now(timezone.utc).isoformat(),
        "source_modules": modules,
        "source_hashes": source_hashes,
        "template_hash": sha256_of(TEMPLATE) if TEMPLATE.exists() else "unknown",
        "output_ods": str(args.ods),
        "output_ods_hash": sha256_of(args.ods) if args.ods.exists() else "unknown",
        "test_summary": test_summary,
        "libreoffice_test_version": libreoffice_version(),
    }

    args.output.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
