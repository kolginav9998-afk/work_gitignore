#!/usr/bin/env python3
"""
Конвертирует текстовый вывод tests/static_checks.py (строки "PASSED: N", "FAILED: N",
"  [FAIL] ...") в validation_report.json для публикации как артефакт PR (build_manifest
requirement из github_engineering-задания).
"""
import json
import re
import sys
from pathlib import Path


def parse(text: str) -> dict:
    passed = re.search(r"^PASSED:\s*(\d+)", text, re.MULTILINE)
    failed = re.search(r"^FAILED:\s*(\d+)", text, re.MULTILINE)
    failures = re.findall(r"^\s*\[FAIL\]\s*(.+)$", text, re.MULTILINE)
    return {
        "passed": int(passed.group(1)) if passed else 0,
        "failed": int(failed.group(1)) if failed else 0,
        "failures": failures,
        "ok": bool(failed) and int(failed.group(1)) == 0,
    }


def main():
    if len(sys.argv) != 3:
        print("usage: validation_report_to_json.py <input.txt> <output.json>")
        return 1
    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    report = parse(input_path.read_text(encoding="utf-8"))
    output_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {output_path}: passed={report['passed']} failed={report['failed']}")
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
