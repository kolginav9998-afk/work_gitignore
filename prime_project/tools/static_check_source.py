#!/usr/bin/env python3
"""
prime-static CI check: verifies the Basic SOURCE FILES themselves (not a built ODS - see
tests/static_checks.py for that) before any build is attempted. Fast, no LibreOffice needed.

Checks:
- all 15 required PRIME_*.bas modules exist under src/basic/
- every module starts with "Option Explicit"
- no PRIME module references legacy WMSDB_*/WMS_*.bas-style identifiers (the old Firebird
  backend) - legacy_reference/ is exempt, it is archived source, not part of the runtime
- every button target and sheet-event target used by tools/build_ods.py's BUTTON_MAP /
  SHEET_EVENT_HANDLERS resolves to a Sub/Function that actually exists in src/basic/
- required top-level files exist (template, docs)
"""
import re
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_BASIC = REPO_ROOT / "src" / "basic"

REQUIRED_MODULES = [
    "PRIME_00_Config", "PRIME_01_Runtime", "PRIME_02_Store", "PRIME_03_Catalog",
    "PRIME_04_Posting", "PRIME_05_Orders", "PRIME_06_Issues", "PRIME_07_Workflows",
    "PRIME_08_ReturnsInventory", "PRIME_09_StockSearch", "PRIME_10_ActsReports",
    "PRIME_11_Kits", "PRIME_12_UI", "PRIME_13_Diagnostics", "PRIME_14_MigrationInstaller",
    "PRIME_15_Transfers", "PRIME_16_Journal",
]

FORBIDDEN_IN_RUNTIME = ["WMSDB", "WMSDBST", "WMSDBO", "WMSDBIU", "WMSDBR", "WMSDBS", "WMSARCH"]

failures = []
passed = []


def check(name, condition, detail=""):
    if condition:
        passed.append(name)
    else:
        failures.append(f"{name}: {detail}")


def main():
    for mod in REQUIRED_MODULES:
        path = SRC_BASIC / f"{mod}.bas"
        check(f"module file exists: {mod}", path.exists(), str(path))

    sub_function_names = set()
    for path in sorted(SRC_BASIC.glob("*.bas")):
        text = path.read_text(encoding="utf-8")
        check(f"Option Explicit: {path.name}", text.lstrip().startswith("Option Explicit"),
              "file must start with 'Option Explicit'")

        # Only real CODE lines matter here - a comment explaining what the legacy 1.4.1 code
        # used to do (for rationale/traceability) is documentation, not a runtime dependency.
        code_only = "\n".join(
            line for line in text.splitlines() if not line.strip().startswith("'")
        )
        for bad in FORBIDDEN_IN_RUNTIME:
            check(f"no legacy '{bad}' reference: {path.name}", bad not in code_only,
                  f"found forbidden legacy identifier '{bad}' in runtime CODE (not a comment)")

        for m in re.finditer(r'^(Public|Private)\s+(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)', text, re.MULTILINE):
            sub_function_names.add(m.group(3))

    # Cross-check the button/event map in tools/build_ods.py against actual declared Subs.
    build_ods_source = (REPO_ROOT / "tools" / "build_ods.py").read_text(encoding="utf-8")
    referenced_subs = set(re.findall(r'"Standard\.PRIME_\d\d_\w+\.(\w+)"', build_ods_source))
    for sub_name in sorted(referenced_subs):
        check(f"button/event target exists: {sub_name}", sub_name in sub_function_names,
              f"tools/build_ods.py references {sub_name} but no PRIME_*.bas module declares it")

    required_top_level = [
        REPO_ROOT / "src" / "templates" / "POKATAK_WMS_1.4.1_template.ods",
        REPO_ROOT / "tools" / "build_ods.py",
        REPO_ROOT / "tests" / "static_checks.py",
        REPO_ROOT / "tests" / "model_tests.py",
        REPO_ROOT / "docs" / "ARCHITECTURE.md",
    ]
    for p in required_top_level:
        check(f"required file exists: {p.relative_to(REPO_ROOT)}", p.exists(), str(p))

    print(f"PASSED: {len(passed)}")
    print(f"FAILED: {len(failures)}")
    for f in failures:
        print(f"  [FAIL] {f}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
