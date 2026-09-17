#!/usr/bin/env python3
"""
Direct-function-invocation runtime verification of the full PRIME 2.1.0 order-to-inventory
workflow in a REAL headless LibreOffice instance (not a Python model, not a pure static check).

This is the strongest evidence class available in this sandboxed environment (no interactive
display server - see docs/KNOWN_ISSUES.md item 13): it opens the actual built .ods, injects a
throwaway diagnostic Basic module (tools/e2e_diagnostic_module.bas, NOT part of the shipped
product) that builds PrimeDocPlan values programmatically and calls the real PRIME_PostDocument/
PRIME_LocationContourBalance engine functions directly - bypassing CurrentSelection/
CurrentController.select(), which are independently documented as unreliable via external
invoke() in this environment. Results are written to a log file (exceptions do not propagate
through invoke() here - an established quirk of this environment) and printed.

Usage: python3 tools/run_e2e_diagnostic.py <path-to-built-ods>
"""
import shutil
import subprocess
import sys
import time
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_BASIC_DIR = REPO_ROOT / "src" / "basic"
DIAGNOSTIC_MODULE = REPO_ROOT / "tools" / "e2e_diagnostic_module.bas"
PORT = 2094
PROFILE_DIR = Path("/tmp/prime_e2e_diag_profile")
LOG_PATH = Path("/tmp/prime_e2e_diagnostic.log")

PRIME_MODULES = [
    "PRIME_00_Config", "PRIME_01_Runtime", "PRIME_02_Store", "PRIME_03_Catalog",
    "PRIME_04_Posting", "PRIME_05_Orders", "PRIME_06_Issues", "PRIME_07_Workflows",
    "PRIME_08_ReturnsInventory", "PRIME_09_StockSearch", "PRIME_10_ActsReports",
    "PRIME_11_Kits", "PRIME_12_UI", "PRIME_13_Diagnostics", "PRIME_14_MigrationInstaller",
    "PRIME_15_Transfers", "PRIME_16_Journal",
]


def make_prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def script_uri(module_dot_sub):
    return f"vnd.sun.star.script:{module_dot_sub}?language=Basic&location=document"


def kill_stale():
    subprocess.run(["pkill", "-9", "-f", f"soffice.bin.*{PROFILE_DIR}"],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def start_soffice():
    kill_stale()
    if PROFILE_DIR.exists():
        shutil.rmtree(PROFILE_DIR)
    cmd = [
        "xvfb-run", "-a", "soffice",
        "--headless", "--invisible", "--nocrashreport", "--nodefault",
        "--norestore", "--nologo", "--nofirststartwizard",
        f"-env:UserInstallation=file://{PROFILE_DIR}",
        f"--accept=socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext",
    ]
    return subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)


def connect(retries=30):
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx)
    last_err = None
    for _ in range(retries):
        try:
            return resolver.resolve(f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext")
        except Exception as e:  # noqa: BLE001
            last_err = e
            time.sleep(1)
    raise RuntimeError(f"could not connect: {last_err}")


def invoke(doc, module_dot_sub, args=()):
    return doc.getScriptProvider().getScript(script_uri(module_dot_sub)).invoke(args, (), ())


def main():
    if len(sys.argv) < 2:
        print("usage: run_e2e_diagnostic.py <path-to-built-ods>")
        return 1
    source_ods = Path(sys.argv[1]).resolve()
    work_copy = Path("/tmp/prime_e2e_diag_work.ods")
    shutil.copyfile(source_ods, work_copy)

    if LOG_PATH.exists():
        LOG_PATH.unlink()

    proc = start_soffice()
    try:
        time.sleep(3)
        ctx = connect()
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(str(work_copy)), "_blank", 0, (make_prop("Hidden", True),))

        libs = doc.BasicLibraries
        if not libs.hasByName("Standard"):
            libs.createLibrary("Standard")
        if not libs.isLibraryLoaded("Standard"):
            libs.loadLibrary("Standard")
        lib = libs.getByName("Standard")

        print("Re-injecting current PRIME_*.bas sources (in case the .ods predates this run)...")
        for module_name in PRIME_MODULES:
            src = (SRC_BASIC_DIR / f"{module_name}.bas").read_text(encoding="utf-8")
            if lib.hasByName(module_name):
                lib.replaceByName(module_name, src)
            else:
                lib.insertByName(module_name, src)

        print("Injecting throwaway diagnostic module PRIME_ZZ_E2ETest...")
        diag_src = DIAGNOSTIC_MODULE.read_text(encoding="utf-8")
        if lib.hasByName("PRIME_ZZ_E2ETest"):
            lib.replaceByName("PRIME_ZZ_E2ETest", diag_src)
        else:
            lib.insertByName("PRIME_ZZ_E2ETest", diag_src)

        print("Ensuring schema/business sheets exist (two separate invokes, see known chaining bug)...")
        try:
            invoke(doc, "Standard.PRIME_14_MigrationInstaller.PRIME_Install_EnsureSchema")
        except Exception as e:  # noqa: BLE001
            print(f"  EnsureSchema exception (may be pre-existing/idempotent): {e}")
        try:
            invoke(doc, "Standard.PRIME_14_MigrationInstaller.PRIME_Install_EnsureAllBusinessSheetsOnly")
        except Exception as e:  # noqa: BLE001
            print(f"  EnsureAllBusinessSheetsOnly exception (may be pre-existing/idempotent): {e}")

        print("Running PRIME_ZZ_RunFullWorkflow (direct function invocation)...")
        try:
            result = invoke(doc, "Standard.PRIME_ZZ_E2ETest.PRIME_ZZ_RunFullWorkflow", (str(LOG_PATH),))
            print(f"Function returned: {result!r}")
        except Exception as e:  # noqa: BLE001
            print(f"invoke() raised (unusual - normally errors are swallowed, see module comment): {e}")

        doc.close(False)
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=15)
        except subprocess.TimeoutExpired:
            proc.kill()
        kill_stale()

    print()
    print("=== Log contents ===")
    if LOG_PATH.exists():
        text = LOG_PATH.read_text(encoding="utf-8", errors="replace")
        print(text)
        return 1 if ("FAIL" in text or "FATAL" in text) else 0
    else:
        print("NO LOG FILE WAS WRITTEN - the diagnostic Sub itself never completed (see above).")
        return 1


if __name__ == "__main__":
    sys.exit(main())
