#!/usr/bin/env python3
"""
prime-libreoffice-linux-smoke CI check: open the built .ods headless, store it under a new
name, close, reopen the stored copy, and verify the essential PRIME structure (hidden system
sheets, embedded Basic library) survived the round trip.

Deliberately does NOT invoke any PRIME Basic macro (unlike tests/functional_smoke.py) - that
external-invoke() chaining path was found to be unreliable in this headless environment (see
docs/KNOWN_ISSUES.md #11 and docs/TEST_MATRIX.md). This check only exercises the UNO
open/store/close/reopen lifecycle itself, which is a different and more reliable code path,
so it is a legitimate persistence check rather than a proof that macro execution works.
"""
import subprocess
import sys
import time
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue

EXPECTED_HIDDEN_SHEETS = [
    "SYS_PRIME_META", "SYS_PRIME_SEQ", "SYS_PRIME_TX",
    "DB_PRIME_PRODUCTS", "DB_PRIME_MOVEMENTS", "DB_PRIME_LOTS",
]


def make_prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def connect(port: int, profile_dir: Path):
    proc = subprocess.Popen(
        ["xvfb-run", "-a", "soffice", "--headless", "--invisible", "--nocrashreport",
         "--nodefault", "--norestore", "--nologo", "--nofirststartwizard",
         f"-env:UserInstallation=file://{profile_dir}",
         f"--accept=socket,host=localhost,port={port};urp;StarOffice.ComponentContext"],
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
    )
    time.sleep(4)
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx)
    ctx = None
    last_exc = None
    for _ in range(30):
        try:
            ctx = resolver.resolve(
                f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
            break
        except Exception as exc:  # noqa: BLE001
            last_exc = exc
            time.sleep(1)
    if ctx is None:
        proc.terminate()
        raise RuntimeError(f"could not connect to soffice: {last_exc}")
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return proc, desktop


def main():
    if len(sys.argv) < 2:
        print("usage: smoke_open_save_reopen.py <path-to-ods>")
        return 1
    ods_path = Path(sys.argv[1]).resolve()
    reopened_path = ods_path.with_name(ods_path.stem + "_reopened.ods")
    profile_dir = Path("/tmp/prime_smoke_profile")

    failures = []
    proc, desktop = connect(2004, profile_dir)
    try:
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(str(ods_path)), "_blank", 0, (make_prop("Hidden", True),))
        try:
            for sheet in EXPECTED_HIDDEN_SHEETS:
                if not doc.Sheets.hasByName(sheet):
                    failures.append(f"missing sheet before save: {sheet}")

            save_args = (make_prop("FilterName", "calc8"),)
            doc.storeToURL(uno.systemPathToFileUrl(str(reopened_path)), save_args)
        finally:
            doc.close(False)

        doc2 = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(str(reopened_path)), "_blank", 0, (make_prop("Hidden", True),))
        try:
            for sheet in EXPECTED_HIDDEN_SHEETS:
                if not doc2.Sheets.hasByName(sheet):
                    failures.append(f"missing sheet after reopen: {sheet}")
            libs = doc2.BasicLibraries
            if not libs.hasByName("Standard"):
                failures.append("Standard Basic library missing after reopen")
        finally:
            doc2.close(False)
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=15)
        except Exception:  # noqa: BLE001
            proc.kill()
        reopened_path.unlink(missing_ok=True)

    if failures:
        print("FAILED:")
        for f in failures:
            print(f"  [FAIL] {f}")
        return 1
    print("PASSED: open/store/close/reopen round trip preserved PRIME structure")
    return 0


if __name__ == "__main__":
    sys.exit(main())
