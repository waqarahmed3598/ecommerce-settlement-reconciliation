#!/usr/bin/env python3
"""One-command demo.

Runs the reconciliation engine on the synthetic sample data in sample_data/,
then runs the independent validator, which recomputes every figure from the
raw CSVs and asserts the workbook matches.

    python3 demo.py

Expected headline result (October 2026 sample month, fictional store):
revenue $9,824.97 - true net profit $313.68 - 14 exceptions found -
$268.89 net leakage (expected payouts $7,897.58 vs actually paid $7,628.69) -
32/32 validation checks PASS.
"""
import subprocess
import sys
from pathlib import Path

BASE = Path(__file__).parent


def run(script: str) -> None:
    print(f"\n===== python3 {script} =====", flush=True)
    subprocess.run([sys.executable, str(BASE / script)], check=True)


if __name__ == "__main__":
    run("reconcile.py")
    run("validate.py")
    print("\nDemo complete: see output/Ecommerce-Profit-Settlement-Reconciliation.xlsx")
