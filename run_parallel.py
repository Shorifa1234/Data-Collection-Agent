"""
run_parallel.py
---------------
Run multiple vendor scrapers simultaneously using your VPS's multi-core CPU.

Usage:
    # Run 2 vendors at the same time (recommended for 8-core / 16GB VPS)
    python run_parallel.py "Gabby" "Curry"

    # Run 3 at once (fine for 8-core, watch RAM if pages are JS-heavy)
    python run_parallel.py "Gabby" "Curry" "Parker Southern"

    # With flags passed to orchestrator
    python run_parallel.py "Gabby" "Curry" --headless true
    python run_parallel.py "Gabby" "Curry" --test

    # Run all 5 vendors in batches of 2
    python run_parallel.py --batch 2 "Parker Southern" "Gabby" "Curry" "Allied Maker" "Regina Andrew"

How it works:
    Each vendor is launched as a separate subprocess calling orchestrator.py.
    Output from each vendor is prefixed with [VendorName] and interleaved in real time.
    A timing summary is printed at the end.

Resource guide (8 cores / 16 GB RAM):
    2 vendors: safe for any scraper type (including JS-heavy sites)
    3 vendors: fine for lightweight scrapers; watch RAM if all use Playwright
    4 vendors: possible but may hit RAM limits; monitor with `htop`
"""

from __future__ import annotations

import argparse
import asyncio
import datetime
import subprocess
import sys
import time
from pathlib import Path

ROOT_DIR = Path(__file__).parent


async def _stream_output(proc, prefix: str) -> None:
    """Stream stdout from a subprocess, prefixing each line with the vendor name."""
    while True:
        line = await proc.stdout.readline()
        if not line:
            break
        print(f"[{prefix}] {line.decode('utf-8', errors='replace').rstrip()}", flush=True)


async def run_vendor(vendor: str, extra_args: list[str]) -> tuple[str, float, int]:
    """Launch orchestrator.py for one vendor and return (name, elapsed, returncode)."""
    cmd = [
        sys.executable, str(ROOT_DIR / "orchestrator.py"),
        vendor,
        *extra_args,
    ]
    t0 = time.monotonic()
    proc = await asyncio.create_subprocess_exec(
        *cmd,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.STDOUT,
        cwd=str(ROOT_DIR),
    )
    await _stream_output(proc, vendor)
    await proc.wait()
    elapsed = time.monotonic() - t0
    return vendor, elapsed, proc.returncode


async def run_batch(vendors: list[str], extra_args: list[str]) -> list[tuple[str, float, int]]:
    """Run all vendors in the batch concurrently."""
    tasks = [run_vendor(v, extra_args) for v in vendors]
    return await asyncio.gather(*tasks)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Run multiple vendor scrapers in parallel."
    )
    parser.add_argument("vendors", nargs="+", help="Vendor name(s) to scrape")
    parser.add_argument(
        "--batch", type=int, default=0,
        help="Process vendors in batches of N (0 = all at once)"
    )
    parser.add_argument("--test", action="store_true", help="Pass --test to orchestrator")
    parser.add_argument("--headless", default="true", help="Pass --headless flag (default: true)")
    parser.add_argument("--force-regen", action="store_true", help="Pass --force-regen to orchestrator")

    args = parser.parse_args()

    # Build extra args to pass through to orchestrator
    extra: list[str] = ["--headless", args.headless]
    if args.test:
        extra.append("--test")
    if args.force_regen:
        extra.append("--force-regen")

    vendors = args.vendors
    batch_size = args.batch or len(vendors)

    print("=" * 60)
    print(f"[Parallel Runner]  {len(vendors)} vendor(s)  |  batch={batch_size}")
    print(f"[Parallel Runner]  Started at {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")
    print("=" * 60)

    all_results: list[tuple[str, float, int]] = []
    wall_start = time.monotonic()

    for i in range(0, len(vendors), batch_size):
        batch = vendors[i : i + batch_size]
        print(f"\n>>> Batch {i // batch_size + 1}: {', '.join(batch)}\n")
        results = asyncio.run(run_batch(batch, extra))
        all_results.extend(results)

    wall_elapsed = time.monotonic() - wall_start

    # Summary
    print("\n" + "=" * 60)
    print("PARALLEL RUN SUMMARY")
    print("=" * 60)
    ok_count = 0
    for vendor, elapsed, rc in all_results:
        status = "OK" if rc == 0 else f"FAILED (exit {rc})"
        mins, secs = divmod(int(elapsed), 60)
        print(f"  {vendor:<30} {mins:02d}m {secs:02d}s   {status}")
        if rc == 0:
            ok_count += 1

    mins, secs = divmod(int(wall_elapsed), 60)
    print(f"\n  Wall time: {mins:02d}m {secs:02d}s   ({ok_count}/{len(all_results)} succeeded)")
    print("=" * 60)

    sys.exit(0 if ok_count == len(all_results) else 2)


if __name__ == "__main__":
    main()
