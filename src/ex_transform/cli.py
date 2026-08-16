"""Unified command line entry point for running, diagnosing, and testing."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .app import run as run_app
from .diagnostics import collect_environment_report, format_report


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="excel-search")
    parser.add_argument("command", nargs="?", choices=("start", "diagnose", "test", "check"), default="start")
    parser.add_argument("--no-excel-probe", action="store_true", help=argparse.SUPPRESS)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.command == "diagnose":
        print(format_report(collect_environment_report(probe_excel=not args.no_excel_probe)))
        return 0
    if args.command == "test":
        try:
            import pytest
        except ImportError:
            print("pytest is required for the test command", file=sys.stderr)
            return 2
        project_root = Path(__file__).resolve().parents[2]
        return int(pytest.main([str(project_root / "tests")]))
    if args.command == "check":
        source_root = Path(__file__).resolve().parent
        try:
            for source_file in source_root.glob("*.py"):
                compile(source_file.read_text(encoding="utf-8"), str(source_file), "exec")
        except (OSError, SyntaxError) as exc:
            print(f"Source compilation failed: {exc}", file=sys.stderr)
            return 1
        print("Source compilation: PASS")
        return 0
    return run_app()


if __name__ == "__main__":
    raise SystemExit(main())
