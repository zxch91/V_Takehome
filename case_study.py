"""Validus case study: validate cashflows and calculate IRR, NAV and FX hedges."""

import argparse
from pathlib import Path
from zipfile import BadZipFile

from calculations import analyse
from cashflows import load_cashflows
from reporting import write_results


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", nargs="?", type=Path, default=Path("CST_Case_Study_data__1_.xlsx"))
    parser.add_argument("--fund", default="Fund I")
    parser.add_argument("--sheet")
    parser.add_argument("--output", type=Path, default=Path("outputs/results.xlsx"))
    args = parser.parse_args(argv)
    try:
        if args.input.resolve() == args.output.resolve():
            raise ValueError("output must differ from the input workbook")
        records, corrections = load_cashflows(args.input, args.sheet)
        results = analyse(records, args.fund)
        write_results(args.output, records, corrections, results)
    except (OSError, ValueError, BadZipFile) as error:
        parser.exit(1, f"Error: {error}\n")
    for row in results["IRR"]:
        print(f"{row['scope']} {row['currency']}: {row['annual_irr']:.6%}")
    print(f"Saved {args.output} ({len(corrections)} input corrections)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
