"""Simple helper to run user-supplied SQL files through pyodbc and pandas.

Usage checklist
----------------
1. Drop your SQL text files into a directory (for example, ``sql/``).
   Name them however you like; the filename (without ``.sql``) becomes the
   default query key when using the ``--directory`` option.
2. Fill in ``CONNECTION_STRING`` below with your ODBC details (Netezza, etc.),
   or pass ``--connection-string`` on the command line.
3. Either list query files explicitly in ``QUERY_FILES`` or point the script at
   a directory of ``.sql`` files with ``--directory path/to/sql``.
4. Run ``python manual_sql_loader.py`` to execute the queries. The script prints
   a short summary and can optionally write each DataFrame out as a CSV.

This script does **not** attempt to read SQL out of the workbook; it simply
executes whatever you provide.
"""
from __future__ import annotations

import argparse
from contextlib import closing
from pathlib import Path
from typing import Dict, Iterable, Mapping, Tuple

import pandas as pd
import pyodbc

# ---------------------------------------------------------------------------
# User configuration
# ---------------------------------------------------------------------------

# Populate with your connection string (or pass --connection-string at runtime).
CONNECTION_STRING: str = ""

# Explicit mapping of query name -> SQL file path. Paths are resolved relative
# to this script unless they are absolute. You can keep this empty and rely on
# the --directory flag instead.
QUERY_FILES: Mapping[str, str] = {
    # "submissions": "sql/submissions.sql",
    # "other_query": "sql/other_query.sql",
}

# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent


def resolve_query_files(query_files: Mapping[str, str]) -> Dict[str, Path]:
    resolved: Dict[str, Path] = {}
    for name, path_text in query_files.items():
        path = Path(path_text)
        if not path.is_absolute():
            path = BASE_DIR / path
        resolved[name] = path
    return resolved


def collect_directory_queries(directory: Path, pattern: str = "*.sql", recursive: bool = False) -> Dict[str, Path]:
    if recursive:
        paths = sorted(directory.rglob(pattern))
    else:
        paths = sorted(directory.glob(pattern))
    queries: Dict[str, Path] = {}
    for path in paths:
        if path.is_file():
            key = path.stem
            if key in queries:
                raise ValueError(f"Duplicate query key detected from directory scan: {key}")
            queries[key] = path
    return queries


def read_sql_file(path: Path) -> str:
    if not path.exists():
        raise FileNotFoundError(f"SQL file not found: {path}")
    return path.read_text(encoding="utf-8")


def load_queries(
    connection_string: str,
    query_files: Mapping[str, Path],
    *,
    autocommit: bool = True,
) -> Dict[str, pd.DataFrame]:
    if not connection_string:
        raise ValueError("Connection string is empty; set CONNECTION_STRING or pass --connection-string.")
    if not query_files:
        raise ValueError("No queries provided. Configure QUERY_FILES or use --directory.")

    statements: Dict[str, str] = {}
    for name, path in query_files.items():
        sql_text = read_sql_file(path)
        statements[name] = sql_text

    results: Dict[str, pd.DataFrame] = {}
    with closing(pyodbc.connect(connection_string, autocommit=autocommit)) as conn:
        for name, sql in statements.items():
            results[name] = pd.read_sql(sql, conn)
    return results


def save_results(dfs: Mapping[str, pd.DataFrame], output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    for name, df in dfs.items():
        output_path = output_dir / f"{name}.csv"
        df.to_csv(output_path, index=False)


def build_query_map(args: argparse.Namespace) -> Dict[str, Path]:
    if args.directory:
        directory = args.directory
        if not directory.is_absolute():
            directory = (Path.cwd() / directory).resolve()
        queries = collect_directory_queries(directory, pattern=args.pattern, recursive=args.recursive)
        if args.append_default:
            queries.update(resolve_query_files(QUERY_FILES))
        return queries
    return resolve_query_files(QUERY_FILES)


def parse_args(argv: Iterable[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Execute user-provided SQL files with pyodbc and pandas.")
    parser.add_argument(
        "--connection-string",
        dest="connection_string",
        help="ODBC connection string (overrides CONNECTION_STRING constant).",
    )
    parser.add_argument(
        "--directory",
        type=Path,
        help="Directory containing .sql files (filenames become query names).",
    )
    parser.add_argument(
        "--pattern",
        default="*.sql",
        help="Glob pattern when scanning --directory (default: *.sql).",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Recurse into subdirectories when using --directory.",
    )
    parser.add_argument(
        "--append-default",
        action="store_true",
        help="When using --directory, also include queries listed in QUERY_FILES.",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="List discovered queries without executing them.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="Optional directory to write query results as CSV files.",
    )
    parser.add_argument(
        "--no-autocommit",
        action="store_true",
        help="Disable autocommit on the ODBC connection.",
    )
    parser.add_argument(
        "--summary-only",
        action="store_true",
        help="Print DataFrame shapes only; skip saving results unless --output-dir is set.",
    )
    return parser.parse_args(argv)


def main(argv: Iterable[str] | None = None) -> int:
    args = parse_args(argv)
    query_map = build_query_map(args)

    if not query_map:
        print("No SQL queries configured. Add entries to QUERY_FILES or use --directory.")
        return 1

    if args.list:
        print("Queries ready for execution:")
        for name, path in sorted(query_map.items()):
            print(f"  {name}: {path}")
        return 0

    connection_string = args.connection_string or CONNECTION_STRING
    try:
        results = load_queries(
            connection_string,
            query_map,
            autocommit=not args.no_autocommit,
        )
    except Exception as exc:  # pragma: no cover - CLI surface area
        print(f"Error: {exc}")
        return 2

    if args.output_dir:
        save_results(results, args.output_dir)
        print(f"Saved results to {args.output_dir}")

    if not args.summary_only or not args.output_dir:
        print("Query execution summary:")
        for name, df in results.items():
            print(f"  {name}: {df.shape[0]} rows, {df.shape[1]} columns")

    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
