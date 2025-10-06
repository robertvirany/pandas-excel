"""Generate pandas SQL loader code from Power Query definitions in an Excel workbook."""
from __future__ import annotations

import argparse
import re
import textwrap
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET


STRING_LITERAL_RE = re.compile(r'"((?:[^"]|"")*)"', re.DOTALL)
NATIVE_QUERY_RE = re.compile(
    r'(?:Value|Sql)\.NativeQuery\s*\(\s*[^,]+,\s*"((?:[^"]|"")*)"',
    re.IGNORECASE | re.DOTALL,
)
ODBC_QUERY_RE = re.compile(
    r'Odbc\.(?:Query|Execute)\s*\(\s*[^,]+,\s*"((?:[^"]|"")*)"',
    re.IGNORECASE | re.DOTALL,
)
QUERY_PARAM_RE = re.compile(r'Query\s*=\s*"((?:[^"]|"")*)"', re.IGNORECASE | re.DOTALL)
SQL_DATABASE_HINT_RE = re.compile(
    r'Sql\.Database\s*\(\s*"((?:[^"]|"")*)"\s*,\s*"((?:[^"]|"")*)"',
    re.IGNORECASE,
)
ODBC_DATASOURCE_HINT_RE = re.compile(
    r'Odbc\.DataSource\s*\(\s*"((?:[^"]|"")*)"',
    re.IGNORECASE,
)
ODBC_QUERY_SOURCE_HINT_RE = re.compile(
    r'Odbc\.(?:Query|Execute)\s*\(\s*"((?:[^"]|"")*)"',
    re.IGNORECASE,
)

SQL_KEYWORDS = ("SELECT", "INSERT", "UPDATE", "DELETE", "WITH", "MERGE")


@dataclass
class ExtractedQuery:
    name: str
    formula: str
    sql_blocks: List[str]
    connection_hints: List[str]
    source_part: str


def unescape_m_string(value: str) -> str:
    return value.replace('""', '"')


def parse_query_xml(xml_text: str, fallback_name: str, source_part: str) -> ExtractedQuery:
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return ExtractedQuery(fallback_name, xml_text, [], [], source_part)

    def local_name(tag: str) -> str:
        return tag.split('}', 1)[-1] if '}' in tag else tag

    name = fallback_name
    formula: Optional[str] = None
    for elem in root.iter():
        tag = local_name(elem.tag).lower()
        text = (elem.text or "").strip()
        if tag == "name" and text:
            name = text
        elif tag in {"queryformula", "formula"} and text:
            formula = elem.text
    if formula is None:
        formula = "\n".join(part.strip() for part in xml_text.splitlines() if part.strip())
    return ExtractedQuery(name, formula, [], [], source_part)


def extract_sql_blocks(formula: str) -> List[str]:
    results: List[str] = []
    seen: set[str] = set()

    def add(sql_raw: str) -> None:
        sql_text = unescape_m_string(sql_raw).strip()
        if not sql_text:
            return
        upper = sql_text.upper()
        if not any(keyword in upper for keyword in SQL_KEYWORDS):
            return
        if sql_text not in seen:
            seen.add(sql_text)
            results.append(sql_text)

    for pattern in (NATIVE_QUERY_RE, ODBC_QUERY_RE, QUERY_PARAM_RE):
        for match in pattern.finditer(formula):
            add(match.group(1))

    if not results:
        for match in STRING_LITERAL_RE.finditer(formula):
            sql_candidate = unescape_m_string(match.group(1)).strip()
            upper = sql_candidate.upper()
            if any(keyword in upper for keyword in SQL_KEYWORDS):
                add(match.group(1))

    return results


def extract_connection_hints(formula: str) -> List[str]:
    hints: List[str] = []
    seen: set[str] = set()
    for match in SQL_DATABASE_HINT_RE.finditer(formula):
        server = unescape_m_string(match.group(1))
        database = unescape_m_string(match.group(2))
        hint = f"Sql.Database(server='{server}', database='{database}')"
        if hint not in seen:
            hints.append(hint)
            seen.add(hint)
    for match in ODBC_DATASOURCE_HINT_RE.finditer(formula):
        source = unescape_m_string(match.group(1))
        hint = f"Odbc.DataSource('{source}')"
        if hint not in seen:
            hints.append(hint)
            seen.add(hint)
    for match in ODBC_QUERY_SOURCE_HINT_RE.finditer(formula):
        source = unescape_m_string(match.group(1))
        hint = f"Odbc.Query('{source}', ...)"
        if hint not in seen:
            hints.append(hint)
            seen.add(hint)
    return hints


def read_query_definitions(workbook_path: Path) -> List[ExtractedQuery]:
    queries: List[ExtractedQuery] = []
    with zipfile.ZipFile(workbook_path, "r") as archive:
        for name in archive.namelist():
            if not name.startswith("xl/queries/") or not name.endswith(".xml"):
                continue
            data = archive.read(name)
            try:
                xml_text = data.decode("utf-8")
            except UnicodeDecodeError:
                xml_text = data.decode("utf-8", errors="replace")
            fallback = Path(name).stem
            query = parse_query_xml(xml_text, fallback, source_part=name)
            query.sql_blocks = extract_sql_blocks(query.formula)
            query.connection_hints = extract_connection_hints(query.formula)
            queries.append(query)
    return queries


def generate_loader_code(
    queries: Iterable[ExtractedQuery],
    workbook: Path,
) -> str:
    sql_queries: Dict[str, Tuple[str, ExtractedQuery]] = {}
    duplicate_counter: Dict[str, int] = {}

    ordered_queries: List[Tuple[str, str, ExtractedQuery]] = []
    for query in queries:
        if not query.sql_blocks:
            continue
        for idx, sql in enumerate(query.sql_blocks, start=1):
            base_key = query.name or "Query"
            key = base_key
            if len(query.sql_blocks) > 1:
                key = f"{base_key}_{idx}"
            if key in sql_queries:
                duplicate_counter[key] = duplicate_counter.get(key, 1) + 1
                key = f"{key}_{duplicate_counter[key]}"
            sql_queries[key] = (sql, query)
            ordered_queries.append((key, sql, query))

    header_lines = [
        f'"""Auto-generated SQL loaders from {workbook.name}."""',
        "import pandas as pd",
        "import pyodbc",
        "from contextlib import closing",
        "",
        "# This file was generated by generate_sql_loader.py. Edit with care.",
    ]
    header = "\n".join(header_lines)

    if not ordered_queries:
        body_lines = [
            header.rstrip(),
            "# No SQL-backed Power Queries were detected.",
            "",
            "QUERIES = {}",
            "",
            "",
            "def load_data(connection_string: str, autocommit: bool = True) -> dict:",
            "    \"\"\"Return an empty dict because no SQL queries were found.\"\"\"",
            "    return {}",
            "",
        ]
        return "\n".join(body_lines)

    lines: List[str] = [header.rstrip(), "", "QUERIES = {"]
    for key, sql, query in ordered_queries:
        if query.connection_hints:
            comment = "; ".join(query.connection_hints)
            lines.append(f"    # Source hints: {comment}")
        lines.append(f"    {key!r}: \"\"\"\n{sql}\n\"\"\",")
    lines.append("}")
    lines.append("")
    lines.append("")
    lines.append("def load_data(connection_string: str, autocommit: bool = True) -> dict:")
    lines.append(
        "    \"\"\"Execute the extracted SQL queries and return a dict of DataFrames.\"\"\"")
    lines.append("    dfs = {}")
    lines.append(
        "    with closing(pyodbc.connect(connection_string, autocommit=autocommit)) as conn:")
    lines.append("        for name, sql in QUERIES.items():")
    lines.append("            dfs[name] = pd.read_sql(sql, conn)")
    lines.append("    return dfs")
    lines.append("")
    return "\n".join(lines)


def write_loader_file(code: str, output_path: Path, overwrite: bool) -> None:
    if output_path.exists() and not overwrite:
        raise FileExistsError(f"Refusing to overwrite existing file: {output_path}")
    output_path.write_text(code, encoding="utf-8")


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path, help="Path to the Excel workbook (.xlsx)")
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("load_queries.py"),
        help="Destination path for the generated loader (default: load_queries.py)",
    )
    parser.add_argument("--overwrite", action="store_true", help="Allow overwriting the output file")
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    queries = read_query_definitions(args.workbook)
    code = generate_loader_code(queries, args.workbook)
    write_loader_file(code, args.output, overwrite=args.overwrite)

    total = len(queries)
    with_sql = sum(1 for q in queries if q.sql_blocks)
    print(f"Discovered {with_sql} SQL-backed query(ies) out of {total} total.")
    print(f"Generated loader written to {args.output}")
    if with_sql == 0:
        print("No SQL statements detected; the loader returns an empty dict.")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
