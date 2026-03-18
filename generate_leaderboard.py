"""Build a MapTap leaderboard from an exported iMessage group chat PDF."""

from __future__ import annotations

import datetime
import html
import re
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "TapMap.pdf"
OUTPUT_FILE = BASE_DIR / "leaderboard.csv"
EXCEL_OUTPUT_FILE = BASE_DIR / "leaderboard.xlsx"
MARKDOWN_OUTPUT_FILE = BASE_DIR / "leaderboard.md"
HTML_OUTPUT_FILE = BASE_DIR / "tapmap-leaderboard.html"
START_DATE = datetime.date(2026, 3, 16)
RIGHT_SIDE_INDENT = 60

# Map raw sender names to short names shown on the leaderboard.
NAME_MAP = {
    "Alma": "Alma",
    "Truman": "Truman",
    "Michelle": "Michelle",
    "Michelle Ben": "Michelle",
    "Cooper": "Cooper",
    "Cooper Cooper": "Cooper",
    "CooperCooper": "Cooper",
    "Ben": "Ben",
    "Ben Awad": "Ben",
    "BenAwad": "Ben",
    "Matt": "Matt",
    "Matt Ruth": "Matt",
    "MattRuth": "Matt",
}
DEFAULT_SELF = "Alma"
KNOWN_PLAYERS = {"Alma", "Truman", "Michelle", "Cooper", "Ben", "Matt"}


HEADER_RE = re.compile(r"^\s*(?P<name>.+?)\s*[—-]\s*(?P<meta>.+?)\s*$")
MAPTAP_RE = re.compile(r"^\s*www\.maptap\.gg\s+(?P<date>.+?)\s*$", re.IGNORECASE)
FINAL_RE = re.compile(r"^\s*Final score:\s*(?P<score>\d+)\s*$", re.IGNORECASE)
TIMESTAMP_RE = re.compile(r"^\s*(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|Today|Yesterday)\b", re.IGNORECASE)
NAME_ONLY_RE = re.compile(r"^\s*[A-Za-z][A-Za-z0-9 .'\-]+\s*$")
SCORE_TOKEN_RE = re.compile(r"\d+")

LINE_INDEX_TEXT = 0
LINE_INDEX_INDENT = 1
LINE_INDEX_RIGHT = 2


def _normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _alpha_name(text: str) -> str:
    return re.sub(r"[^A-Za-z.'\- ]", "", text)


def _compact_text(text: str) -> str:
    return re.sub(r"\s+", "", text).strip()


def _compact_alpha(text: str) -> str:
    return re.sub(r"[^A-Za-z]", "", _normalize_spaces(text))


def _line_indent(text: str) -> int:
    return len(text) - len(text.lstrip(" "))


def _extract_round_scores(text: str) -> list[int]:
    scores: list[int] = []
    for token in SCORE_TOKEN_RE.findall(text):
        raw = int(token)
        if 0 <= raw <= 100:
            scores.append(raw)
            continue

        if len(token) == 3:
            first_two = int(token[:2])
            last = int(token[2:])
            if 0 <= first_two <= 100:
                scores.extend([first_two, last])
                continue

            first = int(token[0])
            last_two = int(token[1:])
            if 0 <= first <= 100 and 0 <= last_two <= 100:
                scores.extend([first, last_two])
                continue
    return scores

MONTHS = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}


def _extract_year(meta_line: str) -> int | None:
    match = re.search(r"\b(19|20)\d{2}\b", meta_line)
    return int(match.group(0)) if match else None


def _parse_game_date(date_text: str, year_hint: int | None) -> str | None:
    text = date_text.strip().replace("  ", " ")
    match = re.match(r"^([A-Za-z]+)\s+(\d{1,2})(?:,\s*(\d{4}))?\s*$", text)
    if not match:
        return None

    month_name, day_text, year_text = match.groups()
    month = MONTHS.get(month_name.strip().lower())
    if month is None:
        return None

    year = int(year_text) if year_text else (year_hint or datetime.datetime.now().year)
    day = int(day_text)
    return f"{year:04d}-{month:02d}-{day:02d}"


def _looks_like_name(text: str) -> bool:
    stripped = _normalize_spaces(text)
    alpha_clean = _normalize_spaces(_alpha_name(text))
    if not alpha_clean or not NAME_ONLY_RE.match(alpha_clean):
        return False
    if TIMESTAMP_RE.match(stripped):
        return False
    if MAPTAP_RE.match(stripped):
        return False
    if FINAL_RE.match(stripped):
        return False
    if alpha_clean.startswith("www"):
        return False
    if " added " in stripped:
        return False
    if stripped.isdigit():
        return False
    return _resolve_name(alpha_clean) is not None


def _resolve_name(raw_name: str) -> str | None:
    candidates = (
        raw_name,
        _normalize_spaces(raw_name),
        _alpha_name(raw_name),
        _normalize_spaces(_alpha_name(raw_name)),
        _compact_text(raw_name),
        _compact_alpha(raw_name),
    )
    for candidate in candidates:
        mapped = NAME_MAP.get(candidate.strip(), candidate.strip())
        if mapped in KNOWN_PLAYERS:
            return mapped
    return None


def _read_messages(path: Path) -> list[tuple[str, int, bool]]:
    try:
        from pypdf import PdfReader
    except ModuleNotFoundError:
        raise SystemExit("Missing dependency: pypdf. Install with `python3 -m pip install pypdf`.")

    try:
        reader = PdfReader(str(path))
    except Exception as exc:
        raise SystemExit(f"Unable to open PDF: {path} ({exc})")

    lines: list[tuple[str, int, bool]] = []
    for page in reader.pages:
        page_text = page.extract_text(extraction_mode="layout") or ""
        page_text = page_text.replace("\r\n", "\n").replace("\r", "\n")
        for raw_line in page_text.split("\n"):
            if not raw_line.strip():
                continue
            indent = _line_indent(raw_line)
            is_right = indent >= RIGHT_SIDE_INDENT
            lines.append((raw_line.strip(), indent, is_right))
    return lines


def _lookbehind_sender(
    lines: list[tuple[str, int, bool]], start: int, max_lookbehind: int = 6
) -> tuple[str | None, int | None]:
    for dist in range(1, max_lookbehind + 1):
        j = start - dist
        if j < 0:
            break
        if lines[j][LINE_INDEX_RIGHT]:
            continue
        candidate = _resolve_name(lines[j][LINE_INDEX_TEXT])
        if candidate is not None:
            return candidate, dist
    return None, None


def parse_records(path: Path) -> list[dict]:
    lines = _read_messages(path)
    candidate_records: list[dict] = []
    total = len(lines)
    current_year_hint: int | None = None
    pending_sender: str | None = None
    pending_sender_ttl = 0
    idx = 0

    while idx < total:
        line = lines[idx][LINE_INDEX_TEXT]
        line_is_right = lines[idx][LINE_INDEX_RIGHT]

        header_match = HEADER_RE.match(line)
        if header_match:
            player = _resolve_name(header_match.group("name").strip())
            if player is not None:
                current_year_hint = _extract_year(header_match.group("meta"))
                pending_sender = player
                pending_sender_ttl = 3
            idx += 1
            continue

        resolved_name = _resolve_name(line)
        if resolved_name is not None and _looks_like_name(line):
            if not line_is_right:
                pending_sender = resolved_name
                pending_sender_ttl = 3
            idx += 1
            continue

        if MAPTAP_RE.match(line):
            sender = DEFAULT_SELF if line_is_right else None
            if sender is None:
                sender = pending_sender if pending_sender_ttl > 0 else None
            if sender is None:
                nearby_sender, distance = _lookbehind_sender(lines, idx, max_lookbehind=6)
                if nearby_sender is not None and distance is not None and distance <= 6:
                    sender = nearby_sender
            if sender is None:
                idx += 1
                continue
            sender = sender or DEFAULT_SELF

            parsed_date = _parse_game_date(MAPTAP_RE.match(line).group("date"), current_year_hint)
            if not parsed_date:
                idx += 1
                continue
            game_date_obj = datetime.date.fromisoformat(parsed_date)
            if game_date_obj < START_DATE:
                idx += 1
                continue

            final_score: int | None = None
            round_scores: list[int] = []
            lookahead = idx + 1
            while lookahead < total:
                block_line, _, block_is_right = lines[lookahead]
                if MAPTAP_RE.match(block_line):
                    break
                if block_is_right != line_is_right and (final_score is not None or round_scores):
                    break
                if block_is_right != line_is_right and not round_scores and final_score is None:
                    lookahead += 1
                    continue
                if FINAL_RE.match(block_line):
                    final_score = int(FINAL_RE.match(block_line).group("score"))
                    if len(round_scores) >= 5:
                        break
                    lookahead += 1
                    continue
                header_match_inside = HEADER_RE.match(block_line)
                if header_match_inside is not None and _resolve_name(header_match_inside.group("name").strip()) is not None:
                    break
                if _looks_like_name(block_line):
                    break
                if len(round_scores) < 5:
                    for value in _extract_round_scores(block_line):
                        round_scores.append(int(value))
                        if len(round_scores) >= 5:
                            break
                if final_score is not None and len(round_scores) >= 5:
                    break
                lookahead += 1

            if final_score is not None and len(round_scores) >= 5:
                candidate_records.append(
                    {
                        "Name": sender,
                        "game_date": parsed_date,
                        "r1": round_scores[0],
                        "r2": round_scores[1],
                        "r3": round_scores[2],
                        "r4": round_scores[3],
                        "r5": round_scores[4],
                        "final_score": final_score,
                    }
                )

            idx = lookahead
            pending_sender = sender
            pending_sender_ttl = 0
            continue

        if pending_sender is not None and pending_sender_ttl > 0:
            pending_sender_ttl -= 1
            if pending_sender_ttl == 0:
                pending_sender = None

        idx += 1

    if not candidate_records:
        return []

    best_per_player_day: dict[tuple[str, str], dict] = {}
    for rec in candidate_records:
        key = (rec["Name"], rec["game_date"])
        existing = best_per_player_day.get(key)
        if existing is None or rec["final_score"] > existing["final_score"]:
            best_per_player_day[key] = rec

    return list(best_per_player_day.values())


def build_leaderboard(records: list[dict]) -> pd.DataFrame:
    if not records:
        return pd.DataFrame(
            columns=[
                "Rank",
                "Name",
                "Games Played",
                "Daily Wins",
                "Final",
                "Round 1",
                "Round 2",
                "Round 3",
                "Round 4",
                "Round 5",
            ]
        )

    df = pd.DataFrame(records)
    df["Daily Win"] = (
        (df["final_score"] == df.groupby("game_date")["final_score"].transform("max"))
        .astype(int)
    )

    grouped = (
        df.groupby("Name", as_index=False)
        .agg(
            games_played=("game_date", "count"),
            daily_wins=("Daily Win", "sum"),
            r1_avg=("r1", "mean"),
            r2_avg=("r2", "mean"),
            r3_avg=("r3", "mean"),
            r4_avg=("r4", "mean"),
            r5_avg=("r5", "mean"),
            final_avg=("final_score", "mean"),
        )
        .sort_values("final_avg", ascending=False)
        .reset_index(drop=True)
    )

    grouped["Rank"] = grouped.index + 1
    leaderboard = grouped[
        [
            "Rank",
            "Name",
            "games_played",
            "daily_wins",
            "final_avg",
            "r1_avg",
            "r2_avg",
            "r3_avg",
            "r4_avg",
            "r5_avg",
        ]
    ].rename(
        columns={
            "games_played": "Games Played",
            "daily_wins": "Daily Wins",
            "final_avg": "Final",
            "r1_avg": "Round 1",
            "r2_avg": "Round 2",
            "r3_avg": "Round 3",
            "r4_avg": "Round 4",
            "r5_avg": "Round 5",
        }
    )

    leaderboard[["Games Played", "Daily Wins"]] = leaderboard[["Games Played", "Daily Wins"]].astype(int)
    leaderboard[
        [
            "Final",
            "Round 1",
            "Round 2",
            "Round 3",
            "Round 4",
            "Round 5",
        ]
    ] = leaderboard[
        [
            "Final",
            "Round 1",
            "Round 2",
            "Round 3",
            "Round 4",
            "Round 5",
        ]
    ].round(2)

    return leaderboard


def write_excel(leaderboard: pd.DataFrame, path: Path) -> bool:
    # Try OpenPyXL first for better formatting support.
    try:
        import openpyxl  # type: ignore
    except ModuleNotFoundError:
        openpyxl = None

    if openpyxl is not None:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            leaderboard.to_excel(writer, index=False, sheet_name="Leaderboard")
            workbook = writer.book
            sheet = writer.sheets["Leaderboard"]

            for idx, column in enumerate(leaderboard.columns, start=1):
                width = max(len(str(column)), 10)
                for value in leaderboard[column]:
                    width = max(width, len("" if pd.isna(value) else str(value)))
                sheet.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = min(width + 2, 24)

            sheet.freeze_panes = "A2"
            sheet.print_area = f"A1:{openpyxl.utils.get_column_letter(len(leaderboard.columns))}{len(leaderboard) + 1}"
            sheet.page_setup.paperSize = 1
            sheet.page_setup.orientation = "landscape"
            sheet.page_setup.fitToHeight = 0
            sheet.page_setup.fitToWidth = 1
            sheet.page_margins.left = 0.25
            sheet.page_margins.right = 0.25
            sheet.page_margins.top = 0.25
            sheet.page_margins.bottom = 0.25
            sheet.page_margins.header = 0.2
            sheet.page_margins.footer = 0.2
        return True

    # Fallback if openpyxl is unavailable.
    try:
        import xlsxwriter  # type: ignore
    except ModuleNotFoundError:
        print(
            "Excel export unavailable: install one writer dependency with"
            " `python3 -m pip install openpyxl xlsxwriter`."
        )
        return False

    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        leaderboard.to_excel(writer, index=False, sheet_name="Leaderboard")
        workbook = writer.book
        sheet = writer.sheets["Leaderboard"]

        header_format = workbook.add_format({"bold": True})
        for col, column in enumerate(leaderboard.columns):
            sheet.write(0, col, column, header_format)

        for idx, column in enumerate(leaderboard.columns):
            max_len = max([len(str(column))] + [len("" if pd.isna(v) else str(v)) for v in leaderboard[column]])
            sheet.set_column(idx, idx, min(max_len + 2, 24))

        sheet.set_paper(1)
        sheet.set_landscape()
        sheet.fit_to_pages(1, 0)
        sheet.set_print_scale(100)
        sheet.set_margins(0.25, 0.25, 0.25, 0.25)
        sheet.freeze_panes(1, 0)
    return True


def _fmt_md(value: object) -> str:
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value.is_integer():
            return f"{int(value)}"
        if isinstance(value, float):
            return f"{value:.2f}".rstrip("0").rstrip(".")
        return str(int(value))
    return str(value)


def _generated_at() -> str:
    return datetime.datetime.now(datetime.timezone.utc).astimezone().strftime("%Y-%m-%d %I:%M:%S %Z%z")


def write_markdown(leaderboard: pd.DataFrame, path: Path) -> None:
    columns = list(leaderboard.columns)
    header = "| " + " | ".join(columns) + " |"
    divider = "| " + " | ".join(["---"] * len(columns)) + " |"
    rows: list[str] = []

    for _, row in leaderboard.iterrows():
        rows.append("| " + " | ".join(_fmt_md(row[col]) for col in columns) + " |")

    body = "\n".join([header, divider, *rows])
    content = (
        "# MapTap Leaderboard\n"
        f"Generated: {_generated_at()}\n\n"
        + body
        + "\n"
    )
    path.write_text(content, encoding="utf-8")


def write_html(leaderboard: pd.DataFrame, path: Path) -> None:
    final_col_idx = (
        leaderboard.columns.get_loc("Final") if "Final" in leaderboard.columns else max(len(leaderboard.columns) - 1, 0)
    )
    header_cells = "\n".join(
        (
            f'      <th class="final-col">{html.escape(str(col))}</th>'
            if col_idx == final_col_idx
            else f"      <th>{html.escape(str(col))}</th>"
        )
        for col_idx, col in enumerate(leaderboard.columns)
    )

    if leaderboard.empty:
        body_rows = '      <tr><td colspan="10">No games found yet.</td></tr>'
    else:
        row_lines: list[str] = []
        for row_idx, row in enumerate(leaderboard.itertuples(index=False), start=1):
            class_name = "rank-top" if row_idx == 1 else "rank-med" if row_idx == 2 else "rank-low" if row_idx == 3 else ""
            cells = []
            for col_idx, value in enumerate(row):
                if col_idx == final_col_idx:
                    cells.append(f'        <td class="final-col">{html.escape(_fmt_md(value))}</td>')
                else:
                    cells.append(f'        <td>{html.escape(_fmt_md(value))}</td>')
            body_rows = "\n".join(
                [
                    "    <tr" + (f' class="{class_name}"' if class_name else "") + ">",
                    *cells,
                    "    </tr>",
                ]
            )
            row_lines.append(body_rows)
        body_rows = "\n".join(row_lines)

    content = f"""<!doctype html>
<html lang=\"en\">
  <head>
    <meta charset=\"UTF-8\" />
    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
    <title>MapTap Leaderboard</title>
    <style>
      :root {{
        --bg: #0f1120;
        --panel: #171a2c;
        --panel-soft: #1f2440;
        --text: #e9eeff;
        --muted: #a9b0d0;
        --line: #2e3765;
        --accent-1: #6ff7ff;
        --accent-2: #ffd166;
        --accent-3: #8de969;
      }}
      * {{ box-sizing: border-box; }}
      body {{
        margin: 0;
        min-height: 100vh;
        font-family: "Inter", "Segoe UI", Arial, sans-serif;
        background:
          radial-gradient(circle at 18% 15%, rgba(111, 247, 255, 0.15), transparent 42%),
          radial-gradient(circle at 82% 8%, rgba(141, 233, 105, 0.14), transparent 36%),
          var(--bg);
        color: var(--text);
      }}
      .wrap {{
        max-width: 980px;
        margin: 24px auto;
        padding: 10px 16px 24px;
      }}
      .card {{
        background: linear-gradient(140deg, rgba(23, 26, 44, 0.96), rgba(21, 25, 47, 0.9));
        border: 1px solid rgba(120, 130, 190, 0.28);
        border-radius: 16px;
        box-shadow: 0 12px 45px rgba(0, 0, 0, 0.35);
        padding: 16px 16px 18px;
      }}
      h1 {{
        margin: 6px 4px 2px;
        font-size: clamp(1.45rem, 3.5vw, 2rem);
        letter-spacing: 0.2px;
      }}
      .meta {{
        margin: 2px 4px 16px;
        color: var(--muted);
        font-size: 0.95rem;
      }}
      .table-wrap {{
        width: 100%;
        overflow-x: auto;
      }}
      table {{
        width: 100%;
        table-layout: auto;
        border-collapse: collapse;
        border: 1px solid var(--line);
        border-radius: 12px;
        overflow: hidden;
        font-size: 0.86rem;
      }}
      thead tr {{
        background: linear-gradient(90deg, #24305e 0%, #1a2450 100%);
      }}
      th, td {{
        text-align: center;
        padding: 8px 7px;
        border-bottom: 1px solid var(--line);
        white-space: nowrap;
      }}
      th {{
        font-weight: 700;
        font-size: 0.92rem;
        letter-spacing: 0.2px;
      }}
      tbody tr {{
        transition: background-color 0.2s ease;
      }}
      tbody tr:nth-child(even) {{ background: rgba(255, 255, 255, 0.02); }}
      tbody tr:hover {{ background: rgba(111, 247, 255, 0.08); }}
      tr.rank-top td:first-child {{ background: rgba(255, 209, 102, 0.16); }}
      tr.rank-med td:first-child {{ background: rgba(111, 247, 255, 0.16); }}
      tr.rank-low td:first-child {{ background: rgba(141, 233, 105, 0.16); }}
      th.final-col,
      td.final-col {{
        color: #b5ffb9;
        font-weight: 700;
        box-shadow: inset 0 0 0 1px rgba(141, 233, 105, 0.45);
        background: rgba(141, 233, 105, 0.15);
      }}
      td:first-child {{ font-weight: 700; }}
      @media (max-width: 860px) {{
        .wrap {{ padding: 10px; margin: 12px auto; }}
        .card {{ padding: 12px 10px 14px; }}
        table, th, td {{ font-size: 0.79rem; }}
        th, td {{ padding: 8px 6px; }}
      }}
      @media print {{
        body {{
          background: white;
          color: black;
        }}
        .card {{
          box-shadow: none;
          border: none;
        }}
        table {{ border-color: #666; }}
        th, td {{ border-color: #666; color: #111; }}
        thead tr {{ background: #efefef; }}
      }}
    </style>
  </head>
  <body>
    <main class=\"wrap\">
      <section class=\"card\">
        <h1>MapTap Leaderboard</h1>
        <div class=\"meta\">Generated: {_generated_at()}</div>
        <div class=\"table-wrap\">
        <table aria-label=\"MapTap leaderboard\">
          <thead>
            <tr>
{header_cells}
            </tr>
          </thead>
          <tbody>
{body_rows}
          </tbody>
        </table>
        </div>
      </section>
    </main>
  </body>
</html>
"""
    path.write_text(content, encoding="utf-8")


def main() -> None:
    if not INPUT_FILE.exists():
        raise SystemExit(f"Input file not found: {INPUT_FILE} (expected TapMap.pdf)")

    records = parse_records(INPUT_FILE)
    leaderboard = build_leaderboard(records)
    leaderboard.to_csv(OUTPUT_FILE, index=False)
    wrote_xlsx = write_excel(leaderboard, EXCEL_OUTPUT_FILE)
    write_markdown(leaderboard, MARKDOWN_OUTPUT_FILE)
    write_html(leaderboard, HTML_OUTPUT_FILE)

    print(leaderboard.to_string(index=False))
    print(f"\nSaved: {OUTPUT_FILE}")
    print(f"Saved: {MARKDOWN_OUTPUT_FILE}")
    print(f"Saved: {HTML_OUTPUT_FILE}")
    if wrote_xlsx:
        print(f"Saved: {EXCEL_OUTPUT_FILE}")


if __name__ == "__main__":
    main()
