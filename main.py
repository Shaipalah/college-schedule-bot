import hashlib
import html
import io
import json
import os
import re
import sys
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse

import pandas as pd
import pdfplumber
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

PAGE_URL = "https://rgkript.ru/raspisanie-zanyatiy/"
DEFAULT_GROUP = "ТЭТ-22"
STATE_FILE = Path(__file__).resolve().parent / "schedule_state.json"

DAYS_OF_WEEK = [
    "Понедельник",
    "Вторник",
    "Среда",
    "Четверг",
    "Пятница",
    "Суббота",
]

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0.0.0 Safari/537.36"
    )
}


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def load_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {"xls": None, "pdfs": {}}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return {"xls": None, "pdfs": {}}
        return {
            "xls": data.get("xls"),
            "pdfs": data.get("pdfs") if isinstance(data.get("pdfs"), dict) else {},
        }
    except json.JSONDecodeError:
        return {"xls": None, "pdfs": {}}


def save_state(path: Path, state: dict[str, Any]) -> None:
    path.write_text(
        json.dumps(state, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def create_http_session() -> requests.Session:
    retry = Retry(
        total=4,
        connect=4,
        read=4,
        backoff_factor=2,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry)
    session = requests.Session()
    session.headers.update(REQUEST_HEADERS)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


HTTP_SESSION = create_http_session()


def fetch_bytes(url: str) -> bytes:
    resp = HTTP_SESSION.get(url, timeout=(20, 45))
    resp.raise_for_status()
    return resp.content


def parse_schedule_page(html_text: str) -> tuple[dict[str, str] | None, list[dict[str, str]]]:
    soup = BeautifulSoup(html_text, "html.parser")
    content = soup.select_one("#page-525 .page-content") or soup.select_one(".page-content") or soup

    # 1. General .xls link inside the table with "УЧЕБНЫЕ ГРУППЫ" / "ПРЕПОДАВАТЕЛИ" images
    xls_info: dict[str, str] | None = None
    xls_links = content.select('table a[href$=".xls"], table a[href$=".XLS"]')
    if not xls_links:
        xls_links = [
            a
            for a in content.find_all("a", href=True)
            if urlparse(a["href"]).path.lower().endswith(".xls")
        ]

    if xls_links:
        first_href = urljoin(PAGE_URL, xls_links[0]["href"].strip())
        heading_tag = None
        for h3 in content.find_all("h3"):
            h3_text = h3.get_text(" ", strip=True).lower()
            if "расписание учебных занятий" in h3_text:
                heading_tag = h3
                break
        if not heading_tag:
            for h3 in content.find_all("h3"):
                if "расписание" in h3.get_text(" ", strip=True).lower():
                    heading_tag = h3
                    break
        title = normalize_space(heading_tag.get_text(" ", strip=True)) if heading_tag else "Расписание занятий"
        xls_info = {"url": first_href, "title": title}

    # 2. Replacement .pdf links ("ЗАМЕНЫ на ...")
    pdf_links: list[dict[str, str]] = []
    seen_urls: set[str] = set()

    candidates = content.select('blockquote a[href$=".pdf"], blockquote a[href$=".PDF"]')
    if not candidates:
        candidates = [
            a
            for a in content.find_all("a", href=True)
            if urlparse(a["href"]).path.lower().endswith(".pdf")
        ]

    for a in candidates:
        href = urljoin(PAGE_URL, a["href"].strip())
        link_text = normalize_space(a.get_text())
        parent_text = normalize_space(a.parent.get_text(" ", strip=True)) if a.parent else ""
        title = link_text if "замен" in link_text.lower() else parent_text

        if "замен" not in title.lower():
            continue

        if href not in seen_urls:
            seen_urls.add(href)
            pdf_links.append({"url": href, "title": title or Path(urlparse(href).path).name})

    return xls_info, pdf_links


def format_slot_lines(lines: list[str]) -> str:
    if not lines:
        return ""
    formatted: list[str] = []
    for line in lines:
        if re.match(r"^а\.", line, re.IGNORECASE) and formatted:
            formatted[-1] = f"{formatted[-1]} ({line})"
        else:
            formatted.append(line)
    return " / ".join(formatted)


def parse_xls_for_group(xls_bytes: bytes, group_name: str) -> str | None:
    df = pd.read_excel(
        io.BytesIO(xls_bytes),
        sheet_name="Учебные группы",
        header=None,
        engine="xlrd",
        dtype=object,
    )

    target_norm = normalize_space(group_name).casefold()
    group_row: int | None = None

    # Group name is in column B (index 1)
    for row_idx in range(len(df)):
        val = df.iat[row_idx, 1]
        if pd.isna(val):
            continue
        if normalize_space(str(val)).casefold() == target_norm:
            group_row = row_idx
            break

    if group_row is None:
        return None

    # Find date/week header line above the group block
    week_header = ""
    for r in range(max(0, group_row - 6), group_row):
        for c in range(min(df.shape[1], 7)):
            val = df.iat[r, c]
            if pd.isna(val):
                continue
            text = normalize_space(str(val))
            if any(k in text.lower() for k in ("учебных занятий", "семестр", "недел")) or re.search(
                r"\d{2}\.\d{2}\.\d{4}", text
            ):
                week_header = text
                break
        if week_header:
            break

    # Days row is directly below the group row
    days_row = group_row + 1
    day_cols: list[tuple[int, str]] = []
    for c in range(1, min(df.shape[1], 8)):
        val = df.iat[days_row, c]
        if not pd.isna(val):
            day_text = normalize_space(str(val))
            if day_text:
                day_cols.append((c, day_text))

    if not day_cols:
        day_cols = [(i + 1, day) for i, day in enumerate(DAYS_OF_WEEK)]

    # 6 four-row slots starting after the days row
    slots_start = days_row + 1
    schedule_by_day: dict[str, list[tuple[str, str]]] = {day: [] for _, day in day_cols}

    for slot_idx in range(6):
        r0 = slots_start + slot_idx * 4
        if r0 >= len(df):
            break

        time_val = df.iat[r0, 0]
        time_str = normalize_space(str(time_val)) if not pd.isna(time_val) else f"{slot_idx + 1} пара"
        if not time_str:
            time_str = f"{slot_idx + 1} пара"

        for col_idx, day_name in day_cols:
            cell_lines: list[str] = []
            for offset in range(4):
                rr = r0 + offset
                if rr >= len(df) or col_idx >= df.shape[1]:
                    continue
                val = df.iat[rr, col_idx]
                if pd.isna(val):
                    continue
                text = normalize_space(str(val))
                if text:
                    cell_lines.append(text)

            entry = format_slot_lines(cell_lines)
            if entry:
                schedule_by_day[day_name].append((time_str, entry))

    lines_out: list[str] = [
        f"📅 <b>Расписание для группы {html.escape(group_name)}</b>"
    ]
    if week_header:
        lines_out.append(f"<i>{html.escape(week_header)}</i>")
    lines_out.append("")

    for _, day_name in day_cols:
        lessons = schedule_by_day.get(day_name, [])
        lines_out.append(f"<b>{html.escape(day_name)}</b>:")
        if not lessons:
            lines_out.append("  — нет пар")
        else:
            for time_str, entry in lessons:
                lines_out.append(f"  <code>{html.escape(time_str)}</code> — {html.escape(entry)}")
        lines_out.append("")

    return "\n".join(lines_out).strip()


def clean_table_cell(cell: Any) -> str:
    if cell is None:
        return ""
    return normalize_space(str(cell))


def parse_pdf_for_group(pdf_bytes: bytes, group_name: str) -> tuple[list[str], list[str]]:
    target_norm = normalize_space(group_name).casefold()
    group_pattern = re.compile(rf"(?<!\w){re.escape(group_name)}(?!\w)", re.IGNORECASE)

    headers_found: list[str] = []
    matches: list[str] = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            if not group_pattern.search(page_text):
                continue

            for raw_line in page_text.splitlines():
                line = normalize_space(raw_line)
                if ("замены" in line.lower() or re.search(r"\d{2}\.\d{2}\.\d{4}", line)) and line not in headers_found:
                    if not group_pattern.search(line):
                        headers_found.append(line)

            page_matches_before = len(matches)
            tables = page.extract_tables() or []
            for table in tables:
                for row in table:
                    if not row:
                        continue
                    cleaned = [clean_table_cell(c) for c in row]
                    if not any(cleaned):
                        continue
                    first_col = cleaned[0].casefold()
                    if first_col == target_norm or group_pattern.search(cleaned[0]):
                        parts = [p for p in cleaned[1:] if p]
                        if len(parts) >= 3:
                            pair, scheduled, *replacement = parts
                            formatted = f"• {pair}: {scheduled} → {' '.join(replacement)}"
                        elif parts:
                            formatted = f"• {' | '.join(parts)}"
                        else:
                            formatted = f"• {' | '.join(cleaned)}"
                        matches.append(formatted)

            # Fallback to raw text lines if table extraction missed anything on this page
            if len(matches) == page_matches_before:
                for raw_line in page_text.splitlines():
                    line = normalize_space(raw_line)
                    if group_pattern.search(line):
                        matches.append(f"• {line}")

    return headers_found, matches


def send_telegram_message(token: str, chat_id: str, text: str) -> None:
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    max_len = 4000
    chunks: list[str] = []
    remaining = text
    while len(remaining) > max_len:
        split_at = remaining.rfind("\n", 0, max_len)
        if split_at <= 0:
            split_at = max_len
        chunks.append(remaining[:split_at])
        remaining = remaining[split_at:].lstrip("\n")
    if remaining:
        chunks.append(remaining)

    for chunk in chunks:
        resp = HTTP_SESSION.post(
            url,
            data={
                "chat_id": chat_id,
                "text": chunk,
                "parse_mode": "HTML",
                "disable_web_page_preview": True,
            },
            timeout=30,
        )
        resp.raise_for_status()


def send_telegram_document(
    token: str,
    chat_id: str,
    filename: str,
    file_bytes: bytes,
    caption: str = "",
) -> None:
    url = f"https://api.telegram.org/bot{token}/sendDocument"
    data: dict[str, Any] = {"chat_id": chat_id}
    if caption:
        data["caption"] = caption[:1024]
    resp = HTTP_SESSION.post(
        url,
        data=data,
        files={"document": (filename, file_bytes, "application/pdf")},
        timeout=60,
    )
    resp.raise_for_status()


def main() -> int:
    token = os.environ.get("TELEGRAM_TOKEN", "").strip()
    chat_id = os.environ.get("TELEGRAM_CHAT_ID", "").strip()
    group_name = os.environ.get("GROUP_NAME", DEFAULT_GROUP).strip() or DEFAULT_GROUP

    if not token or not chat_id:
        print("Error: TELEGRAM_TOKEN and TELEGRAM_CHAT_ID must be set.", file=sys.stderr)
        return 1

    state = load_state(STATE_FILE)
    state_changed = False
    sent_anything = False

    print(f"Fetching {PAGE_URL} for group '{group_name}'...")
    try:
        page_resp = HTTP_SESSION.get(PAGE_URL, timeout=(20, 45))
        page_resp.raise_for_status()
    except requests.exceptions.RequestException as exc:
        print(
            f"Warning: could not reach {PAGE_URL} ({exc}). Skipping this check.",
            file=sys.stderr,
        )
        return 0

    if not page_resp.encoding or page_resp.encoding.lower() == "iso-8859-1":
        page_resp.encoding = page_resp.apparent_encoding or "utf-8"

    xls_info, pdf_links = parse_schedule_page(page_resp.text)
    print(f"Found XLS: {xls_info}")
    print(f"Found replacement PDFs: {len(pdf_links)}")

    # 1. Check general .xls schedule
    if xls_info:
        xls_url = xls_info["url"]
        try:
            xls_bytes = fetch_bytes(xls_url)
        except requests.exceptions.RequestException as exc:
            print(f"Warning: failed to download XLS {xls_url} ({exc}).", file=sys.stderr)
            xls_bytes = None

        if xls_bytes is not None:
            xls_hash = sha256_bytes(xls_bytes)
            prev_xls = state.get("xls") or {}

            if prev_xls.get("url") != xls_url or prev_xls.get("sha256") != xls_hash:
                print(f"XLS changed or new: {xls_url} (sha256={xls_hash[:12]}...)")
                schedule_msg = parse_xls_for_group(xls_bytes, group_name)
                if schedule_msg:
                    send_telegram_message(token, chat_id, schedule_msg)
                    sent_anything = True
                    print(f"Sent XLS schedule for group '{group_name}'.")
                else:
                    print(f"Group '{group_name}' not found in sheet 'Учебные группы'.")

                state["xls"] = {
                    "url": xls_url,
                    "sha256": xls_hash,
                    "title": xls_info["title"],
                }
                state_changed = True
            else:
                print("XLS schedule is unchanged.")

    # 2. Check replacement .pdf files
    pdfs_state: dict[str, Any] = state.setdefault("pdfs", {})
    for pdf_item in pdf_links:
        pdf_url = pdf_item["url"]
        pdf_title = pdf_item["title"]
        try:
            pdf_bytes = fetch_bytes(pdf_url)
        except requests.exceptions.RequestException as exc:
            print(f"Warning: failed to download PDF {pdf_url} ({exc}).", file=sys.stderr)
            continue

        pdf_hash = sha256_bytes(pdf_bytes)

        prev_pdf = pdfs_state.get(pdf_url) or {}
        if prev_pdf.get("sha256") != pdf_hash:
            print(f"PDF changed or new: {pdf_url} (sha256={pdf_hash[:12]}...)")
            headers_found, group_matches = parse_pdf_for_group(pdf_bytes, group_name)
            if group_matches:
                msg_lines = [
                    f"🔄 <b>{html.escape(pdf_title)}</b>",
                    f"Группа: <b>{html.escape(group_name)}</b>",
                ]
                if headers_found:
                    msg_lines.append(f"<i>{html.escape(headers_found[0])}</i>")
                msg_lines.append("")
                msg_lines.extend(html.escape(line) for line in group_matches)
                send_telegram_message(token, chat_id, "\n".join(msg_lines))

                filename = Path(urlparse(pdf_url).path).name or "zameny.pdf"
                send_telegram_document(token, chat_id, filename, pdf_bytes, caption=pdf_title)
                sent_anything = True
                print(f"Sent PDF replacements and file '{filename}' for group '{group_name}'.")
            else:
                print(f"No replacements for '{group_name}' in '{pdf_title}'.")

            pdfs_state[pdf_url] = {
                "sha256": pdf_hash,
                "title": pdf_title,
            }
            state_changed = True
        else:
            print(f"PDF '{pdf_title}' is unchanged.")

    if state_changed:
        save_state(STATE_FILE, state)
        print("Updated schedule_state.json.")

    if not sent_anything:
        print(f"No new schedule/replacement changes for '{group_name}'. Exiting 0.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
