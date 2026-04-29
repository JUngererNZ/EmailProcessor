#!/usr/bin/env python3
"""
FML Email Pipeline
Stages:
  1. .msg files  →  individual markdown files
  2. Markdown    →  chronological timeline summary
  3. Markdown    →  populated FML checklist
  4. All markdown → consolidated summary (Ollama / qwen2.5)
"""

import os
import re
import sys
import datetime
from pathlib import Path

import extract_msg          # pip install extract-msg
from bs4 import BeautifulSoup  # pip install beautifulsoup4
import html2text            # pip install html2text
import ollama               # pip install ollama


# ---------------------------------------------------------------------------
# Stage 1 helpers: .msg → markdown
# ---------------------------------------------------------------------------

def extract_tables_to_md(html_content):
    """Convert HTML tables found in email body to markdown tables."""
    soup = BeautifulSoup(html_content, 'html.parser')
    tables_md = []

    for table in soup.find_all('table'):
        rows = []
        all_rows = table.find_all('tr')
        if not all_rows:
            continue

        header_row = all_rows[0].find_all(['th', 'td'])
        if header_row:
            header = [th.get_text(strip=True) for th in header_row]
            rows.append('| ' + ' | '.join(header) + ' |')
            rows.append('| ' + ' | '.join(['---'] * len(header)) + ' |')

        for row in all_rows[1:]:
            cells = row.find_all(['td', 'th'])
            if cells:
                rows.append('| ' + ' | '.join(c.get_text(strip=True) for c in cells) + ' |')

        if rows:
            tables_md.append('\n'.join(rows))

    return '\n\n'.join(tables_md)


def flatten_html_body(html_body):
    """Strip nested reply separators and convert HTML to markdown."""
    if not html_body:
        return ""

    soup = BeautifulSoup(html_body, 'html.parser')

    for div in soup.find_all('div', string=re.compile(
            r'(From:|Original Message:|Forwarded message:)', re.I)):
        div.decompose()

    h = html2text.HTML2Text()
    h.ignore_links = False
    h.ignore_images = True
    h.body_width = 0
    return h.handle(str(soup)).strip()


def convert_msg_to_md(msg_file_path, output_dir):
    """Convert a single .msg file to a .md file; return metadata dict or None."""
    try:
        with extract_msg.Message(msg_file_path) as msg:
            sender   = msg.sender  or "Unknown Sender"
            subject  = msg.subject or "No Subject"
            date_str = msg.date.strftime("%a, %d %b %Y %H:%M") if msg.date else "Unknown Date"

            html_body    = msg.htmlBody if hasattr(msg, 'htmlBody') and msg.htmlBody else ""
            body_content = flatten_html_body(html_body)
            tables_md    = extract_tables_to_md(html_body)

            key_points = []
            if tables_md:
                key_points.append("Contains task tables with assignments")
            action_keywords = ['action', 'task', 'assign', 'due', 'deadline', 'complete', 'review']
            if any(kw in body_content.lower() for kw in action_keywords):
                key_points.append("Action items mentioned")

            safe_subject = re.sub(r'[^\w\-_\.]', '_', subject)[:50]
            md_filename  = f"{safe_subject}_{os.path.splitext(os.path.basename(msg_file_path))[0]}.md"
            md_path      = output_dir / md_filename

            md_content = f"""# Email: {subject}

**From:** [{sender}](mailto:{sender.split('@')[0] if '@' in sender else sender})
**Date:** {date_str}
**Subject:** {subject}

## Content
{body_content}

{tables_md}

## Key Points
"""
            if key_points:
                md_content += '\n'.join(f"- {p}" for p in key_points)
            else:
                md_content += "- Review email content for details"

            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(md_content)

            print(f"  ✓ {msg_file_path.name} → {md_path.name}")

            return {
                'date':     msg.date,
                'date_str': date_str,
                'sender':   sender,
                'subject':  subject,
                'md_path':  str(md_path),
                'key_points': key_points,
                'body':     body_content,
            }

    except Exception as e:
        print(f"  ✗ Error processing {msg_file_path}: {e}")
        return None


# ---------------------------------------------------------------------------
# Stage 2: Timeline summary (rule-based, no LLM)
# ---------------------------------------------------------------------------

def generate_timeline_summary(emails_data, output_path):
    """Write a chronological timeline markdown from processed email metadata."""
    emails_data = sorted([e for e in emails_data if e], key=lambda x: x['date'])

    chrono = "## Chronological Email Summary\n\n"
    for email in emails_data:
        sender_link = (
            f"[{email['sender']}](mailto:{email['sender'].split('@')[0]})"
            if '@' in email['sender'] else email['sender']
        )
        chrono += f"### {email['date_str']} — {sender_link}\n"
        chrono += f"**Subject:** {email['subject']}\n**Key Points:**\n"
        for point in (email['key_points'] or ["Review email for details"]):
            chrono += f"- {point}\n"
        chrono += "\n"

    exec_summary = "## Executive Summary\n\n"
    if emails_data:
        latest = max(emails_data, key=lambda x: x['date'])
        exec_summary += (
            f"Thread spans {emails_data[0]['date_str']} → {latest['date_str']}. "
            f"{len(emails_data)} email(s) processed.\n\n"
            "Current status: Ongoing shipment processing — see chronological summary above.\n\n"
            "Next steps: Monitor for updates to ensure timely completion."
        )
    else:
        exec_summary += "No email data available."

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(f"# Email Communications Timeline\n\n{chrono}\n{exec_summary}\n")

    print(f"  ✓ Timeline → {output_path.name}")


# ---------------------------------------------------------------------------
# Stage 3: Checklist extraction and population
# ---------------------------------------------------------------------------

def extract_checklist_data_from_markdown(md_path):
    """Parse shipment fields and checklist status from a processed markdown file."""
    try:
        with open(md_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        print(f"  ✗ Could not read {md_path}: {e}")
        return {}, {}

    shipment_data  = {}
    checklist_items = {}

    # Field name → dict key mapping
    field_map = {
        'File Ref':                  'file_ref',
        'Client Ref':                'client_ref',
        'Consignee':                 'consignee',
        'Description':               'description',
        'PIN No':                    'pin_no',
        'Serial No':                 'serial_no',
        'Vessel':                    'vessel',
        'Voy':                       'voy',
        'Bill No':                   'bill_no',
        'ETA':                       'eta',
        'Tariff Code, Weight & Cube':'tariff_code',
        'Container #':               'container_no',
        'TRANSPORTER':               'transporter',
    }

    checklist_map = {
        'TRANSPORT INSTRUCTION': 'transport_instruction',
        'X-HAUL INSTRUCTION':    'x_haul_instruction',
        'ANF / SWB':             'anf_swb',
        'LINE INVOICE':          'line_invoice',
        'CUSTOMS ENTRY (WE/RIT)':'customs_entry',
    }

    for line in content.split('\n'):
        line = line.strip()
        if '|' not in line:
            continue

        # Shipment fields
        for label, key in field_map.items():
            if label + ':' in line:
                parts = line.split('|')
                for part in parts:
                    if label + ':' in part:
                        shipment_data[key] = part.replace(label + ':', '').strip()
                        break

        # Checklist items
        for label, key in checklist_map.items():
            if f'| {label} |' in line or f'| {label}' in line:
                parts = line.split('|')
                if len(parts) >= 5:
                    checklist_items[key] = {
                        'drafted':   '✅' in parts[2] or 'X' in parts[2],
                        'completed': '✅' in parts[3] or 'X' in parts[3],
                        'comments':  parts[4].strip() if len(parts) > 4 else '',
                    }

    return shipment_data, checklist_items


def generate_populated_checklist(shipment_data, checklist_items, output_path):
    """Write an FML checklist markdown populated with extracted shipment data."""

    def tick(key, field):
        return '✅' if checklist_items.get(key, {}).get(field) else '⬜'

    def comment(key):
        return checklist_items.get(key, {}).get('comments', '')

    def val(key):
        return shipment_data.get(key, '')

    content = f"""# FML-CHECKLIST — {val('file_ref') or 'Unknown'}

## Shipment Details
| Field | Value |
|-------|-------|
| File Ref | {val('file_ref')} |
| Client Ref | {val('client_ref')} |
| Consignee | {val('consignee')} |
| Description | {val('description')} |
| PIN No | {val('pin_no')} |
| Serial No | {val('serial_no')} |
| Vessel | {val('vessel')} |
| Voy | {val('voy')} |
| Bill No | {val('bill_no')} |
| ETA | {val('eta')} |
| Tariff Code, Weight & Cube | {val('tariff_code')} |
| Container # | {val('container_no')} |
| TRANSPORTER | {val('transporter')} |

## Checklist
| Task | Drafted | Completed | Comments |
|------|---------|-----------|----------|
| TRANSPORT INSTRUCTION  | {tick('transport_instruction','drafted')} | {tick('transport_instruction','completed')} | {comment('transport_instruction')} |
| X-HAUL INSTRUCTION     | {tick('x_haul_instruction','drafted')}   | {tick('x_haul_instruction','completed')}   | {comment('x_haul_instruction')} |
| ANF / SWB              | {tick('anf_swb','drafted')}              | {tick('anf_swb','completed')}              | {comment('anf_swb')} |
| LINE INVOICE           | {tick('line_invoice','drafted')}         | {tick('line_invoice','completed')}         | {comment('line_invoice')} |
| CUSTOMS ENTRY (WE/RIT) | {tick('customs_entry','drafted')}        | {tick('customs_entry','completed')}        | {comment('customs_entry')} |
| CARGO DUES             | ⬜ | ⬜ | |
| RELEASE LETTER         | ⬜ | ⬜ | |
| BV INSPECTION          | ⬜ | ⬜ | |
| GRN / XE               | ⬜ | ⬜ | |
| FINAL DOCUMENTATION    | ⬜ | ⬜ | |

*Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*
"""
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)

    print(f"  ✓ Checklist → {output_path.name}")


# ---------------------------------------------------------------------------
# Stage 4: Consolidated summary (Ollama / qwen2.5)
# ---------------------------------------------------------------------------

OLLAMA_MODEL   = "qwen2.5"
OLLAMA_TIMEOUT = 120  # seconds — increase if summaries are timing out on long files

def _summarise_with_ollama(content: str) -> str:
    """Call Ollama to summarise content; returns summary string or error message."""
    prompt = f"Summarize the following markdown content concisely in two sentences:\n\n{content[:2000]}"
    try:
        client   = ollama.Client(timeout=OLLAMA_TIMEOUT)
        response = client.chat(
            model=OLLAMA_MODEL,
            messages=[{"role": "user", "content": prompt}],
        )
        return response["message"]["content"].strip()
    except ollama.ResponseError as e:
        return f"[Ollama error: {e}]"
    except Exception as e:
        return f"[Summary unavailable: {e}]"


def generate_consolidated_summary(markdown_dir, output_path, folder_name):
    """
    Summarise every .md in markdown_dir (excluding timeline/checklist),
    then write a single consolidated file: latest full content + all summaries.
    """
    md_files = sorted(
        [f for f in markdown_dir.glob('*.md')
         if not f.name.startswith('timeline_') and not f.name.startswith('FML-CHECKLIST')],
        key=lambda f: f.stat().st_mtime,
        reverse=True,
    )

    if not md_files:
        print("  ⚠  No source markdown files found for consolidation.")
        return

    latest_content = ""
    summaries = []

    for i, md_file in enumerate(md_files):
        try:
            with open(md_file, 'r', encoding='utf-8-sig') as f:
                content = f.read()
        except UnicodeDecodeError:
            with open(md_file, 'r', encoding='utf-16') as f:
                content = f.read()

        if i == 0:
            latest_content = content

        print(f"  Summarising {md_file.name} ...")
        summary = _summarise_with_ollama(content)
        summaries.append(f"### Summary of {md_file.name}\n{summary}\n")

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(latest_content)
        f.write("\n\n---\n## File Summaries\n\n")
        f.write("\n".join(summaries))

    print(f"  ✓ Consolidated summary → {output_path.name}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) > 1:
        base_dir = Path(sys.argv[1]).resolve()
    else:
        cwd = Path.cwd()
        base_dir = cwd if (cwd / "emails").exists() else Path(__file__).resolve().parent

    emails_dir   = base_dir / "emails"
    markdown_dir = base_dir / "markdown"
    folder_name  = base_dir.name

    if not emails_dir.exists():
        print(f"✗ emails/ directory not found in {base_dir}")
        return

    markdown_dir.mkdir(parents=True, exist_ok=True)

    timeline_path  = markdown_dir / f"timeline_{folder_name}_summary.md"
    checklist_path = markdown_dir / f"FML-CHECKLIST-{folder_name}.md"
    date_str       = datetime.datetime.now().strftime('%Y%m%d')
    consolidated_path = base_dir / f"{folder_name}_{date_str}.md"

    # --- Stage 1 ---
    print("\n🔄 Stage 1: Converting .msg files to markdown ...")
    emails_data = []
    for msg_file in emails_dir.glob("*.msg"):
        result = convert_msg_to_md(msg_file, markdown_dir)
        if result:
            emails_data.append(result)

    if not emails_data:
        print("✗ No valid emails processed.")
        return

    # --- Stage 2 ---
    print("\n📅 Stage 2: Generating timeline ...")
    generate_timeline_summary(emails_data, timeline_path)

    # --- Stage 3 ---
    print("\n📋 Stage 3: Generating checklist ...")
    md_files = [
        f for f in markdown_dir.glob("*.md")
        if not f.name.startswith("timeline_") and not f.name.startswith("FML-CHECKLIST")
    ]

    source_file = None
    for md_file in md_files:
        try:
            with open(md_file, 'r', encoding='utf-8') as f:
                text = f.read()
            if any(kw in text for kw in ['File Ref:', 'Client Ref:', 'Consignee:', 'Description:']):
                source_file = md_file
                break
        except Exception:
            continue

    if source_file is None and md_files:
        source_file = max(md_files, key=lambda f: f.stat().st_mtime)

    if source_file:
        print(f"  Using shipment data from: {source_file.name}")
        shipment_data, checklist_items = extract_checklist_data_from_markdown(source_file)
        generate_populated_checklist(shipment_data, checklist_items, checklist_path)
    else:
        print("  ⚠  No markdown source found — skipping checklist.")

    # --- Stage 4 ---
    print("\n🤖 Stage 4: Generating consolidated summary ...")
    generate_consolidated_summary(markdown_dir, consolidated_path, folder_name)

    print(f"""
✅ Pipeline complete — {len(emails_data)} email(s) processed.
   Individual MD files : {markdown_dir}
   Timeline            : {timeline_path}
   Checklist           : {checklist_path}
   Consolidated output : {consolidated_path}
""")


if __name__ == "__main__":
    main()