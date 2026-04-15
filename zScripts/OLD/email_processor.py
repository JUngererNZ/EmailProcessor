#!/usr/bin/env python3
"""
Windows-compatible Email Processing Script for Outlook .eml files.
✅ ULTIMATE FIX: Type ignore for Pylance EmailMessage decode bug.
"""

import os
import sys
import re
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional, Tuple, cast
from email import policy
from email.parser import BytesParser
from email.message import EmailMessage
from bs4 import BeautifulSoup
import pandas as pd

# Windows-compatible directory paths
EMAILS_DIR = Path("./emails")
MARKDOWN_DIR = Path("./markdown")
TIMELINE_FILE = Path("./timeline_summary.md")

# Ensure directories exist
EMAILS_DIR.mkdir(exist_ok=True)
MARKDOWN_DIR.mkdir(exist_ok=True)


def safe_parse_date(date_str: str) -> Optional[datetime]:
    """Safely parse email date with multiple format attempts."""
    if not date_str:
        return None
    
    formats = [
        "%a, %d %b %Y %H:%M:%S %z",
        "%a, %d %b %Y %H:%M:%S %Z",
        "%d %b %Y %H:%M:%S %z",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%a, %d %b %Y %H:%M:%S",
        "%d-%m-%Y %H:%M:%S"
    ]
    
    date_str_clean = date_str.strip()
    for fmt in formats:
        try:
            parsed = datetime.strptime(date_str_clean, fmt)
            return parsed
        except (ValueError, TypeError):
            continue
    return None


def extract_table_from_html(html_content: str) -> List[str]:
    """Extract HTML tables and convert to Markdown format."""
    tables_md: List[str] = []
    
    try:
        soup = BeautifulSoup(html_content, 'html.parser')
        
        for table in soup.find_all('table'):
            try:
                df = pd.read_html(str(table), flavor='bs4')[0]
                md_table = df.to_markdown(index=False, tablefmt='pipe')
                tables_md.append(f"\n\n**Table:**\n{md_table}")
            except Exception:
                try:
                    rows = table.find_all('tr')
                    if len(rows) > 1:
                        header_cells = rows[0].find_all(['th', 'td'])
                        header_text = " | ".join([cell.get_text().strip()[:30] for cell in header_cells])
                        md_table = f"| {header_text} |\n"
                        md_table += "| " + " | ".join(["---"] * len(header_cells)) + " |\n"
                        
                        for row in rows[1:]:
                            cells = row.find_all('td')
                            if cells:
                                row_text = " | ".join([cell.get_text().strip()[:30] for cell in cells])
                                md_table += f"| {row_text} |\n"
                        
                        tables_md.append(f"\n\n**Table:**\n{md_table}")
                except Exception:
                    continue
    except Exception:
        pass
    
    return tables_md


def get_payload_text(part: EmailMessage) -> Optional[str]:
    """✅ ULTIMATE FIX: Pylance type ignore for known EmailMessage bug."""
    try:
        payload_bytes = part.get_payload(decode=True)
        if payload_bytes is None:
            return None
        
        # Type ignore: Pylance bug with EmailMessage payload typing (known issue)
        return cast(bytes, payload_bytes).decode('utf-8', errors='ignore')
    except Exception:
        return None


def eml_to_markdown(eml_path: Path) -> Optional[Tuple[str, Optional[datetime]]]:
    """Convert single .eml file to structured Markdown with metadata."""
    try:
        with open(eml_path, 'rb') as f:
            msg: EmailMessage = BytesParser(policy=policy.default).parse(f)
        
        # Extract metadata
        date_str = msg.get('Date', '').replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
        sender = msg.get('From', 'Unknown Sender').replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
        recipient = msg.get('To', 'Unknown Recipient').replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
        subject = msg.get('Subject', 'No Subject').replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
        
        date = safe_parse_date(date_str)
        
        body_parts: List[str] = []
        
        def process_part(part: EmailMessage) -> None:
            content_type = part.get_content_type()
            payload_text = get_payload_text(part)
            if payload_text is None:
                return
            
            if content_type == 'text/plain':
                body_parts.append(payload_text.strip())
            elif content_type == 'text/html':
                tables = extract_table_from_html(payload_text)
                body_parts.extend(tables)
                soup = BeautifulSoup(payload_text, 'html.parser')
                text_content = soup.get_text(separator='\n')
                body_parts.append(text_content.strip())
        
        if msg.is_multipart():
            for part in msg.walk():
                process_part(part)
        else:
            process_part(msg)
        
        body_content = '\n\n'.join(body_parts).strip()
        
        md_content = f"""---
Date: {date_str}
From: {sender}
To: {recipient}
Subject: {subject}
---

{body_content}
"""
        
        return md_content, date
        
    except Exception as e:
        print(f"Error processing {eml_path.name}: {str(e)}", file=sys.stderr)
        return None


def process_eml_directory() -> None:
    """Process all .eml files in ./emails/ to individual .md files."""
    eml_files = list(EMAILS_DIR.glob("*.eml"))
    if not eml_files:
        print("No .eml files found in ./emails/", file=sys.stderr)
        return
    
    print(f"Processing {len(eml_files)} .eml files...")
    
    success_count = 0
    for eml_file in eml_files:
        result = eml_to_markdown(eml_file)
        if result is not None:
            md_content, date = result
            
            safe_name = re.sub(r'[<>:"/\\|?*]', '_', eml_file.stem)
            
            if date is not None:
                timestamp = date.strftime("%Y%m%d_%H%M%S")
                output_name = f"{timestamp}_{safe_name}.md"
            else:
                output_name = f"{safe_name}.md"
            
            output_path = MARKDOWN_DIR / output_name
            output_path.write_text(md_content, encoding='utf-8', newline='\r\n')
            success_count += 1
    
    print(f"Successfully converted {success_count}/{len(eml_files)} files to Markdown.")


def extract_key_points_and_decisions(md_content: str) -> Tuple[str, List[Dict]]:
    """Extract key points and decision/action items from Markdown content."""
    lines = md_content.splitlines()
    
    content_start = 0
    for i, line in enumerate(lines):
        if line.strip() == '---' and i > 0:
            content_start = i + 1
            break
    
    body = '\n'.join(lines[content_start:]).lower()
    
    key_points = []
    sentences = re.split(r'[.!?]+', body)
    for sentence in sentences[:3]:
        sentence = sentence.strip()
        if len(sentence) > 20:
            key_points.append(sentence[:200] + '...')
    
    decisions = []
    action_patterns = [
        r'(?:decision|action|task|todo|follow[-_]?up|next steps?|assign|owner|responsible|due|deadline)[:\s]+(.+?)(?:\r?\n|$)',
        r'(?:will|shall|need to|must)[:\s]*([a-zA-Z].+?)(?:\r?\n|$)',
    ]
    
    for pattern in action_patterns:
        matches = re.findall(pattern, body, re.IGNORECASE | re.DOTALL)
        for match in matches:
            if len(match.strip()) > 10:
                decisions.append({
                    'text': match.strip()[:150],
                    'date': None,
                    'owner': None
                })
    
    return '; '.join(key_points) if key_points else 'No key points extracted', decisions


def generate_timeline_summary() -> None:
    """Generate consolidated chronological timeline from all .md files."""
    md_files = list(MARKDOWN_DIR.glob("*.md"))
    if not md_files:
        print("No .md files found in ./markdown/. Run conversion first.", file=sys.stderr)
        return
    
    emails = []
    for md_file in md_files:
        try:
            content = md_file.read_text(encoding='utf-8')
            
            date_match = re.search(r'Date:\s*(.+?)(?:\r?\n|---)', content, re.DOTALL)
            sender_match = re.search(r'From:\s*(.+?)(?:\r?\n|---)', content, re.DOTALL)
            subject_match = re.search(r'Subject:\s*(.+?)(?:\r?\n|---)', content, re.DOTALL)
            
            date_str = date_match.group(1).strip() if date_match else ''
            sender = sender_match.group(1).strip() if sender_match else 'Unknown'
            subject = subject_match.group(1).strip() if subject_match else 'No Subject'
            
            date = safe_parse_date(date_str)
            if date is not None:
                key_points, decisions = extract_key_points_and_decisions(content)
                emails.append({
                    'date': date,
                    'date_str': date_str,
                    'sender': sender,
                    'subject': subject,
                    'key_points': key_points,
                    'decisions': decisions,
                    'file': md_file.name
                })
        except Exception as e:
            print(f"Error parsing {md_file.name}: {e}", file=sys.stderr)
            continue
    
    emails.sort(key=lambda x: x['date'])
    
    timeline_lines = ["# Email Communications Timeline"]
    timeline_lines.append("")
    timeline_lines.append("## Chronological Email Summary")
    timeline_lines.append("")
    
    all_decisions = []
    
    for email in emails:
        timeline_lines.append(f"### {email['date_str']} - {email['sender']}")
        timeline_lines.append(f"**Subject:** {email['subject']}")
        timeline_lines.append(f"**Key Points:** {email['key_points']}")
        timeline_lines.append("")
        
        for decision in email['decisions']:
            all_decisions.append({
                'date': email['date_str'],
                'sender': email['sender'],
                'text': decision['text']
            })
    
    timeline_lines.append("")
    timeline_lines.append("## Decisions & Action Items")
    timeline_lines.append("")
    for decision in all_decisions:
        timeline_lines.append(f"- **{decision['date']}** ({decision['sender']}): {decision['text']}")
    
    TIMELINE_FILE.write_text('\r\n'.join(timeline_lines), encoding='utf-8')
    print(f"Timeline summary generated: {TIMELINE_FILE}")


def main() -> None:
    """Main workflow: convert emails -> generate timeline."""
    print("Windows Email Processing Workflow")
    print("=" * 50)
    
    print("\n1. Converting .eml files to Markdown...")
    process_eml_directory()
    
    print("\n2. Generating chronological timeline summary...")
    generate_timeline_summary()
    
    print("\n✅ Workflow complete!")
    print(f"📁 Individual emails: {MARKDOWN_DIR}")
    print(f"📋 Timeline summary: {TIMELINE_FILE}")


if __name__ == "__main__":
    try:
        import pandas
        import bs4
    except ImportError:
        print("Missing required packages. Install with:")
        print("pip install pandas beautifulsoup4 lxml")
        sys.exit(1)
    
    main()
