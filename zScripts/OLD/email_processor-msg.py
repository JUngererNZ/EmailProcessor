#!/usr/bin/env python3
"""
Windows-compatible Email Processing Script for Outlook .msg AND .eml files.
✅ ULTIMATE FIX: ZERO PYLANCE WARNINGS - Simplified table extraction + safe imports.
"""

import os
import sys
import re
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Any
import pandas as pd

# Windows-compatible directory paths
EMAILS_DIR = Path("./emails")
MARKDOWN_DIR = Path("./markdown")
TIMELINE_FILE = Path("./timeline_summary.md")

# Ensure directories exist
EMAILS_DIR.mkdir(exist_ok=True)
MARKDOWN_DIR.mkdir(exist_ok=True)


def check_dependencies():
    """Check required packages."""
    try:
        import bs4
        import extract_msg
    except ImportError:
        print("Install dependencies:")
        print("pip install extract-msg pandas beautifulsoup4 lxml")
        sys.exit(1)


def safe_parse_date(date_str: str) -> Optional[datetime]:
    """Safely parse email date."""
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
            return datetime.strptime(date_str_clean, fmt)
        except (ValueError, TypeError):
            continue
    return None


def extract_tables_pandas_only(html_content: str) -> List[str]:
    """✅ SIMPLIFIED: Pandas-only table extraction (Pylance safe)."""
    tables_md: List[str] = []
    
    try:
        # Pandas handles HTML parsing internally - NO BeautifulSoup needed
        dfs = pd.read_html(html_content)
        for i, df in enumerate(dfs):
            if not df.empty:
                md_table = df.to_markdown(index=False, tablefmt='pipe')
                tables_md.append(f"\n\n**Table {i+1}:**\n{md_table}")
    except Exception:
        pass
    
    return tables_md


def process_eml_file(eml_path: Path) -> Optional[Tuple[str, Optional[datetime]]]:
    """Process .eml file - Pylance safe."""
    try:
        from email import policy
        from email.parser import BytesParser
        from email.message import EmailMessage
        
        with open(eml_path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)
        
        date_str = msg.get('Date', '').replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
        sender = msg.get('From', 'Unknown').replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
        recipient = msg.get('To', 'Unknown').replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
        subject = msg.get('Subject', 'No Subject').replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
        
        date = safe_parse_date(date_str)
        body_parts: List[str] = []
        
        # Simple text/HTML extraction
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                try:
                    payload = part.get_payload(decode=True)
                    if payload:
                        text = payload.decode('utf-8', errors='ignore')
                        if content_type == 'text/html':
                            tables = extract_tables_pandas_only(text)
                            body_parts.extend(tables)
                            body_parts.append(text[:2000])  # Truncate long HTML
                        else:
                            body_parts.append(text)
                except:
                    continue
        else:
            try:
                payload = msg.get_payload(decode=True)
                if payload:
                    text = payload.decode('utf-8', errors='ignore')
                    tables = extract_tables_pandas_only(text)
                    body_parts.extend(tables)
                    body_parts.append(text)
            except:
                pass
        
        body_content = '\n\n'.join(body_parts).strip()
        
        md_content = f"""---
Date: {date_str}
From: {sender}
To: {recipient}
Subject: {subject}
Source: EML
---

{body_content}
"""
        return md_content, date
        
    except Exception as e:
        print(f"Error processing EML {eml_path.name}: {str(e)}", file=sys.stderr)
        return None


def process_msg_file(msg_path: Path) -> Optional[Tuple[str, Optional[datetime]]]:
    """Process .msg file - Pylance safe with getattr."""
    try:
        import extract_msg
        
        with extract_msg.Message(msg_path) as msg:
            # Safe attribute access
            date_str = str(msg.date) if hasattr(msg, 'date') and msg.date else ''
            sender = str(getattr(msg, 'sender', 'Unknown'))
            recipient = str(getattr(msg, 'to', 'Unknown'))
            subject = str(getattr(msg, 'subject', 'No Subject'))
            
            date_str = date_str.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
            sender = sender.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
            recipient = recipient.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
            subject = subject.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
            
            date = safe_parse_date(date_str)
            body_parts: List[str] = []
            
            # Safe body extraction
            if hasattr(msg, 'htmlBody') and msg.htmlBody:
                html_content = str(msg.htmlBody)
                tables = extract_tables_pandas_only(html_content)
                body_parts.extend(tables)
                body_parts.append(html_content[:2000])
            elif hasattr(msg, 'body') and msg.body:
                body_content = str(msg.body)
                tables = extract_tables_pandas_only(body_content)
                body_parts.extend(tables)
                body_parts.append(body_content)
            
            content = '\n\n'.join(body_parts).strip()
            
            md_content = f"""---
Date: {date_str}
From: {sender}
To: {recipient}
Subject: {subject}
Source: MSG
---

{content}
"""
            return md_content, date
            
    except Exception as e:
        print(f"Error processing MSG {msg_path.name}: {str(e)}", file=sys.stderr)
        return None


def email_to_markdown(email_path: Path) -> Optional[Tuple[str, Optional[datetime]]]:
    """Unified converter - Pylance safe."""
    if email_path.suffix.lower() == '.eml':
        return process_eml_file(email_path)
    elif email_path.suffix.lower() == '.msg':
        return process_msg_file(email_path)
    else:
        print(f"Unsupported format: {email_path.name}", file=sys.stderr)
        return None


def process_email_directory() -> None:
    """Process all email files."""
    email_files = list(EMAILS_DIR.glob("*.eml")) + list(EMAILS_DIR.glob("*.msg"))
    if not email_files:
        print("No .eml or .msg files found in ./emails/", file=sys.stderr)
        return
    
    print(f"Processing {len(email_files)} email files...")
    
    success_count = 0
    for email_file in email_files:
        result = email_to_markdown(email_file)
        if result:
            md_content, date = result
            
            safe_name = re.sub(r'[<>:"/\\|?*]', '_', email_file.stem)
            prefix = email_file.suffix.lower().replace('.', '')
            
            timestamp = date.strftime("%Y%m%d_%H%M%S") if date else "unknown"
            output_name = f"{timestamp}_{prefix}_{safe_name}.md"
            
            output_path = MARKDOWN_DIR / output_name
            output_path.write_text(md_content, encoding='utf-8', newline='\r\n')
            success_count += 1
    
    print(f"Successfully converted {success_count}/{len(email_files)} files.")


def extract_key_points_and_decisions(md_content: str) -> Tuple[str, List[Dict]]:
    """Extract key points and decisions."""
    lines = md_content.splitlines()
    
    # Skip YAML frontmatter
    content_start = 0
    for i, line in enumerate(lines):
        if line.strip() == '---' and i > 0:
            content_start = i + 1
            break
    
    body = '\n'.join(lines[content_start:]).lower()
    
    # Key points (first 3 sentences)
    key_points = []
    sentences = re.split(r'[.!?]+', body)
    for sentence in sentences[:3]:
        sentence = sentence.strip()
        if len(sentence) > 20:
            key_points.append(sentence[:200] + '...')
    
    # Decisions (action keywords)
    decisions = []
    action_patterns = [
        r'(decision|action|task|todo|follow[-_]?up|next steps?|assign|owner|responsible|due|deadline)[:\s]+(.+?)(?:\r?\n|$)',
        r'(will|shall|need to|must)[:\s]*([a-z].+?)(?:\r?\n|$)'
    ]
    
    for pattern in action_patterns:
        matches = re.findall(pattern, body, re.IGNORECASE | re.DOTALL)
        for match in matches:
            text = match[-1].strip()
            if len(text) > 10:
                decisions.append({'text': text[:150], 'date': None, 'owner': None})
    
    return '; '.join(key_points) if key_points else 'No key points', decisions


def generate_timeline_summary() -> None:
    """Generate timeline from .md files."""
    md_files = list(MARKDOWN_DIR.glob("*.md"))
    if not md_files:
        print("No .md files. Run conversion first.", file=sys.stderr)
        return
    
    emails = []
    for md_file in md_files:
        try:
            content = md_file.read_text(encoding='utf-8')
            date_match = re.search(r'Date:\s*(.+?)(?:\r?\n|---)', content, re.DOTALL)
            sender_match = re.search(r'From:\s*(.+?)(?:\r?\n|---)', content, re.DOTALL)
            subject_match = re.search(r'Subject:\s*(.+?)(?:\r?\n|---)', content, re.DOTALL)
            
            if date_match:
                date_str = date_match.group(1).strip()
                sender = sender_match.group(1).strip() if sender_match else 'Unknown'
                subject = subject_match.group(1).strip() if subject_match else 'No Subject'
                
                date = safe_parse_date(date_str)
                if date:
                    key_points, decisions = extract_key_points_and_decisions(content)
                    emails.append({
                        'date': date, 'date_str': date_str, 'sender': sender,
                        'subject': subject, 'key_points': key_points, 'decisions': decisions
                    })
        except Exception:
            continue
    
    if not emails:
        print("No valid emails found for timeline.", file=sys.stderr)
        return
    
    emails.sort(key=lambda x: x['date'])
    
    timeline_lines = ["# Email Communications Timeline", "", "## Chronological Summary", ""]
    all_decisions = []
    
    for email in emails:
        timeline_lines.extend([
            f"### {email['date_str']} - {email['sender']}",
            f"**Subject:** {email['subject']}",
            f"**Key Points:** {email['key_points']}",
            ""
        ])
        all_decisions.extend([{
            'date': email['date_str'], 'sender': email['sender'], 'text': d['text']
        } for d in email['decisions']])
    
    timeline_lines.extend(["", "## Decisions & Actions", ""])
    for decision in all_decisions:
        timeline_lines.append(f"- **{decision['date']}** ({decision['sender']}): {decision['text']}")
    
    TIMELINE_FILE.write_text('\r\n'.join(timeline_lines), encoding='utf-8')
    print(f"Timeline created: {TIMELINE_FILE}")


def main() -> None:
    """Main workflow."""
    check_dependencies()
    
    print("Email Processor (.MSG + .EML → Markdown Timeline)")
    print("=" * 55)
    
    print("\n1. Converting emails...")
    process_email_directory()
    
    print("\n2. Creating timeline...")
    generate_timeline_summary()
    
    print("\n✅ COMPLETE!")
    print(f"📁 Emails: {MARKDOWN_DIR}")
    print(f"📋 Timeline: {TIMELINE_FILE}")


if __name__ == "__main__":
    main()
