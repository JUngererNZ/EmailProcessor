"""
Outlook .msg to Markdown Converter with Timeline Generator

This script converts .msg email files to markdown format and creates
a chronological timeline summary with action items and decisions.

Requirements:
    pip install extract-msg

Usage:
    python msg_converter.py
"""

import os
import re
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Tuple, Optional
import glob

try:
    import extract_msg
except ImportError:
    print("Error: extract-msg library not found.")
    print("Please install it using: pip install extract-msg")
    exit(1)


class EmailData:
    """Container for parsed email data"""
    def __init__(self, sender: str, date: datetime, subject: str, 
                 body: str, filename: str):
        self.sender = sender
        self.date = date
        self.subject = subject
        self.body = body
        self.filename = filename
        self.action_items = []
        self.key_points = []


def setup_directories():
    """Create necessary directory structure"""
    dirs = ['project/emails', 'project/markdown']
    for directory in dirs:
        Path(directory).mkdir(parents=True, exist_ok=True)
    print("✓ Directory structure verified")


def extract_tables_from_text(text: str) -> List[str]:
    """
    Extract table-like structures from email text and convert to markdown tables
    """
    tables = []
    lines = text.split('\n')
    
    # Look for patterns that suggest tabular data
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # Check if line contains tab-separated or pipe-separated values
        if '\t' in line or '|' in line:
            table_lines = [line]
            j = i + 1
            
            # Collect consecutive table rows
            while j < len(lines):
                next_line = lines[j].strip()
                if '\t' in next_line or '|' in next_line or not next_line:
                    if next_line:
                        table_lines.append(next_line)
                    j += 1
                else:
                    break
            
            if len(table_lines) >= 2:  # At least header + one row
                md_table = convert_to_markdown_table(table_lines)
                if md_table:
                    tables.append(md_table)
                i = j
                continue
        
        i += 1
    
    return tables


def convert_to_markdown_table(lines: List[str]) -> str:
    """Convert tab or pipe-separated lines to markdown table"""
    if not lines:
        return ""
    
    # Parse rows
    rows = []
    for line in lines:
        if '|' in line:
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
        elif '\t' in line:
            cells = [cell.strip() for cell in line.split('\t') if cell.strip()]
        else:
            continue
        
        if cells:
            rows.append(cells)
    
    if len(rows) < 2:
        return ""
    
    # Ensure all rows have same column count
    max_cols = max(len(row) for row in rows)
    normalized_rows = []
    for row in rows:
        while len(row) < max_cols:
            row.append("")
        normalized_rows.append(row)
    
    # Build markdown table
    md_lines = []
    md_lines.append("| " + " | ".join(normalized_rows[0]) + " |")
    md_lines.append("|" + "|".join(["---"] * max_cols) + "|")
    for row in normalized_rows[1:]:
        md_lines.append("| " + " | ".join(row) + " |")
    
    return "\n".join(md_lines)


def extract_action_items(text: str, date: datetime) -> List[Dict]:
    """
    Extract action items from email text looking for patterns like:
    - Task assignments with owners and deadlines
    - Table rows with Owner/Deadline columns
    """
    action_items = []
    
    # Pattern 1: Look for "Owner: Name" and "Deadline: Date" patterns
    owner_pattern = r'(?:owner|assigned to|responsible)[:\s]+([A-Za-z\s]+?)(?:\s*[,\n]|deadline|\|)'
    deadline_pattern = r'(?:deadline|due|by)[:\s]+([0-9]{1,2}\s+[A-Za-z]+\s+[0-9]{4})'
    
    # Pattern 2: Extract from markdown tables
    table_pattern = r'\|([^|]+)\|([^|]+)\|([^|]+)\|'
    
    lines = text.split('\n')
    current_item = {}
    
    for i, line in enumerate(lines):
        line_lower = line.lower()
        
        # Check for table rows (skip header and separator)
        if '|' in line and not line.strip().startswith('|---'):
            matches = re.findall(table_pattern, line)
            if matches:
                for match in matches:
                    cells = [cell.strip() for cell in match]
                    if len(cells) >= 3:
                        # Assume format: Task | Owner | Deadline
                        action_item = {
                            'date': date.strftime('%d %b %Y'),
                            'owner': cells[1],
                            'action': cells[0],
                            'deadline': cells[2]
                        }
                        # Validate it looks like actual data, not header
                        if action_item['owner'] and action_item['owner'].lower() not in ['owner', 'responsible', 'assigned']:
                            action_items.append(action_item)
        
        # Check for inline owner mentions
        owner_match = re.search(owner_pattern, line_lower, re.IGNORECASE)
        if owner_match:
            current_item['owner'] = owner_match.group(1).strip().title()
            current_item['action'] = line.strip()
        
        # Check for deadline mentions
        deadline_match = re.search(deadline_pattern, line, re.IGNORECASE)
        if deadline_match and current_item.get('owner'):
            current_item['deadline'] = deadline_match.group(1)
            current_item['date'] = date.strftime('%d %b %Y')
            action_items.append(current_item.copy())
            current_item = {}
    
    return action_items


def extract_key_points(text: str) -> List[str]:
    """Extract key points from email text"""
    key_points = []
    lines = text.split('\n')
    
    for line in lines:
        line = line.strip()
        # Look for bullet points or important statements
        if line.startswith('-') or line.startswith('•') or line.startswith('*'):
            key_points.append(line.lstrip('-•* ').strip())
        # Look for short declarative sentences (potential key points)
        elif line and len(line) < 150 and not line.startswith('|'):
            # Check if it's a meaningful statement
            if any(keyword in line.lower() for keyword in 
                   ['complete', 'assigned', 'track', 'due', 'deadline', 
                    'decision', 'approved', 'confirmed', 'reviewed']):
                key_points.append(line)
    
    return key_points[:5]  # Limit to top 5 key points


def parse_msg_file(msg_path: str) -> Optional[EmailData]:
    """Parse .msg file and extract email data"""
    try:
        msg = extract_msg.Message(msg_path)
        
        # Extract basic metadata
        sender = msg.sender or "Unknown Sender"
        subject = msg.subject or "No Subject"
        date_obj = msg.date
        
        # Parse date - msg.date can be string or datetime object
        date: datetime
        try:
            if date_obj is None:
                date = datetime.now()
            elif isinstance(date_obj, datetime):
                date = date_obj
            elif isinstance(date_obj, str):
                # Try parsing as string
                try:
                    date = datetime.strptime(date_obj, '%a, %d %b %Y %H:%M:%S %z')
                except ValueError:
                    # Try without timezone
                    date_str_clean = date_obj.split('+')[0].strip() if '+' in date_obj else date_obj
                    date = datetime.strptime(date_str_clean, '%a, %d %b %Y %H:%M:%S')
            else:
                date = datetime.now()
        except Exception:
            date = datetime.now()
        
        # Extract body
        body = msg.body or ""
        
        # Clean up body text
        body = body.replace('\r\n', '\n')
        
        filename = os.path.basename(msg_path)
        
        msg.close()
        
        return EmailData(sender, date, subject, body, filename)
    
    except Exception as e:
        print(f"Error parsing {msg_path}: {str(e)}")
        return None


def create_individual_markdown(email_data: EmailData, output_dir: str):
    """Create individual markdown file for an email"""
    
    # Create safe filename from subject
    safe_subject = re.sub(r'[^\w\s-]', '', email_data.subject)
    safe_subject = re.sub(r'[-\s]+', '_', safe_subject)
    md_filename = f"{email_data.date.strftime('%Y%m%d_%H%M')}_{safe_subject[:50]}.md"
    md_path = os.path.join(output_dir, md_filename)
    
    # Extract tables and action items
    tables = extract_tables_from_text(email_data.body)
    email_data.action_items = extract_action_items(email_data.body, email_data.date)
    email_data.key_points = extract_key_points(email_data.body)
    
    # Format date
    formatted_date = email_data.date.strftime('%a, %d %b %Y %H:%M')
    
    # Build markdown content
    content = f"# Email: {email_data.subject}\n\n"
    content += f"**From:** {email_data.sender}\n"
    content += f"**Date:** {formatted_date}\n"
    content += f"**Subject:** {email_data.subject}\n\n"
    content += "## Content\n"
    
    # Add body content
    body_lines = email_data.body.split('\n')
    for line in body_lines:
        if line.strip() and not line.strip().startswith('>'):
            content += line + '\n'
    
    # Add tables if found
    if tables:
        content += "\n"
        for table in tables:
            content += table + "\n\n"
    
    # Add key points section
    if email_data.key_points:
        content += "## Key Points\n"
        for point in email_data.key_points:
            content += f"- {point}\n"
    
    # Write to file
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"✓ Created: {md_filename}")
    return email_data


def create_timeline_summary(emails: List[EmailData], output_path: str):
    """Create chronological timeline summary"""
    
    # Sort emails by date
    emails.sort(key=lambda x: x.date)
    
    content = "# Email Communications Timeline\n\n"
    content += "## Chronological Email Summary\n\n"
    
    # Add chronological summaries
    for email in emails:
        formatted_date = email.date.strftime('%a, %d %b %Y %H:%M')
        content += f"### {formatted_date} - {email.sender}\n"
        content += f"**Subject:** {email.subject}\n"
        
        if email.key_points:
            content += "**Key Points:**\n"
            for point in email.key_points:
                content += f"- {point}\n"
        
        content += "\n"
    
    # Collect all action items
    all_action_items = []
    for email in emails:
        all_action_items.extend(email.action_items)
    
    # Add action items table
    if all_action_items:
        content += "## Decisions & Action Items\n\n"
        content += "| Date | Owner | Action Item | Deadline |\n"
        content += "|------|-------|-------------|----------|\n"
        
        for item in all_action_items:
            content += f"| {item['date']} | {item['owner']} | {item['action']} | {item['deadline']} |\n"
    
    # Write to file
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"✓ Timeline summary created: {output_path}")


def main():
    """Main execution function"""
    print("=" * 60)
    print("MSG to Markdown Converter with Timeline Generator")
    print("=" * 60)
    
    # Setup directories
    setup_directories()
    
    # Find all .msg files
    msg_files = glob.glob('project/emails/*.msg')
    
    if not msg_files:
        print("\n⚠ No .msg files found in project/emails/")
        print("Please place .msg files in the project/emails/ directory")
        return
    
    print(f"\nFound {len(msg_files)} .msg file(s)")
    print("-" * 60)
    
    # Process each .msg file
    emails = []
    for msg_file in msg_files:
        print(f"\nProcessing: {os.path.basename(msg_file)}")
        email_data = parse_msg_file(msg_file)
        
        if email_data:
            email_data = create_individual_markdown(email_data, 'project/markdown')
            emails.append(email_data)
    
    # Create timeline summary
    if emails:
        print("\n" + "-" * 60)
        print("Creating timeline summary...")
        create_timeline_summary(emails, 'project/timeline_summary.md')
    
    print("\n" + "=" * 60)
    print(f"✓ Conversion complete!")
    print(f"  - Processed {len(emails)} email(s)")
    print(f"  - Individual markdown files: project/markdown/")
    print(f"  - Timeline summary: project/timeline_summary.md")
    print("=" * 60)


if __name__ == "__main__":
    main()