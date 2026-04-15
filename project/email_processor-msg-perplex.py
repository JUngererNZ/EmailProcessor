import os
import re
import json
from datetime import datetime
from pathlib import Path
import extract_msg  # pip install extract-msg
from bs4 import BeautifulSoup  # pip install beautifulsoup4
import html2text  # pip install html2text


def extract_tables_to_md(html_content):
    """Extract HTML tables and convert to markdown tables."""
    soup = BeautifulSoup(html_content, 'html.parser')
    tables_md = []
    
    for table in soup.find_all('table'):
        rows = []
        all_rows = table.find_all('tr')
        if not all_rows:
            continue
        
        # Get headers from first row
        header_row = all_rows[0].find_all(['th', 'td'])
        if header_row:
            header = [th.get_text(strip=True) for th in header_row]
            rows.append('| ' + ' | '.join(header) + ' |')
            rows.append('| ' + ' | '.join(['---'] * len(header)) + ' |')
        
        # Get data rows
        for row in all_rows[1:]:
            cells = row.find_all(['td', 'th'])
            if cells:
                cell_data = [cell.get_text(strip=True) for cell in cells]
                rows.append('| ' + ' | '.join(cell_data) + ' |')
        
        if rows:
            tables_md.append('\n'.join(rows))
    
    return '\n\n'.join(tables_md)


def remove_disclaimer(text):
    """Remove FML Freight Solutions email disclaimer from text."""
    disclaimer_patterns = [
        # Original patterns
        r'E-?mail Disclaimer:.*?(?=^(?!(.*Disclaimer:|$)))',
        r'This message may contain information which is confidential or private in nature.*?FML Freight Solutions \(Pty\) Ltd reserves the right to monitor e-mails sent or received\.',
        r'Please note that FML Freight Solutions \(Pty\) Ltd reserves the right to monitor e-mails sent or received\.',
        r'FML Freight Solutions.*?scanned for viruses before being sent\.',
        
        # NEW: Additional disclaimer variation patterns
        r'This message may contain information which is confidential or private in nature.*?Neither FML Freight Solutions \(Pty\) Ltd nor the sender accepts any responsibility or liability for viruses',
        r'However, the recipient should also scan this e-?mail and any attached files for viruses and the like\.',
        r'Neither FML Freight Solutions \(Pty\) Ltd nor the sender accepts any responsibility or liability for viruses.*?resulting from the access of this e-?mail',
        r'If you have received this message in error, please notify the sender immediately by e-?mail, facsimile or telephone and thereafter return and/or delete it from your system\.',
    ]
    
    cleaned_text = text
    for pattern in disclaimer_patterns:
        cleaned_text = re.sub(pattern, '', cleaned_text, flags=re.IGNORECASE | re.DOTALL | re.MULTILINE)
    
    # Clean up extra newlines and whitespace
    cleaned_text = re.sub(r'\n\s*\n\s*\n', '\n\n', cleaned_text)
    cleaned_text = re.sub(r'\n{3,}', '\n\n', cleaned_text)
    return cleaned_text.strip()


def flatten_html_body(html_body):
    """Flatten nested replies and clean HTML for markdown conversion."""
    if not html_body:
        return ""
    
    soup = BeautifulSoup(html_body, 'html.parser')
    
    # Remove or flatten reply separators (common Outlook patterns)
    for div in soup.find_all('div', string=re.compile(r'(From:|Original Message:|Forwarded message:)', re.I)):
        div.decompose()
    
    # Convert to markdown first
    h = html2text.html2text(str(soup))
    
    # Remove disclaimer from markdown content
    h = remove_disclaimer(h)
    
    return h.strip()


def convert_msg_to_md(msg_file_path, output_dir):
    """Convert single .msg file to .md file."""
    try:
        with extract_msg.Message(msg_file_path) as msg:
            # Basic metadata
            sender = msg.sender or "Unknown Sender"
            subject = msg.subject or "No Subject"
            date_str = msg.date.strftime("%a, %d %b %Y %H:%M") if msg.date else "Unknown Date"
            
            # Extract and flatten body content
            html_body = msg.htmlBody if hasattr(msg, 'htmlBody') and msg.htmlBody else ""
            body_content = flatten_html_body(html_body)
            
            # Extract tables from HTML body (before disclaimer removal for better table detection)
            tables_md = extract_tables_to_md(html_body)
            
            # Generate key points (simple extraction)
            key_points = []
            if tables_md:
                key_points.append("- Contains task tables with assignments")
            
            # Simple keyword-based action item detection
            action_keywords = ['action', 'task', 'assign', 'due', 'deadline', 'complete', 'review']
            for keyword in action_keywords:
                if keyword.lower() in body_content.lower():
                    key_points.append(f"- {keyword.capitalize()} items mentioned")
                    break
            
            # Create markdown filename
            safe_subject = re.sub(r'[^\w\-_\.]', '_', subject)[:50]
            md_filename = f"{safe_subject}_{os.path.splitext(os.path.basename(msg_file_path))[0]}.md"
            md_path = output_dir / md_filename
            
            # Build markdown content
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
                md_content += '\n'.join([f"- {point}" for point in key_points])
            else:
                md_content += "- Review email content for details"
            
            # Save individual markdown file
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(md_content)
            
            print(f"✓ Converted: {msg_file_path} -> {md_path}")
            
            # Return metadata (no longer used for timeline)
            return {
                'date': msg.date,
                'date_str': date_str,
                'sender': sender,
                'subject': subject,
                'md_path': str(md_path),
                'key_points': key_points,
                'body': body_content
            }
    
    except Exception as e:
        print(f"✗ Error processing {msg_file_path}: {e}")
        import traceback
        traceback.print_exc()
        return None


def main():
    """Main processing function."""
    # Updated paths for your specific project location
    base_dir = Path(r"C:\Users\Jason\OneDrive - FML Freight Solutions\FML-PROJECTS\EmailProcessor\project")
    emails_dir = base_dir / "emails"
    markdown_dir = base_dir / "markdown"
    
    # Create output directories
    markdown_dir.mkdir(parents=True, exist_ok=True)
    
    if not emails_dir.exists():
        print(f"Error: {emails_dir} not found. Please create directory with .msg files.")
        return
    
    # Process all .msg files
    print("🔄 Processing .msg files...")
    emails_data = []
    
    for msg_file in emails_dir.glob("*.msg"):
        email_data = convert_msg_to_md(msg_file, markdown_dir)
        if email_data:
            emails_data.append(email_data)
    
    if emails_data:
        print(f"\n✅ Complete! Processed {len(emails_data)} emails.")
        print(f"📁 Individual MD files: {markdown_dir}")
        print("ℹ️  No timeline summary created per user request.")
    else:
        print("❌ No valid emails processed.")


if __name__ == "__main__":
    main()
