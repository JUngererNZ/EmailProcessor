import os
import re
import extract_msg
from bs4 import BeautifulSoup
from datetime import datetime

# Configuration
INPUT_DIR = 'project/emails/'
OUTPUT_DIR = 'project/markdown/'
SUMMARY_FILE = 'project/timeline_summary.md'

# Ensure directories exist
os.makedirs(OUTPUT_DIR, exist_ok=True)

def html_to_markdown_table(html_content):
    """Parses HTML to find tables and convert them to Markdown format."""
    if not html_content:
        return ""
    
    soup = BeautifulSoup(html_content, 'html.parser')
    tables = soup.find_all('table')
    md_tables = []

    for table in tables:
        rows = []
        for tr in table.find_all('tr'):
            cells = [td.get_text(strip=True) for td in tr.find_all(['td', 'th'])]
            if cells:
                rows.append(f"| {' | '.join(cells)} |")
        
        if rows:
            # Create the header separator
            num_columns = rows[0].count('|') - 1
            separator = f"| {' | '.join(['---'] * num_columns)} |"
            rows.insert(1, separator)
            md_tables.append("\n".join(rows))
    
    return "\n\n".join(md_tables)

def process_emails():
    email_data = []

    # 1. Process individual .msg files
    for filename in os.listdir(INPUT_DIR):
        if filename.endswith('.msg'):
            msg_path = os.path.join(INPUT_DIR, filename)
            msg = extract_msg.Message(msg_path)
            
            # Metadata
            sender = msg.sender
            subject = msg.subject
            date_raw = msg.date # datetime object
            date_str = date_raw.strftime("%a, %d %b %Y %H:%M") if date_raw else "Unknown Date"
            
            # Content Extraction
            body = msg.body
            html_body = msg.htmlBody
            markdown_tables = html_to_markdown_table(html_body)
            
            # Simple logic to extract "Key Points" (lines starting with bullets or containing 'due')
            key_points = [line.strip() for line in body.split('\n') if line.strip().startswith(('-', '*', '•'))]
            
            # Individual Markdown Content
            md_content = f"# Email: {subject}\n\n"
            md_content += f"**From:** {sender}\n"
            md_content += f"**Date:** {date_str}\n"
            md_content += f"**Subject:** {subject}\n\n"
            md_content += f"## Content\n{body}\n\n"
            if markdown_tables:
                md_content += f"{markdown_tables}\n\n"
            
            md_content += "## Key Points\n"
            md_content += "\n".join(key_points) if key_points else "- No specific key points identified."

            # Save individual file
            clean_name = re.sub(r'[^\w\-_\. ]', '_', filename.replace('.msg', ''))
            with open(os.path.join(OUTPUT_DIR, f"{clean_name}.md"), 'w', encoding='utf-8') as f:
                f.write(md_content)

            # Store data for timeline
            email_data.append({
                'datetime': date_raw if date_raw else datetime.min,
                'date_str': date_str,
                'sender': sender,
                'subject': subject,
                'key_points': key_points,
                'tables': markdown_tables
            })
            msg.close()

    # 2. Generate Timeline Summary
    email_data.sort(key=lambda x: x['datetime']) # Chronological sort

    with open(SUMMARY_FILE, 'w', encoding='utf-8') as f:
        f.write("# Email Communications Timeline\n\n")
        f.write("## Chronological Email Summary\n\n")
        
        for email in email_data:
            f.write(### {email['date_str']} - {email['sender']}\n")
            f.write(f"**Subject:** {email['subject']}\n")
            f.write("**Key Points:**\n")
            points = "\n".join(email['key_points']) if email['key_points'] else "- (Summary not available)"
            f.write(f"{points}\n\n")

        f.write("## Decisions & Action Items\n\n")
        # Aggregating all tables found in the emails
        all_tables = [e['tables'] for e in email_data if e['tables']]
        if all_tables:
            f.write("\n\n".join(all_tables))
        else:
            f.write("No action items or tables found.")

    print(f"Success: Processed {len(email_data)} emails. Summary available at {SUMMARY_FILE}")

if __name__ == "__main__":
    process_emails()