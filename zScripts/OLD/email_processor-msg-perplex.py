import os
import re
import sys
import json
import urllib.request
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
            
        header_row = all_rows[0].find_all(['th', 'td'])
        if header_row:
            header = [th.get_text(strip=True) for th in header_row]
            rows.append('| ' + ' | '.join(header) + ' |')
            rows.append('| ' + ' | '.join(['---'] * len(header)) + ' |')
        
        for row in all_rows[1:]:
            cells = row.find_all(['td', 'th'])
            if cells:
                cell_data = [cell.get_text(strip=True) for cell in cells]
                rows.append('| ' + ' | '.join(cell_data) + ' |')
        
        if rows:
            tables_md.append('\n'.join(rows))
    
    return '\n\n'.join(tables_md)

def flatten_html_body(html_body):
    """Flatten nested replies and clean HTML for markdown conversion."""
    if not html_body:
        return ""
    
    soup = BeautifulSoup(html_body, 'html.parser')
    for div in soup.find_all('div', string=re.compile(r'(From:|Original Message:|Forwarded message:)', re.I)):
        div.decompose()
    
    h = html2text.HTML2Text()
    h.ignore_links = False
    h.ignore_images = True
    h.body_width = 0
    markdown = h.handle(str(soup))
    
    return markdown.strip()

def convert_msg_to_md(msg_file_path, output_dir):
    """Convert single .msg file to .md file."""
    try:
        with extract_msg.Message(msg_file_path) as msg:
            sender = msg.sender or "Unknown Sender"
            subject = msg.subject or "No Subject"
            date_str = msg.date.strftime("%a, %d %b %Y %H:%M") if msg.date else "Unknown Date"
            
            body_content = flatten_html_body(msg.htmlBody if hasattr(msg, 'htmlBody') and msg.htmlBody else "")
            tables_md = extract_tables_to_md(msg.htmlBody if hasattr(msg, 'htmlBody') and msg.htmlBody else "")
            
            safe_subject = re.sub(r'[^\w\-_\.]', '_', subject)[:50]
            md_filename = f"{safe_subject}_{os.path.splitext(os.path.basename(msg_file_path))[0]}.md"
            md_path = output_dir / md_filename
            
            md_content = f"# Email: {subject}\n**From:** {sender}\n**Date:** {date_str}\n\n{body_content}\n\n{tables_md}"
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(md_content)
            
            print(f"✓ Converted: {msg_file_path}")
            return {
                'date': msg.date,
                'date_str': date_str,
                'sender': sender,
                'subject': subject,
                'body': body_content
            }
    except Exception as e:
        print(f"✗ Error processing {msg_file_path}: {e}")
        return None

def generate_timeline_summary_with_llm(emails_data, output_path, folder_name):
    """Generate chronological timeline summary using local Ollama LLM."""
    emails_data = sorted([e for e in emails_data if e], key=lambda x: x['date'])
    email_summaries = [f"Date: {e['date_str']}\nFrom: {e['sender']}\nSubject: {e['subject']}\nContent: {e['body'][:300]}..." for e in emails_data]
    all_emails_text = "\n\n".join(email_summaries)
    
    prompt = f"Based on these emails for shipment {folder_name}, create a markdown timeline and executive summary:\n\n{all_emails_text}"
    
    payload = {
        "model": "llama3:latest",
        "prompt": prompt,
        "stream": False,
        "options": {"num_predict": 800, "temperature": 0.3}
    }
    
    try:
        req = urllib.request.Request("http://localhost:11434/api/generate", data=json.dumps(payload).encode('utf-8'), headers={'Content-Type': 'application/json'})
        with urllib.request.urlopen(req) as response:
            result = json.loads(response.read().decode('utf-8'))
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(result.get("response", ""))
        print(f"✓ Timeline created: {output_path}")
    except Exception as e:
        print(f"⚠️ Timeline LLM failed: {e}")

def generate_checklist_with_llm(emails_data, template_path, output_path, folder_name):
    """Populate FML-CHECKLIST-TEMPLATE using local Ollama LLM and inject timestamp."""
    if not template_path.exists():
        print(f"⚠️ Template not found at {template_path}")
        return

    with open(template_path, 'r', encoding='utf-8') as f:
        template_content = f.read()

    all_emails_text = "\n\n".join([f"From: {e['sender']}\nSubject: {e['subject']}\nContent: {e['body'][:500]}" for e in emails_data])
    
    prompt = f"""
You are a logistics assistant. Extract shipment info from these emails to fill out the provided template.

EMAILS:
{all_emails_text}

TEMPLATE TO FILL:
{template_content}

INSTRUCTIONS:
1. Populate File Ref, Client Ref, Consignee, ETA, Vessel, etc.
2. Mark '✅' in the Completed column ONLY if confirmed in emails.
3. Return ONLY the completed ASCII table. Do not change the table structure.
4. Ensure you keep the placeholder text {{dd.MM.yyyy HH:mm:ss}} at the bottom so the script can replace it.
"""
    
    payload = {
        "model": "llama3:latest",
        "prompt": prompt,
        "stream": False,
        "options": {"num_predict": 2500, "temperature": 0.1}
    }
    
    try:
        req = urllib.request.Request("http://localhost:11434/api/generate", data=json.dumps(payload).encode('utf-8'), headers={'Content-Type': 'application/json'})
        with urllib.request.urlopen(req) as response:
            result = json.loads(response.read().decode('utf-8'))
            checklist_content = result.get("response", "").strip()
            
            # Post-process the LLM output to inject the real timestamp
            current_time = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
            checklist_content = checklist_content.replace("{{dd.MM.yyyy HH:mm:ss}}", current_time)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(checklist_content)
        print(f"✓ AI-Populated checklist created with timestamp: {output_path}")
    except Exception as e:
        print(f"⚠️ Checklist LLM failed: {e}")

def main():
    base_dir = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else Path.cwd()
    emails_dir = base_dir / "emails"
    markdown_dir = base_dir / "markdown"
    template_path = base_dir / "FML-CHECKLIST-TEMPLATE.md"
    
    markdown_dir.mkdir(parents=True, exist_ok=True)
    folder_name = base_dir.name
    timeline_path = markdown_dir / f"timeline_{folder_name}_summary.md"
    checklist_path = markdown_dir / f"FML-CHECKLIST-{folder_name}.md"
    
    if not emails_dir.exists():
        print(f"Error: {emails_dir} not found.")
        return
    
    emails_data = []
    for msg_file in emails_dir.glob("*.msg"):
        data = convert_msg_to_md(msg_file, markdown_dir)
        if data: emails_data.append(data)
    
    if emails_data:
        generate_timeline_summary_with_llm(emails_data, timeline_path, folder_name)
        generate_checklist_with_llm(emails_data, template_path, checklist_path, folder_name)
        print(f"\n✅ Processing complete for {folder_name}.")

if __name__ == "__main__":
    main()