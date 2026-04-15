import os
import re
import sys
import json
from datetime import datetime
from pathlib import Path
import extract_msg  # pip install extract-msg
from bs4 import BeautifulSoup  # pip install beautifulsoup4
import html2text  # pip install html2text
from gpt4all import GPT4All  # pip install gpt4all

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

def flatten_html_body(html_body):
    """Flatten nested replies and clean HTML for markdown conversion."""
    if not html_body:
        return ""
    
    soup = BeautifulSoup(html_body, 'html.parser')
    
    # Remove or flatten reply separators (common Outlook patterns)
    for div in soup.find_all('div', string=re.compile(r'(From:|Original Message:|Forwarded message:)', re.I)):
        div.decompose()
    
    # Get main content, preferring body > div > p hierarchy
    content = soup.get_text(separator='\n\n', strip=True)
    
    # Convert to markdown while preserving structure
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
            # Basic metadata
            sender = msg.sender or "Unknown Sender"
            subject = msg.subject or "No Subject"
            date_str = msg.date.strftime("%a, %d %b %Y %H:%M") if msg.date else "Unknown Date"
            
            # Extract and flatten body content
            body_content = flatten_html_body(msg.htmlBody if hasattr(msg, 'htmlBody') and msg.htmlBody else "")
            
            # Extract tables from HTML body
            tables_md = extract_tables_to_md(msg.htmlBody if hasattr(msg, 'htmlBody') and msg.htmlBody else "")
            
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
            
            # Return metadata for timeline
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
        return None

def extract_actions_from_content(content, date_str):
    """Extract action items, owners, deadlines using simple heuristics."""
    actions = []
    
    # Simple regex patterns for common action formats
    patterns = [
        r'(\w+[\w\s]+?)\s+(?:assigned?|to|for)\s+([A-Z][\w\s]+?)(?:\s+(?:due|by|deadline)[:\s]*(\d{1,2}\s+[A-Za-z]{3}\s+\d{4}))?',
        r'task[:\s]*([^\n]+?)\s+owner[:\s]*([A-Z][\w\s]+?)(?:\s+deadline[:\s]*(\d{1,2}\s+[A-Za-z]{3}\s+\d{4}))?',
        r'([A-Z][\w\s]+?)\s+to\s+([^\n]+?)(?:\s+(?:by|due)\s+(\d{1,2}\s+[A-Za-z]{3}\s+\d{4}))?'
    ]
    
    for pattern in patterns:
        matches = re.finditer(pattern, content, re.IGNORECASE | re.MULTILINE)
        for match in matches:
            action = match.group(1).strip() if match.group(1) else "Action required"
            owner = match.group(2).strip() if match.group(2) else "TBD"
            deadline = match.group(3).strip() if match.group(3) else "TBD"
            actions.append({'date': date_str.split()[0:3], 'owner': owner, 'action': action, 'deadline': deadline})
    
    return actions

def generate_timeline_summary_with_llm(emails_data, output_path, folder_name):
    """Generate chronological timeline summary using local LLM."""
    # Sort emails by date (oldest first)
    emails_data = sorted([e for e in emails_data if e], key=lambda x: x['date'])
    
    # Prepare data for LLM
    email_summaries = []
    for email in emails_data:
        summary = f"Date: {email['date_str']}\nFrom: {email['sender']}\nSubject: {email['subject']}\nKey Points: {'; '.join(email['key_points'])}\nContent: {email['body'][:200]}..."  # Truncate for prompt
        email_summaries.append(summary)
    
    all_emails_text = "\n\n".join(email_summaries)
    
    # Limit total prompt to fit context window (approx 1500 tokens for safety)
    max_email_length = 1000  # characters, roughly 200-300 tokens
    if len(all_emails_text) > max_email_length:
        all_emails_text = all_emails_text[:max_email_length] + "..."
    
    # LLM Prompt for timeline
    prompt = f"""
You are an AI assistant helping to summarize email communications for freight logistics shipments.

Based on the following email data for shipment {folder_name}, create a chronological timeline summary in markdown format.

Email Data:
{all_emails_text}

Please generate a timeline summary that includes:
1. Chronological summary of key communications
2. Extracted action items, owners, and deadlines
3. Important decisions or changes
4. Current status of the shipment

Format as a clean markdown document with sections for timeline and action items.
"""
    
    try:
        # Initialize GPT4All model (downloads model if needed)
        model = GPT4All(model_name="Llama-3.2-1B-Instruct-Q4_0.gguf", allow_download=True)
        
        # Generate timeline
        timeline_content = model.generate(prompt, max_tokens=512)
        
        # Save timeline
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(f"# Email Communications Timeline - {folder_name}\n\n{timeline_content}")
        
        print(f"✓ AI-generated timeline created: {output_path}")
        return timeline_content
    except Exception as e:
        print(f"⚠️  Local LLM timeline generation failed: {e}")
        print("ℹ️  Falling back to simple timeline generation.")
        generate_timeline_summary(emails_data, output_path)
        return None

def extract_checklist_data_from_markdown(markdown_file_path):
    """Extract shipment details and checklist status from processed markdown file."""
    try:
        with open(markdown_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        shipment_data = {}
        checklist_items = {}
        
        # Extract shipment details from table format
        lines = content.split('\n')
        in_shipment_table = False
        
        for line in lines:
            line = line.strip()
            if 'File Ref:' in line and '|' in line:
                # Extract File Ref - handle both formats
                if '| File Ref:' in line:
                    parts = line.split('|')
                    if len(parts) >= 2:
                        shipment_data['file_ref'] = parts[1].replace('File Ref:', '').strip()
                elif 'File Ref:' in line and line.count('|') >= 2:
                    parts = line.split('|')
                    for part in parts:
                        if 'File Ref:' in part:
                            shipment_data['file_ref'] = part.replace('File Ref:', '').strip()
                            break
            elif 'Client Ref:' in line and '|' in line:
                if '| Client Ref:' in line:
                    parts = line.split('|')
                    if len(parts) >= 2:
                        shipment_data['client_ref'] = parts[1].replace('Client Ref:', '').strip()
                elif 'Client Ref:' in line and line.count('|') >= 2:
                    parts = line.split('|')
                    for part in parts:
                        if 'Client Ref:' in part:
                            shipment_data['client_ref'] = part.replace('Client Ref:', '').strip()
                            break
                if '| Client Ref:' in line:
                    parts = line.split('|')
                    if len(parts) >= 2:
                        shipment_data['client_ref'] = parts[1].replace('Client Ref:', '').strip()
                elif 'Client Ref:' in line and line.count('|') >= 2:
                    parts = line.split('|')
                    for part in parts:
                        if 'Client Ref:' in part:
                            shipment_data['client_ref'] = part.replace('Client Ref:', '').strip()
                            break
            elif 'Consignee:' in line and '|' in line:
                if '| Consignee:' in line:
                    parts = line.split('|')
                    if len(parts) >= 2:
                        shipment_data['consignee'] = parts[1].replace('Consignee:', '').strip()
                elif 'Consignee:' in line and line.count('|') >= 2:
                    parts = line.split('|')
                    for part in parts:
                        if 'Consignee:' in part:
                            shipment_data['consignee'] = part.replace('Consignee:', '').strip()
                            break
            elif 'Description:' in line and '|' in line:
                if '| Description:' in line:
                    parts = line.split('|')
                    if len(parts) >= 2:
                        shipment_data['description'] = parts[1].replace('Description:', '').strip()
                elif 'Description:' in line and line.count('|') >= 2:
                    parts = line.split('|')
                    for part in parts:
                        if 'Description:' in part:
                            shipment_data['description'] = part.replace('Description:', '').strip()
                            break
            elif 'PIN No:' in line and '|' in line:
                if '| PIN No:' in line:
                    parts = line.split('|')
                    if len(parts) >= 2:
                        shipment_data['pin_no'] = parts[1].replace('PIN No:', '').strip()
                elif 'PIN No:' in line and line.count('|') >= 2:
                    parts = line.split('|')
                    for part in parts:
                        if 'PIN No:' in part:
                            shipment_data['pin_no'] = part.replace('PIN No:', '').strip()
                            break
            elif 'Serial No:' in line and '|' in line:
                if '| Serial No:' in line:
                    parts = line.split('|')
                    if len(parts) >= 2:
                        shipment_data['serial_no'] = parts[1].replace('Serial No:', '').strip()
                elif 'Serial No:' in line and line.count('|') >= 2:
                    parts = line.split('|')
                    for part in parts:
                        if 'Serial No:' in part:
                            shipment_data['serial_no'] = part.replace('Serial No:', '').strip()
                            break
            elif 'Vessel:' in line and '|' in line:
                if '| Vessel:' in line:
                    parts = line.split('|')
                    if len(parts) >= 2:
                        shipment_data['vessel'] = parts[1].replace('Vessel:', '').strip()
                elif 'Vessel:' in line and line.count('|') >= 2:
                    parts = line.split('|')
                    for part in parts:
                        if 'Vessel:' in part:
                            shipment_data['vessel'] = part.replace('Vessel:', '').strip()
                            break
            elif 'Voy:' in line and '|' in line:
                if '| Voy:' in line:
                    parts = line.split('|')
                    if len(parts) >= 2:
                        shipment_data['voy'] = parts[1].replace('Voy:', '').strip()
                elif 'Voy:' in line and line.count('|') >= 2:
                    parts = line.split('|')
                    for part in parts:
                        if 'Voy:' in part:
                            shipment_data['voy'] = part.replace('Voy:', '').strip()
                            break
            elif 'Bill No:' in line and '|' in line:
                if '| Bill No:' in line:
                    parts = line.split('|')
                    if len(parts) >= 2:
                        shipment_data['bill_no'] = parts[1].replace('Bill No:', '').strip()
                elif 'Bill No:' in line and line.count('|') >= 2:
                    parts = line.split('|')
                    for part in parts:
                        if 'Bill No:' in part:
                            shipment_data['bill_no'] = part.replace('Bill No:', '').strip()
                            break
            elif 'ETA:' in line and '|' in line:
                if '| ETA:' in line:
                    parts = line.split('|')
                    if len(parts) >= 2:
                        eta_part = parts[1].replace('ETA:', '').strip()
                        shipment_data['eta'] = eta_part.split('|')[0].strip() if '|' in eta_part else eta_part
                elif 'ETA:' in line and line.count('|') >= 2:
                    parts = line.split('|')
                    for part in parts:
                        if 'ETA:' in part:
                            eta_part = part.replace('ETA:', '').strip()
                            shipment_data['eta'] = eta_part.split('|')[0].strip() if '|' in eta_part else eta_part
                            break
            elif '| Tariff Code, Weight & Cube:' in line:
                parts = line.split('|')
                if len(parts) >= 2:
                    shipment_data['tariff_code'] = parts[1].strip()
            elif '| Container #:' in line:
                parts = line.split('|')
                if len(parts) >= 2:
                    shipment_data['container_no'] = parts[1].strip()
            elif '| TRANSPORTER:' in line:
                parts = line.split('|')
                if len(parts) >= 2:
                    shipment_data['transporter'] = parts[1].strip()
            
            # Extract checklist items
            if '| TRANSPORT INSTRUCTION |' in line:
                parts = line.split('|')
                if len(parts) >= 5:
                    checklist_items['transport_instruction'] = {
                        'drafted': '✅' in parts[2] or 'X' in parts[2],
                        'completed': '✅' in parts[3] or 'X' in parts[3],
                        'comments': parts[4].strip() if len(parts) > 4 else ''
                    }
            elif '| X-HAUL INSTRUCTION |' in line:
                parts = line.split('|')
                if len(parts) >= 5:
                    checklist_items['x_haul_instruction'] = {
                        'drafted': '✅' in parts[2] or 'X' in parts[2],
                        'completed': '✅' in parts[3] or 'X' in parts[3],
                        'comments': parts[4].strip() if len(parts) > 4 else ''
                    }
            elif '| ANF / SWB |' in line:
                parts = line.split('|')
                if len(parts) >= 5:
                    checklist_items['anf_swb'] = {
                        'drafted': '✅' in parts[2] or 'X' in parts[2],
                        'completed': '✅' in parts[3] or 'X' in parts[3],
                        'comments': parts[4].strip() if len(parts) > 4 else ''
                    }
            elif '| LINE INVOICE |' in line:
                parts = line.split('|')
                if len(parts) >= 5:
                    checklist_items['line_invoice'] = {
                        'drafted': '✅' in parts[2] or 'X' in parts[2],
                        'completed': '✅' in parts[3] or 'X' in parts[3],
                        'comments': parts[4].strip() if len(parts) > 4 else ''
                    }
            elif '| CUSTOMS ENTRY (WE/RIT) |' in line:
                parts = line.split('|')
                if len(parts) >= 5:
                    checklist_items['customs_entry'] = {
                        'drafted': '✅' in parts[2] or 'X' in parts[2],
                        'completed': '✅' in parts[3] or 'X' in parts[3],
                        'comments': parts[4].strip() if len(parts) > 4 else ''
                    }
            # Add more checklist items as needed...
        
        return shipment_data, checklist_items
        
    except Exception as e:
        print(f"Error extracting checklist data: {e}")
        return {}, {}

def generate_populated_checklist(shipment_data, checklist_items, output_path):
    """Generate FML-CHECKLIST populated with extracted data."""
    
    # Read the template
    template_path = Path(__file__).parent.parent / "FML-CHECKLIST-TEMPLATE.md"
    if not template_path.exists():
        print(f"Warning: Template not found at {template_path}, using basic format")
        template_content = f"""# FML-CHECKLIST - {shipment_data.get('file_ref', 'Unknown')}

## Shipment Details
| Field | Value |
|-------|-------|
| File Ref | {shipment_data.get('file_ref', '')} |
| Client Ref | {shipment_data.get('client_ref', '')} |
| Consignee | {shipment_data.get('consignee', '')} |
| Description | {shipment_data.get('description', '')} |
| PIN No | {shipment_data.get('pin_no', '')} |
| Serial No | {shipment_data.get('serial_no', '')} |
| Vessel | {shipment_data.get('vessel', '')} |
| Voy | {shipment_data.get('voy', '')} |
| Bill No | {shipment_data.get('bill_no', '')} |
| ETA | {shipment_data.get('eta', '')} |
| Tariff Code, Weight & Cube | {shipment_data.get('tariff_code', '')} |
| Container # | {shipment_data.get('container_no', '')} |
| TRANSPORTER | {shipment_data.get('transporter', '')} |

## Checklist
| Task | Drafted | Completed | Comments |
|------|---------|-----------|----------|
| TRANSPORT INSTRUCTION | {'✅' if checklist_items.get('transport_instruction', {}).get('drafted') else '⬜'} | {'✅' if checklist_items.get('transport_instruction', {}).get('completed') else '⬜'} | {checklist_items.get('transport_instruction', {}).get('comments', '')} |
| X-HAUL INSTRUCTION | {'✅' if checklist_items.get('x_haul_instruction', {}).get('drafted') else '⬜'} | {'✅' if checklist_items.get('x_haul_instruction', {}).get('completed') else '⬜'} | {checklist_items.get('x_haul_instruction', {}).get('comments', '')} |
| ANF / SWB | {'✅' if checklist_items.get('anf_swb', {}).get('drafted') else '⬜'} | {'✅' if checklist_items.get('anf_swb', {}).get('completed') else '⬜'} | {checklist_items.get('anf_swb', {}).get('comments', '')} |
| LINE INVOICE | {'✅' if checklist_items.get('line_invoice', {}).get('drafted') else '⬜'} | {'✅' if checklist_items.get('line_invoice', {}).get('completed') else '⬜'} | {checklist_items.get('line_invoice', {}).get('comments', '')} |
| CUSTOMS ENTRY (WE/RIT) | {'✅' if checklist_items.get('customs_entry', {}).get('drafted') else '⬜'} | {'✅' if checklist_items.get('customs_entry', {}).get('completed') else '⬜'} | {checklist_items.get('customs_entry', {}).get('comments', '')} |
"""
    else:
        with open(template_path, 'r', encoding='utf-8') as f:
            template_content = f.read()
        
        # Replace empty fields with extracted data
        replacements = {
            '| File Ref:                                                                                   |': f'| File Ref: {shipment_data.get("file_ref", "")}',
            '| Client Ref:                                                                                 |': f'| Client Ref: {shipment_data.get("client_ref", "")}',
            '| Consignee:                                                                                  |': f'| Consignee: {shipment_data.get("consignee", "")}',
            '| Description:                                                                                |': f'| Description: {shipment_data.get("description", "")}',
            '| PIN No:                                                                                     |': f'| PIN No: {shipment_data.get("pin_no", "")}',
            '| Serial No:                                                                                  |': f'| Serial No: {shipment_data.get("serial_no", "")}',
            '| Vessel:                                                                                     |': f'| Vessel: {shipment_data.get("vessel", "")}',
            '| Voy:                                                                                        |': f'| Voy: {shipment_data.get("voy", "")}',
            '| Bill No:                                                                                    |': f'| Bill No: {shipment_data.get("bill_no", "")}',
            '| Container #:                                                                                |': f'| Container #: {shipment_data.get("container_no", "")}',
            '| TRANSPORTER:                                                                                |': f'| TRANSPORTER: {shipment_data.get("transporter", "")}',
        }
        
        for old, new in replacements.items():
            template_content = template_content.replace(old, new)
        
        # Replace checklist items
        checklist_replacements = {
            '| TRANSPORT INSTRUCTION                 |          |           |                           |': f'| TRANSPORT INSTRUCTION                 | {"✅" if checklist_items.get("transport_instruction", {}).get("drafted") else "⬜"} | {"✅" if checklist_items.get("transport_instruction", {}).get("completed") else "⬜"} | {checklist_items.get("transport_instruction", {}).get("comments", "")} |',
            '| X-HAUL INSTRUCTION                    |          |           |                           |': f'| X-HAUL INSTRUCTION                    | {"✅" if checklist_items.get("x_haul_instruction", {}).get("drafted") else "⬜"} | {"✅" if checklist_items.get("x_haul_instruction", {}).get("completed") else "⬜"} | {checklist_items.get("x_haul_instruction", {}).get("comments", "")} |',
            '| ANF / SWB                             |          |           |                           |': f'| ANF / SWB                             | {"✅" if checklist_items.get("anf_swb", {}).get("drafted") else "⬜"} | {"✅" if checklist_items.get("anf_swb", {}).get("completed") else "⬜"} | {checklist_items.get("anf_swb", {}).get("comments", "")} |',
            '| LINE INVOICE                          |          |           |                           |': f'| LINE INVOICE                          | {"✅" if checklist_items.get("line_invoice", {}).get("drafted") else "⬜"} | {"✅" if checklist_items.get("line_invoice", {}).get("completed") else "⬜"} | {checklist_items.get("line_invoice", {}).get("comments", "")} |',
            '| CUSTOMS ENTRY (WE/RIT)                |          |           |                           |': f'| CUSTOMS ENTRY (WE/RIT)                | {"✅" if checklist_items.get("customs_entry", {}).get("drafted") else "⬜"} | {"✅" if checklist_items.get("customs_entry", {}).get("completed") else "⬜"} | {checklist_items.get("customs_entry", {}).get("comments", "")} |',
        }
        
        for old, new in checklist_replacements.items():
            template_content = template_content.replace(old, new)
    
    # Write the populated checklist
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(template_content)
    
    print(f"✓ Populated checklist created: {output_path}")

def generate_checklist_with_llm(emails_data, template_path, output_path, folder_name):
    """Generate updated checklist using local LLM."""
    # Read template
    with open(template_path, 'r', encoding='utf-8') as f:
        template_content = f.read()
    
    # Truncate template if too long
    max_template_length = 500  # characters
    if len(template_content) > max_template_length:
        template_content = template_content[:max_template_length] + "..."
    
    # Prepare email data for LLM
    email_summaries = []
    for email in emails_data:
        summary = f"Date: {email['date_str']}\nFrom: {email['sender']}\nSubject: {email['subject']}\nContent: {email['body'][:200]}..."
        email_summaries.append(summary)
    
    all_emails_text = "\n\n".join(email_summaries)
    
    # Limit total prompt to fit context window
    max_email_length = 600  # characters for checklist
    if len(all_emails_text) > max_email_length:
        all_emails_text = all_emails_text[:max_email_length] + "..."
    
    # LLM Prompt for checklist
    prompt = f"""
You are an AI assistant helping to create a freight logistics checklist for shipment {folder_name}.

Based on the following email communications, fill in the checklist template with relevant information extracted from the emails.

Email Data:
{all_emails_text}

Checklist Template:
{template_content}

Please:
1. Extract shipment details (File Ref, Client Ref, Consignee, Description, etc.) from the emails
2. Update the checklist items with status information where mentioned
3. Fill in any relevant dates, numbers, or details
4. Keep the same table format but populate with actual data from emails
5. If information is not available, leave as blank or note "TBD"

Return the complete updated checklist in markdown format.
"""
    
    try:
        # Use the same model
        model = GPT4All(model_name="Llama-3.2-1B-Instruct-Q4_0.gguf", allow_download=True)
        
        # Generate checklist
        checklist_content = model.generate(prompt, max_tokens=1024)
        
        # Save checklist
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(checklist_content)
        
        print(f"✓ AI-generated checklist created: {output_path}")
    except Exception as e:
        print(f"⚠️  Local LLM checklist generation failed: {e}")
        print("ℹ️  Creating basic checklist template.")
        # Create a basic checklist
        basic_checklist = f"""# FML Checklist - {folder_name}

Based on email communications, here's a basic checklist for the shipment.

## Shipment Details
- File Reference: Extract from emails
- Client Reference: Extract from emails
- Description: Equipment shipment
- Status: In progress

## Checklist Items
- [ ] Transport Instruction
- [ ] X-Haul Instruction  
- [ ] ANF / SWB
- [ ] Line Invoice
- [ ] Customs Entry
- [ ] Cargo Dues
- [ ] Release Letter
- [ ] BV Inspection
- [ ] GRN / XE
- [ ] Final Documentation

*Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*
"""
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(basic_checklist)
        print(f"✓ Basic checklist created: {output_path}")

def generate_timeline_summary(emails_data, output_path):
    """Generate chronological timeline summary."""
    # Sort emails by date (oldest first)
    emails_data = sorted([e for e in emails_data if e], key=lambda x: x['date'])
    
    # Collect all actions across emails
    all_actions = []
    for email in emails_data:
        actions = extract_actions_from_content(email['body'], email['date_str'])
        all_actions.extend(actions)
    
    # Generate chronological summary section
    chrono_section = "## Chronological Email Summary\n\n"
    for email in emails_data:
        sender_link = f"[{email['sender']}](mailto:{email['sender'].split('@')[0] if '@' in email['sender'] else email['sender']})"
        chrono_section += f"### {email['date_str']} - {sender_link}\n"
        chrono_section += f"**Subject:** {email['subject']}\n"
        chrono_section += "**Key Points:**\n"
        
        if email['key_points']:
            for point in email['key_points']:
                chrono_section += f"- {point}\n"
        else:
            chrono_section += "- Review email for details\n"
        chrono_section += "\n"
    
    # Generate actions table
    actions_table = "## Decisions & Action Items\n\n"
    if all_actions:
        actions_table += "| Date | Owner | Action Item | Deadline |\n"
        actions_table += "|------|-------|-------------|----------|\n"
        for action in sorted(all_actions, key=lambda x: x['date']):
            date_short = ' '.join(action['date'][:3])
            actions_table += f"| {date_short} | {action['owner']} | {action['action']} | {action['deadline']} |\n"
    else:
        actions_table += "*No explicit action items detected*\n"
    
    # Write timeline summary
    timeline_content = f"""# Email Communications Timeline

{chrono_section}

{actions_table}
"""
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(timeline_content)
    
    print(f"✓ Timeline summary created: {output_path}")

def main():
    """Main processing function."""
    if len(sys.argv) > 1:
        base_dir = Path(sys.argv[1]).resolve()
    else:
        cwd = Path.cwd()
        if (cwd / "emails").exists():
            base_dir = cwd
        else:
            script_dir = Path(__file__).resolve().parent
            if (script_dir.parent / "emails").exists():
                base_dir = script_dir.parent
            else:
                base_dir = cwd

    emails_dir = base_dir / "emails"
    markdown_dir = base_dir / "markdown"
    template_path = base_dir / "FML-CHECKLIST-TEMPLATE.md"
    
    # Create output directories
    markdown_dir.mkdir(parents=True, exist_ok=True)
    folder_name = base_dir.name
    timeline_filename = f"timeline_{folder_name}_summary.md"
    timeline_path = markdown_dir / timeline_filename
    checklist_path = markdown_dir / f"FML-CHECKLIST-{folder_name}.md"
    
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
        # Generate timeline (simple version for now)
        generate_timeline_summary(emails_data, timeline_path)
        
        # Extract data from processed markdown files and generate populated checklist
        print("🔍 Extracting shipment data from processed emails...")
        shipment_data = {}
        checklist_items = {}
        
        # Find the most recent processed markdown file that contains shipment data
        markdown_files = [f for f in markdown_dir.glob("*.md") if not f.name.startswith("timeline_") and not f.name.startswith("FML-CHECKLIST")]
        if markdown_files:
            # Try to find a file that contains shipment data (File Ref, Client Ref, etc.)
            shipment_files = []
            for md_file in markdown_files:
                try:
                    with open(md_file, 'r', encoding='utf-8') as f:
                        content = f.read()
                        if any(keyword in content for keyword in ['File Ref:', 'Client Ref:', 'Consignee:', 'Description:']):
                            shipment_files.append(md_file)
                except:
                    continue
            
            # Use shipment file if found, otherwise use most recent
            if shipment_files:
                latest_md = max(shipment_files, key=lambda x: x.stat().st_mtime)
                print(f"📄 Using shipment data from: {latest_md.name}")
            else:
                latest_md = max(markdown_files, key=lambda x: x.stat().st_mtime)
                print(f"📄 Using data from: {latest_md.name}")
            
            shipment_data, checklist_items = extract_checklist_data_from_markdown(latest_md)
            
            if shipment_data or checklist_items:
                generate_populated_checklist(shipment_data, checklist_items, checklist_path)
                print(f"✓ Populated checklist created: {checklist_path}")
            else:
                print("⚠️  No shipment data extracted, creating basic checklist.")
                # Create a basic checklist
                basic_checklist = f"""# FML Checklist - {folder_name}

Based on email communications, here's a basic checklist for the shipment.

## Shipment Details
- File Reference: Extract from emails
- Client Reference: Extract from emails  
- Description: Equipment shipment
- Status: In progress

## Checklist Items
- [ ] Transport Instruction
- [ ] X-Haul Instruction  
- [ ] ANF / SWB
- [ ] Line Invoice
- [ ] Customs Entry
- [ ] Cargo Dues
- [ ] Release Letter
- [ ] BV Inspection
- [ ] GRN / XE
- [ ] Final Documentation

*Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*
"""
                with open(checklist_path, 'w', encoding='utf-8') as f:
                    f.write(basic_checklist)
                print(f"✓ Basic checklist created: {checklist_path}")
        else:
            print("⚠️  No markdown files found, skipping checklist generation.")
        
        print(f"\n✅ Complete! Processed {len(emails_data)} emails.")
        print(f"📁 Individual MD files: {markdown_dir}")
        print(f"📋 Timeline summary: {timeline_path}")
        print(f"📋 Checklist: {checklist_path}")
    else:
        print("❌ No valid emails processed.")

if __name__ == "__main__":
    main()
