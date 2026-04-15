#!/usr/bin/env python3
"""
Script to merge data from markdown files in the 'markdown' folder into the FML-CHECKLIST-TEMPLATE.md
Generates an output file named {parent_folder}_{YYYYMMDD}.md with merged data.
"""

import os
import datetime
import re

def parse_markdown_table(lines):
    """Parse a markdown table from lines, return list of rows, each row is list of cells."""
    rows = []
    for line in lines:
        if line.strip().startswith('|'):
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            if cells:
                rows.append(cells)
    return rows

def extract_data_from_md(content):
    """Extract header and checklist data from markdown content."""
    lines = content.split('\n')
    tables = []
    current_table = []
    in_table = False
    for line in lines:
        if line.strip().startswith('|'):
            current_table.append(line)
            in_table = True
        elif in_table:
            if current_table:
                tables.append(current_table)
                current_table = []
            in_table = False
    if current_table:
        tables.append(current_table)

    header_data = {}
    checklist_data = {}

    if len(tables) >= 2:
        header_table = parse_markdown_table(tables[0])
        checklist_table = parse_markdown_table(tables[1])

        # Header table: assume first column is key, second is value
        for row in header_table[1:]:  # skip header row
            if len(row) >= 2:
                key = row[0].replace('**', '').replace(':', '').strip()
                value = row[1].strip()
                if value:
                    header_data[key] = value

        # Checklist table: Task, Drafted, Completed, Comments
        for row in checklist_table[2:]:  # skip header and separator
            if len(row) >= 4:
                task = row[0].strip()
                drafted = row[1].strip()
                completed = row[2].strip()
                comments = row[3].strip()
                checklist_data[task] = {
                    'drafted': drafted,
                    'completed': completed,
                    'comments': comments
                }

    return header_data, checklist_data

def main():
    cwd = os.getcwd()
    template_path = os.path.join(cwd, 'FML-CHECKLIST-TEMPLATE.md')
    if not os.path.exists(template_path):
        print(f"Template file not found: {template_path}")
        return

    markdown_dir = os.path.join(cwd, 'markdown')
    if not os.path.exists(markdown_dir):
        print(f"Markdown directory not found: {markdown_dir}")
        return

    md_files = [f for f in os.listdir(markdown_dir) if f.endswith('.md')]
    if not md_files:
        print("No markdown files found in markdown directory")
        return

    # Sort by modification time, latest first
    md_files.sort(key=lambda f: os.path.getmtime(os.path.join(markdown_dir, f)), reverse=True)

    merged_content = None
    for md_file in md_files:
        path = os.path.join(markdown_dir, md_file)
        try:
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read()
        except Exception as e:
            continue
        if 'TRANSPORT INSTRUCTION' in content:
            merged_content = content
            break

    if merged_content is None and md_files:
        # Fallback to latest
        latest_md = os.path.join(markdown_dir, md_files[0])
        with open(latest_md, 'r', encoding='utf-8') as f:
            merged_content = f.read()

    if merged_content is None:
        print("No suitable markdown content found")
        return

    # Generate output
    folder_name = os.path.basename(cwd)
    date_str = datetime.datetime.now().strftime('%Y%m%d')
    output_name = f'{folder_name}_{date_str}.md'
    output_path = os.path.join(cwd, output_name)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(merged_content)

    print(f"Output written to {output_path}")

if __name__ == '__main__':
    main()