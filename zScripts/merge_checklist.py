#!/usr/bin/env python3
"""
Script to extract executive summaries from all markdown files in the 'markdown' folder.
Generates an output file named {parent_folder}_executive_summaries_{YYYYMMDD}.md with all executive summaries in order from latest to first.
"""

import os
import datetime

def extract_executive_summary(content):
    """Extract the Executive Summary section from markdown content."""
    lines = content.split('\n')
    in_exec_summary = False
    exec_summary_lines = []
    for line in lines:
        if line.strip().startswith('## Executive Summary'):
            in_exec_summary = True
            exec_summary_lines.append(line)
        elif in_exec_summary and line.strip().startswith('## '):
            # Next section starts
            break
        elif in_exec_summary:
            exec_summary_lines.append(line)
    return '\n'.join(exec_summary_lines).strip()

def main():
    cwd = os.getcwd()
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

    executive_summaries = []
    for md_file in md_files:
        path = os.path.join(markdown_dir, md_file)
        try:
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read()
            exec_summary = extract_executive_summary(content)
            if exec_summary:
                executive_summaries.append(f"### From {md_file}\n\n{exec_summary}\n\n---\n")
        except Exception as e:
            print(f"Error reading {md_file}: {e}")
            continue

    if not executive_summaries:
        print("No executive summaries found in markdown files")
        return

    # Generate output
    folder_name = os.path.basename(cwd)
    date_str = datetime.datetime.now().strftime('%Y%m%d')
    output_name = f'{folder_name}_executive_summaries_{date_str}.md'
    output_path = os.path.join(cwd, output_name)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(f"# Executive Summaries - {folder_name}\n\n")
        f.write('\n'.join(executive_summaries))

    print(f"Output written to {output_path}")

if __name__ == '__main__':
    main()