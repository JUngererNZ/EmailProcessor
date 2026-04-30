# update 17.04.2026 09:13
#!/usr/bin/env python3
import os
import datetime
from gpt4all import GPT4All

def generate_summary(model, content):
    """Generates a brief summary of the markdown content using the local LLM."""
    prompt = f"Summarize the following markdown content concisely in two sentences:\n\n{content[:2000]}"
    with model.chat_session():
        response = model.generate(prompt, max_tokens=500)
    return response.strip()

def main():
    cwd = os.getcwd()
    markdown_dir = os.path.join(cwd, 'markdown')

    if not os.path.exists(markdown_dir):
        print(f"Markdown directory not found: {markdown_dir}")
        return

    model_path = "Llama-3.2-1B-Instruct-Q4_0.gguf"
    print(f"Loading model: {model_path}...")
    model = GPT4All(model_name=model_path, allow_download=False)

    md_files = [f for f in os.listdir(markdown_dir) if f.endswith('.md')]
    if not md_files:
        print("No markdown files found.")
        return

    all_summaries = []
    latest_content = ""

    md_files.sort(key=lambda f: os.path.getmtime(os.path.join(markdown_dir, f)), reverse=True)

    for i, md_file in enumerate(md_files):
        path = os.path.join(markdown_dir, md_file)

        try:
            with open(path, 'r', encoding='utf-8-sig') as f:
                content = f.read()
        except UnicodeDecodeError:
            print(f"Standard UTF-8 failed for {md_file}, trying UTF-16...")
            with open(path, 'r', encoding='utf-16') as f:
                content = f.read()

        if i == 0:
            latest_content = content

        print(f"Summarizing {md_file}...")
        summary = generate_summary(model, content)
        all_summaries.append(f"### Summary of {md_file}\n{summary}\n")

    folder_name = os.path.basename(cwd)
    date_str = datetime.datetime.now().strftime('%Y%m%d')
    output_name = f'{folder_name}_{date_str}.md'
    output_path = os.path.join(cwd, output_name)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(latest_content)
        f.write("\n\n---\n## File Summaries\n\n")
        f.write("\n".join(all_summaries))

    print(f"Output written to {output_path}")

if __name__ == '__main__':
    main()