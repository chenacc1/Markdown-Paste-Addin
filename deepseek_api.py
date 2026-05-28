#!/usr/bin/env python3
"""DeepSeek API integration — ask DeepSeek and convert the answer to Word.

Requires: pip install openai
Setup: set DEEPSEEK_API_KEY environment variable.

Usage:
  python deepseek_api.py "Explain quantum entanglement" output.docx
  python deepseek_api.py --conversation chat.json output.docx  (multi-turn)
  python deepseek_api.py --interactive  (interactive chat mode)
"""

import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path

try:
    from openai import OpenAI
    _HAS_OPENAI = True
except ImportError:
    _HAS_OPENAI = False

from md2docx_lib import parse_markdown, build_docx


API_KEY = os.environ.get("DEEPSEEK_API_KEY", "")
BASE_URL = "https://api.deepseek.com"
DEFAULT_MODEL = "deepseek-chat"

SYSTEM_PROMPT = """You are a helpful assistant. When your response includes structured information, use Markdown formatting:
- Use `#` / `##` / `###` for headings
- Use `|...|` tables for tabular data
- Use ```mermaid code blocks for flowcharts and diagrams
- Use LaTeX math ($...$ or $$...$$) for formulas
- Use `- [ ]` task lists for action items
- Use `>` blockquotes for notable quotes
Keep responses well-structured and use these formats whenever appropriate."""


def ask_deepseek(prompt: str, conversation_history: list = None,
                 model: str = DEFAULT_MODEL, api_key: str = "",
                 stream: bool = False) -> tuple[str, list]:
    """Send a prompt to DeepSeek API. Returns (response_text, updated_history)."""
    if not _HAS_OPENAI:
        raise RuntimeError("openai package required. Run: pip install openai")

    client = OpenAI(api_key=api_key or API_KEY, base_url=BASE_URL)

    messages = conversation_history or []
    if not messages or messages[0].get("role") != "system":
        messages.insert(0, {"role": "system", "content": SYSTEM_PROMPT})

    messages.append({"role": "user", "content": prompt})

    response = client.chat.completions.create(
        model=model,
        messages=messages,
        stream=stream,
        temperature=0.7,
        max_tokens=8192,
    )

    if stream:
        answer = ""
        for chunk in response:
            if chunk.choices[0].delta.content:
                answer += chunk.choices[0].delta.content
                print(chunk.choices[0].delta.content, end="", flush=True)
        print()
    else:
        answer = response.choices[0].message.content

    messages.append({"role": "assistant", "content": answer})
    return answer, messages


def save_conversation_to_docx(messages: list, output_path: str):
    """Convert a multi-turn conversation to a Word document (Q&A format)."""
    chunks = []

    for msg in messages:
        if msg["role"] == "system":
            continue
        elif msg["role"] == "user":
            chunks.append({
                "type": "text",
                "text": f"# ❓ 问题\n\n{msg['content']}"
            })
        elif msg["role"] == "assistant":
            # Parse as markdown to get tables, code, etc.
            parsed = parse_markdown(msg["content"])
            chunks.extend(parsed)
            chunks.append({"type": "hr"})

    build_docx(chunks, output_path,
               title=f"DeepSeek 对话记录 - {datetime.now().strftime('%Y-%m-%d %H:%M')}",
               auto_captions=True, add_toc=False)


def interactive_mode(model: str, api_key: str):
    """Interactive chat with DeepSeek, save to Word on exit."""
    print(f"DeepSeek Chat ({model})")
    print("Type your questions. Enter /save to export to Word, /exit to quit.\n")

    messages = [{"role": "system", "content": SYSTEM_PROMPT}]

    while True:
        try:
            user_input = input("You: ").strip()
        except (EOFError, KeyboardInterrupt):
            break

        if not user_input:
            continue

        if user_input.lower() == "/exit":
            break
        if user_input.lower() == "/save":
            path = input("Save to (.docx): ").strip()
            if path:
                save_conversation_to_docx(messages, path)
                print(f"Saved: {path}")
            continue

        print("\nDeepSeek: ", end="", flush=True)
        try:
            answer, messages = ask_deepseek(
                user_input, conversation_history=messages,
                model=model, api_key=api_key, stream=True
            )
        except Exception as e:
            print(f"\nError: {e}")

    # Auto-save on exit
    filename = f"deepseek-chat-{datetime.now().strftime('%Y%m%d-%H%M%S')}.docx"
    save_conversation_to_docx(messages, filename)
    print(f"\nSaved: {filename}")


def main():
    parser = argparse.ArgumentParser(
        description="DeepSeek API → Word document converter"
    )
    parser.add_argument("prompt", nargs="?", help="Question to ask DeepSeek")
    parser.add_argument("output", nargs="?", default="deepseek-output.docx",
                        help="Output .docx path")
    parser.add_argument("--conversation", "-c", help="Load conversation JSON file")
    parser.add_argument("--model", "-m", default=DEFAULT_MODEL,
                        help=f"Model name (default: {DEFAULT_MODEL})")
    parser.add_argument("--api-key", "-k", default="",
                        help="DeepSeek API key (or set DEEPSEEK_API_KEY env)")
    parser.add_argument("--interactive", "-i", action="store_true",
                        help="Interactive chat mode")
    parser.add_argument("--save-conversation", "-s",
                        help="Save conversation history to JSON file")

    args = parser.parse_args()

    key = args.api_key or API_KEY
    if not key and not args.interactive:
        print("Error: DEEPSEEK_API_KEY not set.")
        print("  Set environment variable: set DEEPSEEK_API_KEY=sk-xxx")
        print("  Or pass: python deepseek_api.py --api-key sk-xxx")
        sys.exit(1)

    if args.interactive:
        interactive_mode(args.model, key)
        return

    if not args.prompt:
        parser.print_help()
        sys.exit(1)

    # Load conversation history
    history = []
    if args.conversation and os.path.isfile(args.conversation):
        history = json.loads(Path(args.conversation).read_text(encoding="utf-8"))

    print(f"Asking DeepSeek ({args.model}): {args.prompt[:100]}...")
    answer, history = ask_deepseek(
        args.prompt, conversation_history=history if history else None,
        model=args.model, api_key=key
    )

    # Save conversation JSON if requested
    if args.save_conversation:
        Path(args.save_conversation).write_text(
            json.dumps(history, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )

    # Convert answer to Word
    chunks = parse_markdown(answer)
    title = args.prompt[:80] + ("..." if len(args.prompt) > 80 else "")
    build_docx(chunks, args.output, title=title, auto_captions=True)
    print(f"\nDone: {args.output}")


if __name__ == "__main__":
    main()
