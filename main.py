from docx import Document
import re

def escape_latex(text):
    """Escape special LaTeX characters in text while preserving Unicode"""
    replacements = {
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
        "~": r"\textasciitilde{}",
        "^": r"\textasciicircum{}",
        "\\": r"\textbackslash{}",
    }
    for key, val in replacements.items():
        text = text.replace(key, val)
    return text

def detect_list_type(text):
    """Determine if a list is ordered (enumerate) or unordered (itemize)"""
    if re.match(r'^\s*(\d+[\.\)]|[a-zA-Z][\.\)])', text):
        return "enumerate"
    return "itemize"

def doc_to_latex(docx_file, tex_file):
    """Convert Word document to LaTeX format with nested list support"""
    try:
        doc = Document(docx_file)
    except FileNotFoundError:
        print(f"Error: File '{docx_file}' not found!")
        return

    latex = [
        r"\documentclass{article}",
        r"\usepackage[utf8]{inputenc}",
        r"\usepackage{graphicx}",
        r"\usepackage{hyperref}",
        r"\usepackage{enumitem} % For custom list formatting",
        r"\begin{document}"
    ]

    list_stack = [] 

    for para in doc.paragraphs:
        style = para.style.name
        text = "".join(
            f"\\textbf{{{escape_latex(run.text)}}}" if run.bold else
            f"\\textit{{{escape_latex(run.text)}}}" if run.italic else
            f"\\underline{{{escape_latex(run.text)}}}" if run.underline else
            escape_latex(run.text)
            for run in para.runs
        ).strip()

        if not text:
            continue

        if style.startswith("Heading"):
            while list_stack:
                latex.append(f"\\end{{{list_stack.pop()[0]}}}")
            
            level = int(style[-1]) if style[-1].isdigit() else 1
            latex.append(f"\\{'sub' * (level - 1)}section{{{text}}}")

        elif style.startswith(("List Paragraph", "List Bullet", "List Number")):
            list_type = detect_list_type(para.text)
            indent = para.paragraph_format.left_indent
            current_level = int(indent.pt // 18) if indent and indent.pt else 0

            while list_stack and list_stack[-1][1] > current_level:
                latex.append(f"\\end{{{list_stack.pop()[0]}}}")

            if not list_stack or list_stack[-1][1] < current_level or list_stack[-1][0] != list_type:
                latex.append(f"\\begin{{{list_type}}}")
                list_stack.append((list_type, current_level))

            item_text = re.sub(r'^\s*([•◦▪]|\d+[\.\)]|[a-zA-Z][\.\)])\s*', '', text)
            latex.append(f"  \\item {item_text}")

        else:
            while list_stack:
                latex.append(f"\\end{{{list_stack.pop()[0]}}}")
            latex.append(text + "\n")

    while list_stack:
        latex.append(f"\\end{{{list_stack.pop()[0]}}}")

    latex.append(r"\end{document}")

    try:
        with open(tex_file, "w", encoding="utf-8") as f:
            f.write('\n'.join(latex))
        print(f"Successfully created LaTeX file: {tex_file}")
    except IOError as e:
        print(f"Error writing to file: {e}")

if __name__ == "__main__":
    doc_to_latex("sample.docx", "output.tex")
