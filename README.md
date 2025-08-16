# Word to LaTeX Converter

A simple Python utility to convert **Microsoft Word (.docx)** documents into **LaTeX (.tex)** format.  
The script preserves **headings, lists (ordered & unordered), bold, italic, underline**, and handles **special LaTeX characters** safely.

> ⚠️ This project is currently discontinued, but the codebase is left here for reference and future use.

---

## Features
- Converts `.docx` files to `.tex`
- Preserves text styles:
  - **Bold**
  - *Italic*
  - Underlined
- Supports nested lists (`itemize` & `enumerate`)
- Adds hyperlink support
- Escapes LaTeX special characters

---

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/word-to-latex.git
   cd word-to-latex
   ```

2. Install dependencies:
   ```bash
   pip install python-docx
   ```

---

## Usage

Run the script with:
```bash
python main.py input.docx output.tex
```

Example:
```bash
python main.py sample.docx output.tex
```

---

## Example Output

**Word input:**
- Heading 1
- Bullet list item
- Numbered list item

**LaTeX output:**
```latex
\section{Heading 1}
\begin{itemize}
  \item Bullet list item
\end{itemize}
\begin{enumerate}
  \item Numbered list item
\end{enumerate}
```

---

## Project Status

❌ **Cancelled** – development has been paused.  
The existing code can still be used as a base for anyone wanting to extend it.

---

## License
MIT License © 2025
