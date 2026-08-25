import argparse
import sys
from pathlib import Path

from docx import Document

from .convert import to_docx
from .sanitize import sanitize_document

SUPPORTED = {".pdf", ".doc", ".docx"}


def collect_inputs(paths):
    files = []
    for p in paths:
        p = Path(p)
        if p.is_dir():
            files.extend(sorted(f for f in p.iterdir() if f.suffix.lower() in SUPPORTED))
        elif p.is_file() and p.suffix.lower() in SUPPORTED:
            files.append(p)
        else:
            print(f"Skipping unsupported path: {p}", file=sys.stderr)
    return files


def process_file(path: Path, out_dir: Path) -> dict:
    docx_path, is_temp = to_docx(path)
    try:
        doc = Document(str(docx_path))
        stats = sanitize_document(doc)
        tag = "" if path.suffix.lower() == ".docx" else "_" + path.suffix.lower().lstrip(".")
        out_path = out_dir / f"{path.stem}{tag}_sanitized.docx"
        doc.save(str(out_path))
    finally:
        if is_temp:
            import shutil

            shutil.rmtree(docx_path.parent, ignore_errors=True)
    return {"input": path, "output": out_path, **stats}


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="sanitizer",
        description="Sanitize PDF/DOC documents: remove watermarks, replace branded words, enforce A4 portrait.",
    )
    parser.add_argument("inputs", nargs="+", help="PDF/DOC/DOCX files or folders")
    parser.add_argument("-o", "--output", default="sanitized", help="Output directory (default: ./sanitized)")
    args = parser.parse_args(argv)

    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    files = collect_inputs(args.inputs)
    if not files:
        print("No supported input files found.", file=sys.stderr)
        sys.exit(1)

    failures = 0
    for f in files:
        try:
            r = process_file(f, out_dir)
            print(
                f"[OK] {f.name} -> {r['output'].name} "
                f"(watermarks removed: {r['watermarks_removed']}, words replaced: {r['words_replaced']}, "
                f"sections fixed: {r['sections_fixed']}, shapes scaled: {r['shapes_scaled']})"
            )
        except Exception as e:
            failures += 1
            print(f"[FAIL] {f.name}: {e}", file=sys.stderr)

    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
