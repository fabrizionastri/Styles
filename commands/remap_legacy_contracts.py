

import argparse
import os

from docx import Document


def ensure_docx_suffix(path: str) -> str:
    """Ensure the provided path ends with .docx without duplicating the suffix."""
    return path if path.lower().endswith('.docx') else f"{path}.docx"


def default_output_path(input_path: str) -> str:
    """Append _remapped before the .docx extension for the default output path."""
    base, ext = os.path.splitext(input_path)
    ext = ext or '.docx'
    return f"{base}_remapped{ext}"


def remap_styles(file_path, output_path):
    doc = Document(file_path)

    # Mapping legacy headings to new article styles
    mapping = {
        'Title': 'Heading 1',
        'Heading 1': 'Article 1',
        'Heading 2': 'Article 2',
        'Heading 3': 'Article 3',
        'Heading 4': 'Article 4',
        'Heading 6': 'Heading 4'
    }

    for paragraph in doc.paragraphs:
        if paragraph.style.name in mapping:
            paragraph.style = mapping[paragraph.style.name]

    doc.save(output_path)
    print(f"Success! Saved to {output_path}")


def main():
    parser = argparse.ArgumentParser(description='Remap legacy Word heading styles.')
    parser.add_argument('input_file', help='Path to the source .docx file (extension optional).')
    parser.add_argument('output_file', nargs='?', help='Optional output .docx path (extension optional).')
    args = parser.parse_args()

    input_path = ensure_docx_suffix(args.input_file)
    output_path = ensure_docx_suffix(args.output_file) if args.output_file else default_output_path(input_path)

    remap_styles(input_path, output_path)


if __name__ == '__main__':
    main()
