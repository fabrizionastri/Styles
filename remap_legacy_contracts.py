from docx import Document

def remap_styles(file_path, output_path):
    doc = Document(file_path)
    
    # Define your mapping dictionary
    mapping = {
        'Heading 1': 'Article 1',
        'Heading 2': 'Article 2',
        'Heading 3': 'Article 3',
        'Heading 4': 'Article 4',
        'Heading 6': 'Heading 4'
    }

    for paragraph in doc.paragraphs:
        if paragraph.style.name in mapping:
            # Apply the new style name
            paragraph.style = mapping[paragraph.style.name]

    doc.save(output_path)
    print(f"Success! Saved to {output_path}")

remap_styles('input.docx', 'output_path.docx')
