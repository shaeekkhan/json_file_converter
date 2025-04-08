from docx import Document


def process_docx(file_path):
    doc = Document(file_path)
    data = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            data.append({"paragraph": text})

    return data
