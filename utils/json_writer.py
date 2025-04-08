import json
import os
from utils.excel_utils import process_excel
from utils.csv_utils import process_csv
from utils.text_utils import process_text
from utils.docx_utils import process_docx


def file_to_json(input_file, output_file):
    _, file_extension = os.path.splitext(input_file)
    file_extension = file_extension.lower()

    if file_extension in ['.xlsx', '.xls']:
        data = process_excel(input_file)
    elif file_extension == '.csv':
        data = process_csv(input_file)
    elif file_extension in ['.txt', '.log', '.md']:
        data = process_text(input_file)
    elif file_extension == '.docx':
        data = process_docx(input_file)
    else:
        print(f"Error: Unsupported file format {file_extension}")
        return

    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

    print(f"Conversion complete. JSON file saved as {output_file}")
