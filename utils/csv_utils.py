import csv


def process_csv(file_path):
    with open(file_path, 'r', newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        data = []
        for row in reader:
            row_data = {k: v for k, v in row.items() if v}
            if row_data:
                data.append(row_data)
    return data
