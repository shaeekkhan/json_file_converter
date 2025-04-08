import sys
from utils.file_selector import select_file
from utils.json_writer import file_to_json

if __name__ == "__main__":
    input_file = select_file()
    if not input_file:
        print("No file selected. Exiting...")
        sys.exit(1)

    output_file = input_file + ".json"
    file_to_json(input_file, output_file)
