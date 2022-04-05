

# --- DEPRECATED ---


def translate(input_code_filename, output_code_filename):
    input_code_lines = open(input_code_filename, 'r').readlines()
    output_code_lines = []

    saw_import_line = False

    for line in enumerate(input_code_lines):
        if is_import_line(line):
            saw_import_from_line = True

        if saw_import_from_line and is_function_definition_line(line):


    output_code_lines += input_code_lines[line_idx + 1:]

    with open(output_code_filename, 'w') as f:
        f.writelines(output_code_lines)


def is_import_line(line: str) -> bool:
    splitted_line = [s.strip() for s in line.split()]

    if splitted_line[0].strip() == 'import':
        return True
    elif (len(splitted_line) >= 3) and ([splitted_line[0], splitted_line[2]] == ['from', 'import']):
        return True
    else:
        return False


def is_function_definition_line(line: str) -> bool:
    if line.split()[0].strip() == 'def':
        return True
    else:
        return False


if __name__ == "__main__":

    line = 'import pandas_cursor'

    print(is_import_line(line))