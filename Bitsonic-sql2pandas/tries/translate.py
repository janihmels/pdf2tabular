
def translate(input_code_filename, output_code_filename):

    output_code_lines = []

    for line in open(input_code_filename, 'r').readlines():
        translation = []

        if '.execute()' in line:
            translation.append()

        elif something:
            pass
        else:
            translation.append(line)

        output_code_lines += translation

    with open(output_code_filename, 'w') as f:
        f.writelines(output_code_lines)


def reformat_line(line: str) -> bool:
    if '.fetchall()' in line:
        return True

    if '.execute()' in line:
        return True

    # ELSE:

    return False