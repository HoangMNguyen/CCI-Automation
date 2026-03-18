import os


def search_function_in_files(directory, function_name):
    results = []

    # Walk through the directory structure
    for dirpath, dirnames, filenames in os.walk(directory):
        for filename in filenames:
            # Check if the file is a .py or .txt file
            if filename.endswith((".py", ".txt")):
                filepath = os.path.join(dirpath, filename)
                with open(filepath, "r", encoding="utf-8", errors="ignore") as file:
                    lines = file.readlines()
                    count = 0
                    matching_lines = []

                    # Check each line for the function name
                    for index, line in enumerate(lines):
                        if function_name in line:
                            count += 1
                            matching_lines.append(
                                (index + 1, line.strip())
                            )  # index+1 because line numbers start from 1

                    if count:
                        results.append({"file": filepath, "count": count, "lines": matching_lines})

    return results


if __name__ == "__main__":
    directory_to_search = os.getcwd()
    function_name_to_search = "np.NaN"
    results = search_function_in_files(directory_to_search, function_name_to_search)
    for result in results:
        print(f"\nFunction found in {result['file']} {result['count']} times.")
        for line_num, line in result["lines"]:
            print(f"    Line {line_num}: {line}")
