import re
from pathlib import Path

# Explicitly declared language suffixes in lowercase
VALID_LANGUAGES = ["py", "c", "cpp", "potigol"]
LANG_PATTERN = "|".join(VALID_LANGUAGES)

# Regular expression to capture file name components:
# Matches S_<problem>_<domain>_<subdomain>[_<source_language>]
# The optional source_language group is only captured if it strictly matches VALID_LANGUAGES.
PATTERN = re.compile(
    rf"^S_(?P<problem>\d+)_(?P<domain>[^_]+)_(?P<subdomain>.+?)(?:_(?P<source_language>{LANG_PATTERN}))?$"
)

# Language normalization mapping
LANGUAGES_MAP = {
    "py": "Python",
    "c": "C",
    "cpp": "C++",
    "potigol": "Potigol",
}


def process_directories(root_dir: str, output_file: str = "info.txt"):
    root = Path(root_dir)
    extracted_data = []

    # Recursively traverse all files in the directory and subdirectories
    for file_path in root.rglob("*"):
        if not file_path.is_file() or file_path.name == output_file:
            continue

        filename_stem = file_path.stem
        match = PATTERN.match(filename_stem)

        if match:
            data = match.groupdict()
            problem_num = int(data["problem"])  # Converted to int for proper numeric sorting
            domain = data["domain"]
            project = data["subdomain"]

            # Use mapped language or default to 'Java' if omitted or not in VALID_LANGUAGES
            raw_lang = data["source_language"]
            if raw_lang:
                source_lang = LANGUAGES_MAP.get(raw_lang, raw_lang.title())
            else:
                source_lang = "Java"

            extracted_data.append({
                "problem": problem_num,
                "domain": domain,
                "project": project,
                "source_lang": source_lang
            })

    # Sort entries by problem number in ascending order
    extracted_data.sort(key=lambda item: item["problem"])

    # Format entries for the output text file
    info_lines = []
    for item in extracted_data:
        info_entry = (
            f"  - Problem: {item['problem']}\n"
            f"  - User: {item['domain']}\n"
            f"  - Project: {item['project']}\n"
            f"  - Source Language: {item['source_lang']}\n"
            f"{'-' * 40}\n"
        )
        info_lines.append(info_entry)

    # Save formatted output to info.txt
    output_path = root / output_file
    with open(output_path, "w", encoding="utf-8") as f:
        if info_lines:
            f.writelines(info_lines)
            print(f"Success! {len(info_lines)} file(s) processed and sorted.")
            print(f"Results saved to: {output_path.resolve()}")
        else:
            f.write("No files matching the expected pattern were found.\n")
            print("No matching files were found.")


if __name__ == "__main__":
    # Target directory (use '.' for current working directory)
    TARGET_DIR = "."
    process_directories(TARGET_DIR)
