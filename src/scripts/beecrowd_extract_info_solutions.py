import re
from pathlib import Path

# Sufixos de linguagem explicitamente declarados nos nomes dos arquivos
VALID_LANGUAGES = ["py", "c", "cpp", "potigol"]
LANG_PATTERN = "|".join(VALID_LANGUAGES)

# A expressão regular captura a linguagem de origem no final APENAS se for um dos sufixos acima.
# Todo o restante do nome (mesmo contendo a palavra 'Java') permanecerá no grupo 'subdomain' (Project).
PATTERN = re.compile(
    rf"^S_(?P<problem>\d+)_(?P<domain>[^_]+)_(?P<subdomain>.+?)(?:_(?P<source_language>{LANG_PATTERN}))?$",
    re.IGNORECASE,
)

LANGUAGES_MAP = {
    "py": "Python",
    "c": "C",
    "cpp": "C++",
    "potigol": "Potigol",
}


def process_directories(root_dir: str, output_file: str = "info.txt"):
    root = Path(root_dir)
    info_lines = []

    # Percorre recursivamente todos os arquivos
    for file_path in root.rglob("*"):
        if not file_path.is_file() or file_path.name == output_file:
            continue

        filename_stem = file_path.stem
        match = PATTERN.match(filename_stem)

        if match:
            data = match.groupdict()
            problem_num = data["problem"]
            domain = data["domain"]
            project = data["subdomain"]

            # Se não houver sufixo de linguagem mapeado (py, c, cpp, potigol), define como Java
            raw_lang = data["source_language"]
            if raw_lang:
                source_lang = LANGUAGES_MAP.get(raw_lang.lower(), raw_lang.title())
            else:
                source_lang = "Java"

            # Formatação de saída para o info.txt
            info_entry = (
                f"  - Problem: {problem_num}\n"
                f"  - User: {domain}\n"
                f"  - Project: {project}\n"
                f"  - Source Language: {source_lang}\n"
                f"{'-' * 40}\n"
            )
            info_lines.append(info_entry)

    # Gravação do resultado no info.txt
    output_path = root / output_file
    with open(output_path, "w", encoding="utf-8") as f:
        if info_lines:
            f.writelines(info_lines)
            print(f"Success! {len(info_lines)} file(s) processed.")
            print(f"Results saved to: {output_path.resolve()}")
        else:
            f.write("No files matching the expected pattern were found.\n")
            print("No matching files were found.")


if __name__ == "__main__":
    TARGET_DIR = "."
    process_directories(TARGET_DIR)
