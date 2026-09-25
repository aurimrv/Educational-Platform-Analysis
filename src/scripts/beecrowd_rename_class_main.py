import os
import re

# Automatically detects the directory where this script is placed
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

def update_class_in_file(file_path, remove_package=True):
    content = None
    used_encoding = None

    for enc in ['utf-8', 'latin-1', 'cp1252', 'utf-8-sig']:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                content = f.read()
            used_encoding = enc
            break
        except Exception:
            continue

    if content is None:
        print(f"[ERROR] Could not read file: {file_path}")
        return

    # 1. Remove package declarations (e.g., package URI;)
    if remove_package:
        content = re.sub(r'^\s*package\s+[\w\.]+;\s*\n?', '', content, flags=re.MULTILINE)

    # 2. Match the class declaration
    class_pattern = re.compile(
        r'(\b(?:public\s+|protected\s+|private\s+)?(?:abstract\s+|final\s+|static\s+)*class\s+)([A-Za-z_][A-Za-z0-9_]*)'
    )

    match = class_pattern.search(content)
    if not match:
        print(f"[SKIPPED] No class definition found in: {file_path}")
        return

    old_class_name = match.group(2)

    # 3. Update class name to Main
    content = class_pattern.sub(r'\1Main', content, count=1)

    # 4. Update constructors matching the old class name
    if old_class_name != 'Main':
        constructor_pattern = re.compile(rf'\b{re.escape(old_class_name)}\b(?=\s*\()')
        content = constructor_pattern.sub('Main', content)

    # 5. Overwrite the file in place (file name remains unchanged)
    with open(file_path, 'w', encoding=used_encoding) as f:
        f.write(content)

    print(f"[UPDATED] {file_path}: '{old_class_name}' -> 'Main'")

def scan_current_folder():
    print(f"Processing all .java files in: {CURRENT_DIR}\n")
    for root, _, files in os.walk(CURRENT_DIR):
        for file in files:
            if file.endswith('.java'):
                full_path = os.path.join(root, file)
                update_class_in_file(full_path)

if __name__ == '__main__':
    scan_current_folder()