import re
import shutil
from pathlib import Path

# Regular expression to capture the problem number from the filename:
# Group 1: Problem number (digits)
PATTERN = re.compile(
    r"^S_(?P<problem>\d+)_(?P<domain>[^_]+)_(?P<subdomain>[^_]+)(?:_(?P<source_language>.+))?$"
)


def organize_java_files(root_dir: str):
    root = Path(root_dir)
    moved_count = 0

    # Retrieve all .java files across the directory tree
    java_files = list(root.rglob("*.java"))

    for file_path in java_files:
        # Avoid processing files already named Main.java inside a problem_* folder
        if file_path.name == "Main.java" and file_path.parent.name.startswith(
            "problem_"
        ):
            continue

        filename_stem = file_path.stem
        match = PATTERN.match(filename_stem)

        if match:
            problem_id = match.group("problem")

            # Create destination folder: problem_<id>
            target_dir = root / f"problem_{problem_id}"
            target_dir.mkdir(parents=True, exist_ok=True)

            target_file = target_dir / "Main.java"

            # Move and rename the file
            shutil.move(str(file_path), str(target_file))
            print(
                f"Moved: {file_path.relative_to(root)} -> {target_file.relative_to(root)}"
            )
            moved_count += 1

            # Remove empty parent directory if the original file was inside a subfolder
            parent = file_path.parent
            if (
                parent != root
                and not any(parent.iterdir())
                and not parent.name.startswith("problem_")
            ):
                parent.rmdir()

    print(f"\nDone! Successfully reorganized {moved_count} Java file(s).")


if __name__ == "__main__":
    # Replace with target path (use "." for current working directory)
    TARGET_DIR = "."
    organize_java_files(TARGET_DIR)
