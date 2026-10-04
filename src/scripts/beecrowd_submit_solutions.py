import os
import re
import sys
import time
from pathlib import Path
from dotenv import load_dotenv
import openpyxl
from playwright.sync_api import sync_playwright

# Load configuration from .env file
load_dotenv()

BEECROWD_EMAIL = os.getenv("BEECROWD_EMAIL")
BEECROWD_PASSWORD = os.getenv("BEECROWD_PASSWORD")
BEECROWD_URL = os.getenv("BEECROWD_URL", "https://judge.beecrowd.com/pt/login")
SOLUTIONS_DIR = os.getenv("SOLUTIONS_DIR", "./")
EXCEL_FILE = os.getenv("EXCEL_FILE", "beecrowd_problems_shared.xlsx")


def validate_environment() -> None:
    """Ensure essential credentials and files exist."""
    if not BEECROWD_EMAIL or not BEECROWD_PASSWORD:
        print("Error: BEECROWD_EMAIL and BEECROWD_PASSWORD must be defined in your .env file.")
        sys.exit(1)

    if not os.path.exists(EXCEL_FILE):
        print(f"Error: Excel file '{EXCEL_FILE}' was not found in the current directory.")
        sys.exit(1)


def load_excel_problem_map(excel_path: str):
    """Load problem metadata and map problem IDs to worksheet rows."""
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active

    headers = [cell.value for cell in sheet[1]]

    id_col = headers.index("ID") + 1 if "ID" in headers else 1
    category_col = headers.index("Category") + 1 if "Category" in headers else None
    lang_col = headers.index("Language") + 1 if "Language" in headers else None
    approved_col = headers.index("BeeCrowd Approved") + 1 if "BeeCrowd Approved" in headers else None
    just_col = headers.index("Justification") + 1 if "Justification" in headers else None

    problem_map = {}
    for row in range(2, sheet.max_row + 1):
        prob_id = sheet.cell(row=row, column=id_col).value
        if prob_id is not None:
            try:
                prob_id_int = int(prob_id)
                category = sheet.cell(row=row, column=category_col).value if category_col else None
                approved = sheet.cell(row=row, column=approved_col).value if approved_col else None
                problem_map[prob_id_int] = {
                    "row": row,
                    "category": str(category) if category else "",
                    "approved": str(approved) if approved else "",
                }
            except ValueError:
                continue

    wb.close()
    return problem_map, {
        "id_col": id_col,
        "category_col": category_col,
        "lang_col": lang_col,
        "approved_col": approved_col,
        "just_col": just_col,
    }


def update_excel_row(
    excel_path: str,
    cols_config: dict,
    problem_id: int,
    row_num: int,
    language_used: str,
    is_approved: bool,
    verdict_text: str,
) -> None:
    """Update Excel columns: Language, BeeCrowd Approved, and Justification."""
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active

    if cols_config["lang_col"]:
        sheet.cell(row=row_num, column=cols_config["lang_col"]).value = language_used

    if cols_config["approved_col"]:
        sheet.cell(row=row_num, column=cols_config["approved_col"]).value = "Yes" if is_approved else "No"

    if cols_config["just_col"]:
        sheet.cell(row=row_num, column=cols_config["just_col"]).value = "" if is_approved else verdict_text

    wb.save(excel_path)
    wb.close()
    print(f"[Problem {problem_id}] Excel updated -> Approved: {'Yes' if is_approved else 'No'} | Justification: '{verdict_text if not is_approved else ''}'")


def login_to_beecrowd(page) -> bool:
    """Navigate to Beecrowd login page and authenticate."""
    print(f"Navigating to {BEECROWD_URL}...")
    page.goto(BEECROWD_URL, wait_until="domcontentloaded")

    try:
        print("Waiting for login form...")
        page.wait_for_selector('input[name="email"]', timeout=10000)

        print("Entering credentials...")
        page.fill('input[name="email"]', BEECROWD_EMAIL)
        page.fill('input[name="password"]', BEECROWD_PASSWORD)

        print("Submitting credentials...")
        page.click('button[type="submit"], input[type="submit"]')

        page.wait_for_url(lambda url: "/login" not in url, timeout=15000)
        print("Successfully logged in!")
        return True

    except Exception as error:
        print(f"Login failed or took too long: {error}")
        return False


def select_language(page, file_ext: str, category: str) -> str:
    """Select Java or PostgreSQL/SQL from dropdown using JavaScript."""
    is_sql = (file_ext.lower() == ".sql") or (category.upper() == "SQL")

    selected_name = page.evaluate(
        """
        (isSql) => {
            const select = document.querySelector("select[name='language_id'], select#language-id");
            if (!select) return null;

            let targetVal = null;
            let targetText = null;

            for (let opt of select.options) {
                const text = opt.text.toLowerCase();
                if (isSql) {
                    if (text.includes("postgresql") || text.includes("postgres") || text.includes("sql")) {
                        targetVal = opt.value;
                        targetText = opt.text;
                        break;
                    }
                } else {
                    if (text.includes("java")) {
                        targetVal = opt.value;
                        targetText = opt.text;
                        break;
                    }
                }
            }

            if (targetVal) {
                select.value = targetVal;
                select.dispatchEvent(new Event('change', { bubbles: true }));

                if (select.selectize) {
                    select.selectize.setValue(targetVal);
                }
                return targetText;
            }

            return null;
        }
    """,
        is_sql,
    )

    if selected_name:
        print(f"Selected language: {selected_name}")
        return selected_name
    else:
        fallback = "PostgreSQL" if is_sql else "Java"
        print(f"Language option not found directly; assuming default ({fallback}).")
        return fallback


def check_is_accepted(verdict_str: str) -> bool:
    """Strict evaluation of submission verdict."""
    s = verdict_str.strip().lower()
    
    # Negative patterns that mark a failure
    rejection_terms = [
        "errada", "wrong", "erro", "error", "excedido", 
        "exceeded", "presentation", "apresentação", "compilação", "compilation"
    ]
    if any(term in s for term in rejection_terms):
        return False

    # Positive approval confirmation
    return "aceito" in s or "accepted" in s


def get_latest_submission_status(page, problem_id: str, max_retries: int = 6) -> str:
    """Navigate to submission history and retrieve judgment result for the specific problem."""
    runs_url = "https://judge.beecrowd.com/pt/runs"
    print(f"[Problem {problem_id}] Checking judgment status at {runs_url}...")

    for attempt in range(max_retries):
        try:
            page.goto(runs_url, wait_until="domcontentloaded")
            page.wait_for_selector("table tbody tr", timeout=8000)

            result = page.evaluate(
                """
                (targetProbId) => {
                    const rows = Array.from(document.querySelectorAll("table tbody tr"));
                    for (let row of rows) {
                        const cells = Array.from(row.querySelectorAll("td")).map(td => td.innerText.trim());
                        const probText = cells[2] || '';
                        
                        // Ensure the row matches the target problem ID
                        if (probText.includes(targetProbId)) {
                            return {
                                problem: probText,
                                status: cells[4] || cells[3] || 'Unknown'
                            };
                        }
                    }
                    return null;
                }
            """,
                str(problem_id),
            )

            if result and result.get("status"):
                status = result["status"]
                print(f"[Problem {problem_id}] Status retrieved: '{status}' (Attempt {attempt + 1}/{max_retries})")

                if any(p in status.lower() for p in ["em fila", "queue", "compilando", "compiling", "executando"]):
                    time.sleep(4)
                    continue

                return status

        except Exception as err:
            print(f"[Problem {problem_id}] Notice while checking status: {err}")

        time.sleep(3)

    return "Pending / Unknown"


def submit_solution(page, problem_id: str, code: str, file_ext: str, category: str) -> tuple[str, bool, str]:
    """Navigate to problem page, inject solution, submit, and return (status, is_accepted, language_used)."""
    problem_url = f"https://judge.beecrowd.com/pt/problems/view/{problem_id}"
    print(f"\n[Problem {problem_id}] Navigating to {problem_url}...")

    page.goto(problem_url, wait_until="domcontentloaded")

    try:
        editor_selector = "#source-code, textarea[name='source_code'], .ace_editor"
        page.wait_for_selector(editor_selector, state="attached", timeout=10000)

        language_used = select_language(page, file_ext, category)

        print(f"[Problem {problem_id}] Injecting solution code...")
        inserted = page.evaluate(
            """
            (codeText) => {
                const textarea = document.querySelector('#source-code') || document.querySelector('textarea[name="source_code"]');
                const aceContainer = document.querySelector('.ace_editor');

                if (aceContainer && window.ace) {
                    const editor = ace.edit(aceContainer);
                    editor.setValue(codeText, -1);
                    if (textarea) textarea.value = codeText;
                    return true;
                }

                if (textarea) {
                    textarea.value = codeText;
                    textarea.dispatchEvent(new Event('input', { bubbles: true }));
                    textarea.dispatchEvent(new Event('change', { bubbles: true }));
                    return true;
                }

                return false;
            }
        """,
            code,
        )

        if not inserted:
            print(f"[Problem {problem_id}] Could not target code editor.")
            return "Editor Error", False, language_used

        time.sleep(1)

        print(f"[Problem {problem_id}] Clicking Submit...")
        submitted = page.evaluate(
            """
            () => {
                const submitBtn = document.querySelector('button[type="submit"]') ||
                                  document.querySelector('input[type="submit"]') ||
                                  document.querySelector('#btn-submit') ||
                                  Array.from(document.querySelectorAll('button, input')).find(el => el.value === 'Enviar' || el.innerText.trim() === 'Enviar');

                if (submitBtn) {
                    submitBtn.click();
                    return true;
                }
                return false;
            }
        """
        )

        if submitted:
            time.sleep(4)
            status = get_latest_submission_status(page, problem_id)
            is_accepted = check_is_accepted(status)
            return status, is_accepted, language_used
        else:
            print(f"[Problem {problem_id}] Could not locate submit button.")
            return "Submit Button Not Found", False, language_used

    except Exception as error:
        print(f"[Problem {problem_id}] Submission error: {error}")
        return f"Error: {error}", False, "Unknown"


def find_solutions(base_dir: str, problem_map: dict):
    """Locate solution files matching problem IDs in Excel."""
    solutions = []
    base_path = Path(base_dir)

    valid_extensions = {".java", ".sql", ".py", ".cpp", ".c"}

    for root, _, files in os.walk(base_path):
        for file in files:
            ext = Path(file).suffix.lower()
            if ext not in valid_extensions:
                continue

            full_path = Path(root) / file

            folder_match = re.search(r"(?:problem_)?(\d+)", Path(root).name, re.IGNORECASE)
            file_match = re.search(r"(\d+)", file)

            problem_id = None
            if folder_match:
                problem_id = folder_match.group(1)
            elif file_match:
                problem_id = file_match.group(1)

            if problem_id:
                prob_int = int(problem_id)
                excel_info = problem_map.get(prob_int, {})

                solutions.append(
                    {
                        "problem_id": problem_id,
                        "prob_int": prob_int,
                        "filepath": full_path,
                        "extension": ext,
                        "category": excel_info.get("category", ""),
                        "row": excel_info.get("row"),
                        "approved": excel_info.get("approved", ""),
                    }
                )

    return solutions


def process_solution_files(page, problem_map: dict, cols_config: dict) -> None:
    """Find, submit all solutions, and dynamically update Excel."""
    solutions = find_solutions(SOLUTIONS_DIR, problem_map)

    if not solutions:
        print(f"No solution files found in '{SOLUTIONS_DIR}'.")
        return

    print(f"Found {len(solutions)} solution file(s) to process.")

    for item in solutions:
        if item["approved"].strip().lower() == "yes":
            print(f"\n[Problem {item['problem_id']}] Already marked as 'Yes' in Excel. Skipping.")
            continue

        with open(item["filepath"], "r", encoding="utf-8") as f:
            code = f.read()

        status, is_accepted, language_used = submit_solution(
            page=page,
            problem_id=item["problem_id"],
            code=code,
            file_ext=item["extension"],
            category=item["category"],
        )

        if item["row"]:
            update_excel_row(
                excel_path=EXCEL_FILE,
                cols_config=cols_config,
                problem_id=item["prob_int"],
                row_num=item["row"],
                language_used=language_used,
                is_approved=is_accepted,
                verdict_text=status,
            )


def main() -> None:
    """Main execution entry point."""
    validate_environment()

    problem_map, cols_config = load_excel_problem_map(EXCEL_FILE)
    print(f"Loaded {len(problem_map)} problem entries from {EXCEL_FILE}.")

    with sync_playwright() as p:
        print("Launching Firefox...")
        browser = p.firefox.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        if login_to_beecrowd(page):
            process_solution_files(page, problem_map, cols_config)

        print("\nProcess finished. Closing browser...")
        time.sleep(3)
        browser.close()


if __name__ == "__main__":
    main()
