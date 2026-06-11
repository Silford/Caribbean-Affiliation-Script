from tkinter import Tk, filedialog
from pathlib import Path
import os
import sys

from fix_encoding_issues import fix_encoding
from url_search import process_url_file
from doi_search import process_doi_file


COLORS = {
    "reset": "\033[0m",
    "bold": "\033[1m",
    "dim": "\033[2m",
    "cyan": "\033[36m",
    "magenta": "\033[35m",
    "green": "\033[32m",
    "yellow": "\033[33m",
    "red": "\033[31m",
    "blue": "\033[34m",
}

APP_NAME = "Caribbean Affiliation Script"
APP_VERSION = "1.0.0"
APP_SUBTITLE = "Excel cleanup, DOI enrichment, and URL affiliation review"


def supports_color():
    return sys.stdout.isatty() and os.environ.get("NO_COLOR") is None


def supports_unicode():
    encoding = (sys.stdout.encoding or "").lower()
    return "utf" in encoding


def style(text, *styles):
    if not supports_color():
        return text

    prefix = "".join(COLORS[item] for item in styles)
    return f"{prefix}{text}{COLORS['reset']}"


def clear_screen():
    os.system("cls" if os.name == "nt" else "clear")


def box_chars():
    if supports_unicode():
        return {
            "tl": "╭", "tr": "╮", "bl": "╰", "br": "╯",
            "h": "─", "v": "│", "sep_l": "├", "sep_r": "┤",
        }

    return {
        "tl": "+", "tr": "+", "bl": "+", "br": "+",
        "h": "-", "v": "|", "sep_l": "+", "sep_r": "+",
    }


def print_rule(width=72, color="cyan"):
    chars = box_chars()
    print(style(chars["h"] * width, color, "dim"))


def print_panel(title, lines, accent="cyan", width=72):
    chars = box_chars()
    width = max(width, len(title) + 4, *(len(line) + 4 for line in lines))
    top = chars["tl"] + chars["h"] * (width - 2) + chars["tr"]
    separator = chars["sep_l"] + chars["h"] * (width - 2) + chars["sep_r"]
    bottom = chars["bl"] + chars["h"] * (width - 2) + chars["br"]

    print(style(top, accent))
    print(style(f"{chars['v']} {title.center(width - 4)} {chars['v']}", accent, "bold"))
    print(style(separator, accent, "dim"))

    for line in lines:
        print(f"{style(chars['v'], accent)} {line.ljust(width - 4)} {style(chars['v'], accent)}")

    print(style(bottom, accent))


def print_header():
    print()
    print(style(APP_NAME.upper(), "cyan", "bold"))
    print(style(f"v{APP_VERSION}  |  {APP_SUBTITLE}", "dim"))
    print_rule()


def menu_line(number, label, description):
    return f" {number:>2}  {label.ljust(28)} {description}"


def print_status(message, status="info"):
    labels = {
        "info": ("INFO", "blue"),
        "success": ("OK", "green"),
        "warning": ("WARN", "yellow"),
        "error": ("ERR", "red"),
    }
    label, color = labels[status]
    print(f"{style(f'[{label}]', color, 'bold')} {message}")


def print_operation(title, input_file, output_file):
    print()
    print_panel(
        title,
        [
            f"Input : {input_file}",
            f"Output: {output_file}",
            "",
            "Progress will update as work completes.",
        ],
        accent="magenta",
    )


def pause():
    input(style("\nPress Enter to return to the main menu...", "dim"))


def pick_excel_file():
    root = Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )

    root.destroy()
    return file_path


def get_output_path(default_name="output.xlsx"):
    prompt = style("Output filename", "cyan", "bold")
    output_name = input(f"{prompt} [{default_name}]: ").strip()

    if not output_name:
        output_name = default_name

    if not output_name.endswith(".xlsx"):
        output_name += ".xlsx"

    return str(Path.cwd() / output_name)


def fix_encoding_menu():
    input_file = pick_excel_file()

    if not input_file:
        print_status("No file selected.", "warning")
        return

    output_file = get_output_path("output.xlsx")

    print_operation("Fix Encoding", input_file, output_file)
    print_status("Fixing encoding issues...", "info")
    fix_encoding(input_file, output_file)

    print_status(f"Saved cleaned workbook to: {output_file}", "success")


def doi_search_menu():
    input_file = pick_excel_file()

    if not input_file:
        print_status("No file selected.", "warning")
        return

    output_file = get_output_path("output.xlsx")

    print_operation("DOI Affiliation Search", input_file, output_file)
    print_status("Searching DOI affiliation data...", "info")
    process_doi_file(input_file, output_file)

    print_status(f"Saved DOI results to: {output_file}", "success")


def url_search_menu():
    input_file = pick_excel_file()

    if not input_file:
        print_status("No file selected.", "warning")
        return

    output_file = get_output_path("output.xlsx")

    print_operation("URL Affiliation Search", input_file, output_file)
    print_status("Searching URL affiliation data...", "info")
    process_url_file(input_file, output_file)

    print_status(f"Saved URL results to: {output_file}", "success")


def print_main_menu():
    clear_screen()
    print_header()
    print_panel(
        "Main Menu",
        [
            "Choose a workflow for your Excel file.",
            "",
            menu_line(1, "Fix encoding issues", "Clean mojibake in Title and Authors"),
            menu_line(2, "Search DOI affiliations", "Query DOI metadata and flag Caribbean links"),
            menu_line(3, "Search URL affiliations", "Inspect URLs, DOI metadata, and pages"),
            menu_line(4, "Exit", "Close the application"),
        ],
    )


def main():
    while True:
        print_main_menu()

        choice = input(style("\nSelect an option: ", "cyan", "bold")).strip()

        if choice == "1":
            run_operation(fix_encoding_menu)
            pause()
        elif choice == "2":
            run_operation(doi_search_menu)
            pause()
        elif choice == "3":
            run_operation(url_search_menu)
            pause()
        elif choice == "4":
            print_status("Goodbye.", "success")
            break
        else:
            print_status("Invalid option. Choose 1, 2, 3, or 4.", "error")
            pause()


def run_operation(callback):
    try:
        callback()
    except Exception as exc:
        print_status(f"Operation failed: {exc}", "error")


if __name__ == "__main__":
    main()
