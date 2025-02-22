import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import os
import re

# Constants for colors
ARTICLE_CONV = {
    "m": "der",  
    "f": "die",  
    "nt": "das", 
    "": "",
}

ARTICLE_COLORS = {
    "der": "0000FF",  # Blue
    "die": "FF0000",  # Red
    "das": "008000",  # Green
    "": "8B4513"      # Brown for no article
}

TERMINAL_COLORS = {
    "der": "\033[1;34m",  # Blue
    "die": "\033[1;31m",  # Red
    "das": "\033[1;32m",  # Green
    "": "\033[38;2;165;42;42m"      # Brown for no article
}

RESET = "\033[0m"

EXCEL_FILE = "German_Words.xlsx"


def get_word_data(word):
    """Fetches the article and definitions of a word from PONS."""
    url = f"https://en.pons.com/translate/german-english/{word.lower()}"
    response = requests.get(url)
    
    if response.status_code != 200:
        print("Error: Unable to fetch data from PONS.")
        return None, None

    soup = BeautifulSoup(response.text, "html.parser")

    # Try to extract the article and definitions
    try:
        article_tag = soup.find("span", class_="genus")
        article = article_tag.text.strip() if article_tag else ""

        # Exclude "seealso" sections
        seealso_sections = soup.find_all("div", class_="seealso")
        for section in seealso_sections:
            section.decompose()  # Remove the seealso section

        definition_tags = soup.find_all("div", class_="translations")
        definitions = []
        for tag in definition_tags:
            target_tag = tag.find_all("div", class_="target", limit=2)
            for target in target_tag:
                definition = target.get_text().strip()
                definitions.append(definition)
                
        definitions = list(dict.fromkeys(definitions))  # Remove duplicates

        if not definitions:
            return article, ""

        # Allow user to select the appropriate definition
        print("Multiple definitions found:")
        for i, definition in enumerate(definitions, start=1):
            print(f"{i}. {definition}")

        while True:
            try:
                choice = int(input("Select the appropriate definition (enter the number): ").strip())
                if 0 < choice <= len(definitions):
                    selected_definition = definitions[choice - 1]
                    break
                else:
                    print("Invalid choice: Please enter a number within the range.")
            except ValueError:
                print("Invalid input: Please enter a number for the definition choice.")

        return article, selected_definition
    except AttributeError:
        print("Warning: Could not determine the article or definitions.")
        return "", ""


def create_or_load_excel():
    """Creates a new Excel file if it doesn't exist, otherwise loads it."""
    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "General"
        ws.append(["Word", "Definition"])
        bold_font = Font(bold=True, size=16)
        alignment_center = Alignment(horizontal="center", vertical="center")
        for col in range(1, 3):
            ws.cell(row=1, column=col).font = bold_font
            ws.cell(row=1, column=col).alignment = alignment_center

        # Create sheets for each article
        for article in ["der", "die", "das", "No Article"]:
            ws = wb.create_sheet(title=article)
            ws.append(["Word", "Definition"])
            for col in range(1, 3):
                ws.cell(row=1, column=col).font = bold_font
                ws.cell(row=1, column=col).alignment = alignment_center

        wb.save(EXCEL_FILE)
    return wb


def check_duplicate(word, definition, wb):
    """Checks if a word already exists in the Excel file."""
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0].lower() == word.lower() and row[1].lower() == definition.lower():
                return row[0], row[1]  # Return word and definition
    return None, None


def create_lesson_sheet(wb, lesson):
    """Creates a new lesson sheet if it doesn't exist."""
    lesson_sheet_name = f"A1.1 - L{lesson}"
    if lesson_sheet_name not in wb.sheetnames:
        lesson_ws = wb.create_sheet(title=lesson_sheet_name)
        lesson_ws.append(["Word", "Definition"])
        bold_font = Font(bold=True, size=16)
        alignment_center = Alignment(horizontal="center", vertical="center")
        for col in range(1, 3):
            lesson_ws.cell(row=1, column=col).font = bold_font
            lesson_ws.cell(row=1, column=col).alignment = alignment_center
    else:
        lesson_ws = wb[lesson_sheet_name]
    return lesson_ws

def sort_and_color_sheet(ws):
    """Sorts the sheet alphabetically by the word itself, ignoring the articles, and optionally changes the colors."""
    data = list(ws.iter_rows(min_row=2, values_only=True))
    sorted_data = sorted(data, key=lambda x: x[0].split(" ", 1)[1].lower() if " " in x[0] else x[0].lower())
    for idx, row in enumerate(sorted_data, start=2):
        for col, value in enumerate(row, start=1):
            ws.cell(row=idx, column=col, value=value)
            if col == 1:  # Apply color to the word cell
                checked_article = value.split(" ", 1)[0]
                color = ARTICLE_COLORS.get(checked_article)
                ws.cell(row=idx, column=col).fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                ws.cell(row=idx, column=col).font = Font(color="FFFFFF", bold=True)

def add_word_to_sheet(ws, full_word, definition):
    """Adds a word to a specific sheet and sorts it."""
    ws.append([full_word, definition])
    sort_and_color_sheet(ws)

def add_word_to_excel(word, article, definition, lesson, wb):
    """Adds a word to the general sheet, the relevant article sheet, and the relevant lesson sheet in the Excel file."""      
    # Combine the article and word
    real_article = ARTICLE_CONV.get(article)
    full_word = f"{real_article} {word}"

    # Color the terminal output based on the article
    terminal_color = TERMINAL_COLORS.get(real_article, RESET)

    # Check for duplicates in the general sheet
    existing_word, existing_definition = check_duplicate(full_word, definition, wb)
    if existing_word:        
        print(f"⚠️ The word already exists! {terminal_color}{existing_word}{RESET} : {existing_definition}")
        return

    # Add the word to the general sheet
    general_ws = wb["General"]
    add_word_to_sheet(general_ws, full_word, definition)

    # Add the word to the relevant article sheet
    article_sheet_name = real_article if real_article else "No Article"
    article_ws = wb[article_sheet_name]
    add_word_to_sheet(article_ws, full_word, definition)

    # Add the word to the relevant lesson sheet
    lesson_ws = create_lesson_sheet(wb, lesson)
    add_word_to_sheet(lesson_ws, full_word, definition)

    wb.save(EXCEL_FILE)
    print(f"✅ Added '{terminal_color}{full_word}{RESET}' : {definition}.")


if __name__ == "__main__":
    wb = create_or_load_excel()

    while True:
        word = input("Enter the word (or 'q' to quit): ").strip()
        if word == "q":
            break

        article, definition = get_word_data(word)
        if not definition:
            print("Skipping word due to missing data.")
            continue

        while True:
            try:
                lesson = int(input("Which lesson did you learn this word in? ").strip())
                break
            except ValueError:
                print("Invalid input: Please enter a number for the lesson.")
        
        add_word_to_excel(word, article, definition, lesson, wb)
