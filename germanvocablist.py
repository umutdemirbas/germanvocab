import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill
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

EXCEL_FILE = "German_Words.xlsx"


def get_word_data(word):
    """Fetches the article and definition of a word from PONS."""
    url = f"https://en.pons.com/translate/german-english/{word.lower()}"
    response = requests.get(url)
    
    if response.status_code != 200:
        print("Error: Unable to fetch data from PONS.")
        return None, None

    soup = BeautifulSoup(response.text, "html.parser")

    # Try to extract the article and definition
    try:
        article_tag = soup.find("span", class_="genus")
        article = article_tag.text.strip() if article_tag else ""

        definition_tag = soup.find("div", class_="entry")
        definition = definition_tag.find("div", class_="target").text.strip().split("\n")[0] if definition_tag else ""

        return article, definition
    except AttributeError:
        print("Warning: Could not determine the article or definition.")
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
        bold_font = Font(bold=True)
        for col in range(1, 3):
            ws.cell(row=1, column=col).font = bold_font

        # Create sheets for each article
        for article in ["der", "die", "das", "No Article"]:
            ws = wb.create_sheet(title=article)
            ws.append(["Word", "Definition"])
            for col in range(1, 3):
                ws.cell(row=1, column=col).font = bold_font

        wb.save(EXCEL_FILE)
    return wb


def check_duplicate(word, wb):
    """Checks if a word already exists in the Excel file."""
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0].lower() == word.lower():
                return row[0], row[1]  # Return article and definition
    return None, None


def create_lesson_sheet(wb, lesson):
    """Creates a new lesson sheet if it doesn't exist."""
    lesson_sheet_name = f"Lesson {lesson}"
    if lesson_sheet_name not in wb.sheetnames:
        lesson_ws = wb.create_sheet(title=lesson_sheet_name)
        lesson_ws.append(["Word", "Definition"])
        bold_font = Font(bold=True)
        for col in range(1, 3):
            lesson_ws.cell(row=1, column=col).font = bold_font
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

    # Check for duplicates in the general sheet
    existing_article, existing_definition = check_duplicate(full_word, wb)
    if existing_article:
        print(f"⚠️ '{word}' already exists! ({existing_article}) - {existing_definition}")
        return

    # Add the word to the general sheet
    general_ws = wb["General"]
    add_word_to_sheet(general_ws, full_word, definition)

    # Add the word to the relevant article sheet
    article_ws = wb[real_article]
    add_word_to_sheet(article_ws, full_word, definition)

    # Add the word to the relevant lesson sheet
    lesson_ws = create_lesson_sheet(wb, lesson)
    add_word_to_sheet(lesson_ws, full_word, definition)

    wb.save(EXCEL_FILE)
    print(f"✅ Added '{full_word}' - {definition} under lesson {lesson}.")


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

        lesson = int(input("Which lesson did you learn this word in? ").strip())
        add_word_to_excel(word, article, definition, lesson, wb)
