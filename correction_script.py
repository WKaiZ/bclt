import os
import pyautogui
import pyperclip
from openpyxl import load_workbook
import re

def read_lines_to_list(file_path):
    with open(file_path, 'r') as file:
        lines = [line.rstrip() for line in file]
    return lines

def find_button(image_path, confidence = 0.8):
    button_location = pyautogui.locateOnScreen(image_path, confidence=confidence)
    if button_location is not None:
        return button_location
    else:
        print("Button not found.")

def move_mouse_to_button_and_click(image_path):
    button_location = find_button(image_path)
    if button_location is not None:
        x, y = pyautogui.center(button_location)
        pyautogui.moveTo(x, y, duration=0.5)
        pyautogui.click()
    else:
        print("No button location provided.")

def type_into_input_box(text, interval = 0.1):
    pyautogui.typewrite(text, interval)

def paste_text(text):
    pyperclip.copy(text)
    pyautogui.hotkey('ctrl', 'v')

def detect_error(image_path, confidence = 0.7):
    try:
        pyautogui.locateOnScreen(image_path, confidence=confidence)
        return True
    except Exception as e:
        return False

def wait_until_complete(image_path, confidence = 0.95):
    while True:
        try:
            if detect_error("error.png"):
                return False
            button_location = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if button_location is not None:
                return True
        except Exception as e:
            continue

def find_cell_coordinates(file_path, sheet_name, target_string):
    wb = load_workbook(filename=file_path)
    sheet = wb[sheet_name]
    
    for row in sheet.iter_rows(values_only=False):
        for cell in row:
            if cell.value == target_string:
                return (cell.row, cell.column)
    return None


citizens = read_lines_to_list('bad_summaries_petitions.txt')
responses  = read_lines_to_list('bad_summaries_responses.txt')
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.7
prompt = """Extract the following details from the FDA response document above: 1. Date of Response: Identify the date on which the FDA issued the response. 2. Responding FDA Center: Specify which FDA center or division responded to the petition (e.g., CDER, CBER, CDRH). 3. Response to Petition: Indicate the FDA's decision or action taken in response to the petition (e.g., approved, denied, partially approved, other). 4. Cited Statutes or Regulations: List all statutes or regulations cited by the FDA in its response. 5. Justification for Response: Summarize the reasoning or justifications provided by the FDA to support its decision. Ensure that the extracted information is accurate, relevant, and organized in a structured format, such as a bullet list or table. If any information is not explicitly stated, indicate 'Not Mentioned'. Remove markdown and put each numbered item in a Python list. Do not use nested lists. Do not include any other text such as comments or explanations."""

for file_name in responses:
    year = re.findall(r'\d{4}', file_name)[0]
    move_mouse_to_button_and_click("prompt.png")
    paste_text(prompt)
    move_mouse_to_button_and_click("upload.png")
    move_mouse_to_button_and_click("upload_from_computer.png")
    paste_text(f"C:\\Users\\wesle\\OneDrive\\Desktop\\bclt\\{year}\\{file_name}")
    pyautogui.hotkey('enter')
    error = wait_until_complete("complete.png")
    if not error:
        pyautogui.moveTo(1153, 149, duration=0.5)
        pyautogui.click()
        pyautogui.moveTo(1100, 800, duration=0.5)
        pyautogui.click()
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('backspace')
        continue
    move_mouse_to_button_and_click("ask.png")
    wait_until_complete("voice.png")
    move_mouse_to_button_and_click("python_copy.png")
    lst = eval(pyperclip.paste())
    wb = load_workbook("outputs_responses.xlsx")
    ws = wb[year]
    r, _ = find_cell_coordinates("outputs_responses.xlsx", year, file_name)
    for i, item in enumerate(lst):
        if item == "Not Mentioned":
            print("Not Mentioned")
        ws.cell(row=r, column= i + 2).value = item
    wb.save("outputs_responses.xlsx")
    wb.close()
    print("Completed file: ", file_name)