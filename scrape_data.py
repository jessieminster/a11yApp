import os

# Ensure python-docx is installed
try:
    os.system("pip install python-docx")
except ImportError:
    print("Could not install python-docx...")

from docx import Document

import win32com.client
# Open a Word document and run the accessibility checker
def run_accessibility_checker(file_path):
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    # Open the Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Run in the background

    try:
        # Open the document
        doc = word.Documents.Open(file_path)

        # Run the accessibility checker
        word.ActiveDocument.CheckAccessibility()

        print("Accessibility checker has been run. Please review the results in Word.")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Close the document and quit Word
        doc.Close(SaveChanges=False)
        word.Quit()

# Example usage
if __name__ == "__main__":
    file_path = "./Minster_Resume.docx"  # Replace with your Word file path
    run_accessibility_checker(file_path)