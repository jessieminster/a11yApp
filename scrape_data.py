import os

# Ensure python-docx is installed
try:
    os.system("pip install python-docx")
except ImportError:
    print("Could not install python-docx...")

from docx import Document
import time
import win32com.client
# Open a Word document and run the accessibility checker
def run_accessibility_checker(file_path):
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    # pythoncom.CoInitialize()

    # Open the Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True  # Run in the background

    try:
        # Open the document
        doc = word.Documents.Open(file_path)
        version = word.Version
        print(f"word active doc: {word.ActiveDocument}")
        print(f"version: {version}")
        # Run the accessibility checker
        # word.ActiveDocument.CheckAccessibility()
        # word.Run("CheckAccessibility")
        #word.CommandBars.ExecuteMso("ReviewAccessibilityChecker")
        word.CommandBars.ExecuteMso("AccessibilityChecker")
        time.sleep(5)

        # try:
        #     print("Starting check")
        #     results = []
        #     for i in range(1, word.TaskPanes.Count + 1):
        #         print(f"pane is {word.TaskPanes.Item(i)}")
                # try:
                #     if hasattr(pane, 'Title') and 'Accessibility' in str(pane.Title):
                #         print(f"Found accessibility pane: {pane.Title}")

                #         if hasattr(pane, 'HTMLDocument'):
                #             html_doc = pane.HTMLDocument
                #             print(f"doc= {html_doc}")
                #             results.append("found accessibility pane")
                #         break
                # except Exception as pane_error:
                #     print(f"pane error {pane_error}")

            # return results if results else "No pane found"
            # checker = doc.Range().AccessibilityChecker
            # issues = []

            # for issue in checker.Issues:
            #     issue_data = {
            #         'title': str(issue.Title),
            #         'description': str(issue.Description),
            #         'severity': str(issue.Severity),
            #         'rule_type': str(issue.Rule),
            #         'location': str(issue.Location) if hasattr(issue, 'Location') else 'Unknown'
            #     }
            #     issues.append(issue_data)

            # print(f"issues are: {issues}")
            # return issues
        # except Exception as e: 
        #     print(f"Could not access results object: {e}")

        #     try: 
        #         checker = doc.AccessibilityChecker
        #         issues = []

        #         for i in range(checker.Issues.Count):
        #             issue = checker.Issues.Item(i + 1)
        #             issue_data = {
        #                 'title': str(issue.Title),
        #                 'description': str(issue.Description),
        #                 'severity': str(issue.Severity),
        #             }
        #             issues.append(issue_data)
        #         return issues
        #     except Exception as e2:
        #         print(f"Could not retrieve results: {e2}")

        print("Accessibility checker has been run. Please review the results in Word.")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Close the document and quit Word
        doc.Close(SaveChanges=False)
        word.Quit()

# Example usage
if __name__ == "__main__":
    file_path = "C:\\Users\\JessieMinster\\Desktop\\A11y\\a11yApp\\Minster_Resume.docx"  # Replace with your Word file path
    run_accessibility_checker(file_path)
