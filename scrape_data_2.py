import os
import time
import win32com.client
import pywinauto
from pywinauto import Application
from pywinauto.findwindows import find_windows
import re
from datetime import datetime

class WordAccessibilityScraper:
    def __init__(self):
        self.app = None
        self.word_window = None
        self.accessibility_pane = None
        
    def connect_to_word(self):
        """Connect to an existing Word application"""
        try:
            # Find Word windows
            word_windows = find_windows(class_name="OpusApp")
            if not word_windows:
                raise Exception("No Word windows found. Please open Word first.")
            
            # Connect to Word application
            self.app = Application().connect(handle=word_windows[0])
            self.word_window = self.app.window(handle=word_windows[0])
            print("Successfully connected to Word for GUI automation")
            return True
            
        except Exception as e:
            print(f"Error connecting to Word for GUI automation: {str(e)}")
            return False
    
    def wait_for_accessibility_checker(self, timeout=15):
        """Wait for the accessibility checker to appear and complete"""
        print("Waiting for accessibility checker to complete analysis...")
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            if self.find_accessibility_pane():
                # Wait a bit more for the analysis to complete
                time.sleep(2)
                return True
            time.sleep(1)
        
        print("Timeout waiting for accessibility checker")
        return False
    
    def find_accessibility_pane(self):
        """Find the accessibility checker task pane"""
        try:
            # Based on your output, look for "Accessibility Assistant" with different control types
            accessibility_control_types = [
                "MsoWorkPane",  # This appears in your output
                "MsoCommandBar", 
                "NetUINativeHWNDHost",
                "Pane",
                "NUIPane"
            ]
            
            # First try to find by exact title "Accessibility Assistant"
            for control_type in accessibility_control_types:
                try:
                    pane = self.word_window.child_window(title="Accessibility Assistant", control_type=control_type)
                    if pane.exists() and pane.is_visible():
                        self.accessibility_pane = pane
                        for child in pane.children():
                            print(f"Child: {child.window_text()}, Type: {child.element_info.control_type}")
                        print(f"Found accessibility pane with control_type: {control_type}")
                        return True
                except Exception as e:
                    print(f"Tried {control_type}: {str(e)}")
                    continue
            
            # Alternative: Look through all children for accessibility-related titles
            try:
                print("Searching through all children for accessibility pane...")
                children = self.word_window.children()
                
                for child in children:
                    try:
                        title = child.window_text()
                        if "Accessibility" in title:
                            print(f"Found child with accessibility title: '{title}', control_type: {child.element_info.control_type}")
                            if child.is_visible():
                                self.accessibility_pane = child
                                for desc in child.children():
                                    print(f"Child of child: {desc.window_text()}, Type: {desc.element_info.control_type}")
                        
                                print(f"Selected visible accessibility pane: {title}")
                                return True
                    except Exception as e:
                        continue
                        
            except Exception as e:
                print(f"Error searching children: {str(e)}")
            
            # Last resort: Try to find any MsoWorkPane that might contain accessibility content
            try:
                print("Trying MsoWorkPane elements...")
                work_panes = self.word_window.children(control_type="MsoWorkPane")
                for pane in work_panes:
                    try:
                        if pane.is_visible():
                            pane_text = pane.window_text().lower()
                            print(f"MsoWorkPane text preview: {pane_text[:100]}...")
                            if any(word in pane_text for word in ['accessibility', 'error', 'warning', 'tip', 'issues']):
                                self.accessibility_pane = pane
                                print("Found accessibility content in MsoWorkPane")
                                return True
                    except:
                        continue
            except Exception as e:
                print(f"Error checking MsoWorkPane elements: {str(e)}")
                
            return False
            
        except Exception as e:
            print(f"Error finding accessibility pane: {str(e)}")
            return False
    
    def scrape_accessibility_results(self):
        """Scrape the accessibility checker results"""
        if not self.accessibility_pane:
            if not self.find_accessibility_pane():
                return None
        
        try:
            results = {
                'timestamp': datetime.now().isoformat(),
                'errors': [],
                'warnings': [],
                'tips': [],
                'summary': {},
                'raw_text': '',
                'all_found_text': []
            }
            
            # Get all text from the accessibility pane
            try:
                pane_text = self.accessibility_pane.window_text()
                results['raw_text'] = pane_text
                print(f"Main pane text: '{pane_text}' (length: {len(pane_text)})")
            except Exception as e:
                print(f"Error getting main pane text: {str(e)}")
            
            # Recursively get all text from child elements
            def get_all_child_text(element, depth=0):
                texts = []
                try:
                    # Get text from current element
                    element_text = element.window_text()
                    if element_text and element_text.strip():
                        texts.append(f"{'  ' * depth}[{element.element_info.control_type}] {element_text}")
                    
                    # Get text from children
                    children = element.children()
                    for child in children:
                        if child.is_visible():
                            child_texts = get_all_child_text(child, depth + 1)
                            texts.extend(child_texts)
                            
                except Exception as e:
                    texts.append(f"{'  ' * depth}[ERROR] {str(e)}")
                
                return texts
            
            # Get all text content
            print("Extracting all text from accessibility pane and children...")
            all_texts = get_all_child_text(self.accessibility_pane)
            results['all_found_text'] = all_texts
            
            # Print what we found for debugging
            print("Found text elements:")
            for text in all_texts[:10]:  # Print first 10 items
                print(f"  {text}")
            if len(all_texts) > 10:
                print(f"  ... and {len(all_texts) - 10} more items")
            
            # Analyze all found text for accessibility issues
            combined_text = ' '.join([text for text in all_texts if text])
            
            # Categorize based on keywords in all text
            for text_item in all_texts:
                if not text_item or not text_item.strip():
                    continue
                    
                lower_text = text_item.lower()
                clean_text = text_item.split('] ', 1)[-1] if '] ' in text_item else text_item
                
                if any(word in lower_text for word in ['error', 'critical', 'must fix']):
                    if clean_text not in results['errors']:
                        results['errors'].append(clean_text.strip())
                elif any(word in lower_text for word in ['warning', 'caution', 'should fix']):
                    if clean_text not in results['warnings']:
                        results['warnings'].append(clean_text.strip())
                elif any(word in lower_text for word in ['tip', 'suggestion', 'recommendation', 'consider']):
                    if clean_text not in results['tips']:
                        results['tips'].append(clean_text.strip())
            
            # Parse for summary information from combined text
            if combined_text:
                # Extract issue counts using regex
                error_matches = re.findall(r'(\d+)\s*error', combined_text, re.IGNORECASE)
                warning_matches = re.findall(r'(\d+)\s*warning', combined_text, re.IGNORECASE)
                tip_matches = re.findall(r'(\d+)\s*tip', combined_text, re.IGNORECASE)
                
                if error_matches:
                    results['summary']['error_count'] = max([int(x) for x in error_matches])
                if warning_matches:
                    results['summary']['warning_count'] = max([int(x) for x in warning_matches])
                if tip_matches:
                    results['summary']['tip_count'] = max([int(x) for x in tip_matches])
                
                # Look for status messages
                if any(phrase in combined_text.lower() for phrase in ['no issues', 'no accessibility issues', 'good to go', 'no problems']):
                    results['summary']['status'] = 'No issues found'
                elif any(phrase in combined_text.lower() for phrase in ['issues found', 'problems found', 'error', 'warning']):
                    results['summary']['status'] = 'Issues found'
                else:
                    results['summary']['status'] = 'Status unclear'
            
            # If we didn't find categorized issues but have text, add it as raw findings
            if not results['errors'] and not results['warnings'] and not results['tips'] and results['all_found_text']:
                results['summary']['note'] = f"Found {len(results['all_found_text'])} text elements but could not categorize them"
            
            return results
            
        except Exception as e:
            print(f"Error scraping results: {str(e)}")
            return None
    
    def save_results_to_file(self, results, filename, document_name=""):
        """Save the scraped results to a text file"""
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write("Word Accessibility Checker Results\n")
                f.write("=" * 50 + "\n\n")
                f.write(f"Document: {document_name}\n")
                f.write(f"Generated: {results['timestamp']}\n\n")
                
                # Summary
                if results['summary']:
                    f.write("SUMMARY:\n")
                    f.write("-" * 20 + "\n")
                    for key, value in results['summary'].items():
                        f.write(f"{key.replace('_', ' ').title()}: {value}\n")
                    f.write("\n")
                
                # Errors
                if results['errors']:
                    f.write(f"ERRORS ({len(results['errors'])}):\n")
                    f.write("-" * 20 + "\n")
                    for i, error in enumerate(results['errors'], 1):
                        f.write(f"{i}. {error}\n")
                    f.write("\n")
                else:
                    f.write("ERRORS: None found\n\n")
                
                # Warnings
                if results['warnings']:
                    f.write(f"WARNINGS ({len(results['warnings'])}):\n")
                    f.write("-" * 20 + "\n")
                    for i, warning in enumerate(results['warnings'], 1):
                        f.write(f"{i}. {warning}\n")
                    f.write("\n")
                else:
                    f.write("WARNINGS: None found\n\n")
                
                # Tips
                if results['tips']:
                    f.write(f"TIPS ({len(results['tips'])}):\n")
                    f.write("-" * 20 + "\n")
                    for i, tip in enumerate(results['tips'], 1):
                        f.write(f"{i}. {tip}\n")
                    f.write("\n")
                else:
                    f.write("TIPS: None found\n\n")
                
                # All found text for debugging
                if results.get('all_found_text'):
                    f.write("ALL EXTRACTED TEXT ELEMENTS:\n")
                    f.write("-" * 30 + "\n")
                    for i, text in enumerate(results['all_found_text'], 1):
                        f.write(f"{i}. {text}\n")
                    f.write("\n")
                
                # Raw text for debugging/reference
                if results['raw_text']:
                    f.write("RAW ACCESSIBILITY PANE CONTENT:\n")
                    f.write("-" * 30 + "\n")
                    f.write(results['raw_text'])
                    f.write("\n\n")
            
            print(f"Results saved to {filename}")
            return True
            
        except Exception as e:
            print(f"Error saving results: {str(e)}")
            return False

def run_accessibility_checker(file_path):
    """Open a Word document and run the accessibility checker with GUI scraping"""
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return
    
    # Initialize the scraper
    scraper = WordAccessibilityScraper()
    
    # Open the Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True  # Make Word visible so we can scrape the GUI
    
    try:
        print(f"Opening document: {os.path.basename(file_path)}")
        
        # Open the document
        doc = word.Documents.Open(file_path)
        version = word.Version
        print(f"Word version: {version}")
        print(f"Active document: {word.ActiveDocument.Name}")
        
        # Connect to Word for GUI automation
        if not scraper.connect_to_word():
            print("Failed to connect for GUI automation, but continuing with basic functionality")
        
        # Run the accessibility checker
        print("Running accessibility checker...")
        try:
            word.CommandBars.ExecuteMso("AccessibilityChecker")
        except:
            # Try alternative command
            try:
                word.CommandBars.ExecuteMso("ReviewAccessibilityChecker")
            except:
                print("Could not execute accessibility checker command")
                return
        
        # Wait for the accessibility checker to complete
        if scraper.word_window:
            if scraper.wait_for_accessibility_checker():
                print("Accessibility checker completed. Scraping results...")
                
                # Scrape the results
                results = scraper.scrape_accessibility_results()
                
                if results:
                    # Generate output filename
                    base_name = os.path.splitext(os.path.basename(file_path))[0]
                    output_dir = os.path.dirname(file_path)
                    results_file = os.path.join(output_dir, f"{base_name}_accessibility_results.txt")
                    
                    # Save results
                    if scraper.save_results_to_file(results, results_file, os.path.basename(file_path)):
                        print(f"\n✅ Accessibility check completed successfully!")
                        print(f"📁 Results saved to: {results_file}")
                        
                        # Print summary
                        if results['summary']:
                            print(f"📊 Summary:")
                            for key, value in results['summary'].items():
                                print(f"   {key.replace('_', ' ').title()}: {value}")
                    else:
                        print("❌ Failed to save results to file")
                        
                else:
                    print("❌ Failed to scrape accessibility results")
                    print("The accessibility pane might not be visible or have unexpected content")
            else:
                print("⚠️  Could not detect accessibility checker completion")
                print("Please check the results manually in Word")
        else:
            print("⚠️  GUI automation not available, accessibility checker opened but results not scraped")
            print("Please review the results manually in Word")
            time.sleep(5)  # Give user time to see the results
        
    except Exception as e:
        print(f"❌ An error occurred: {e}")
        
    finally:
        # Close the document and quit Word
        try:
            doc.Close(SaveChanges=False)
            word.Quit()
            print("Word application closed")
        except:
            print("Error closing Word application")

# Example usage
if __name__ == "__main__":
    # Install required packages if not already installed
    try:
        import pywinauto
    except ImportError:
        print("pywinauto not found. Please install it with: pip install pywinauto")
        exit(1)
    
    file_path = "C:\\Users\\JessieMinster\\Desktop\\A11y\\a11yApp\\ConflictDoc.docx"
    
    print("Word Accessibility Checker with GUI Scraping")
    print("=" * 50)
    
    run_accessibility_checker(file_path)