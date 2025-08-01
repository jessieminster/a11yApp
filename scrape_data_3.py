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

    def print_accessibility_pane_elements(self):
    """Recursively print all elements in the accessibility pane."""
    if not self.accessibility_pane:
        if not self.find_accessibility_pane():
            print("Accessibility pane not found.")
            return

    def print_elements(element, depth=0):
        try:
            text = element.window_text()
            control_type = element.element_info.control_type
            print(f"{'  '*depth}- [{control_type}] '{text}'")
            for child in element.children():
                print_elements(child, depth+1)
        except Exception as e:
            print(f"{'  '*depth}Error: {e}")
    
    def find_accessibility_pane(self):
        """Find the accessibility checker task pane"""
        try:
            # Based on your output, look for "Accessibility Assistant" with different control types
            accessibility_control_types = [
                "MsoWorkPane",  # This appears in your output
            ]
            print(f"children are {self.word_window.children()}")
            # First try to find by exact title "Accessibility Assistant"
            for control_type in accessibility_control_types:
                print(control_type)
                try:
                    pane = self.word_window.child_window(title="Accessibility Assistant", control_type=control_type)
                    if pane.exists() and pane.is_visible():
                        self.accessibility_pane = pane
                        print(f"Found accessibility pane with control_type: {control_type}")
                        print("Elements in Accessibility Pane:")
                        print_elements(self.accessibility_pane)
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
            
            # Function to extract detailed accessibility data
            def extract_accessibility_categories(element, depth=0):
                category_data = {
                    'errors': [],
                    'warnings': [],
                    'tips': [],
                    'all_items': []
                }
                
                def explore_element(elem, current_depth=0):
                    try:
                        elem_text = elem.window_text().strip()
                        elem_type = elem.element_info().control_type
                        
                        print(f"{'  ' * current_depth}Exploring: [{elem_type}] '{elem_text}'")
                        
                        # Look for category headers or expandable sections
                        if elem_text:
                            lower_text = elem_text.lower()
                            
                            # Check if this is a category header
                            if any(word in lower_text for word in ['error', 'errors']):
                                print(f"{'  ' * current_depth}→ Found ERROR category")
                                # Get items under this category
                                category_items = get_category_items(elem, 'error')
                                category_data['errors'].extend(category_items)
                                
                            elif any(word in lower_text for word in ['warning', 'warnings']):
                                print(f"{'  ' * current_depth}→ Found WARNING category")
                                category_items = get_category_items(elem, 'warning')
                                category_data['warnings'].extend(category_items)
                                
                            elif any(word in lower_text for word in ['tip', 'tips', 'suggestion']):
                                print(f"{'  ' * current_depth}→ Found TIP category")
                                category_items = get_category_items(elem, 'tip')
                                category_data['tips'].extend(category_items)
                            
                            # Collect all non-empty text
                            if len(elem_text) > 2:  # Ignore very short text
                                category_data['all_items'].append({
                                    'text': elem_text,
                                    'type': elem_type,
                                    'depth': current_depth
                                })
                        
                        # Recursively explore children
                        try:
                            children = elem.children()
                            for child in children:
                                if child.is_visible():
                                    explore_element(child, current_depth + 1)
                        except:
                            pass
                            
                    except Exception as e:
                        print(f"{'  ' * current_depth}Error exploring element: {e}")
                
                def get_category_items(category_element, category_type):
                    """Extract items from a specific category"""
                    items = []
                    try:
                        # Try to expand the category if it's collapsible
                        try:
                            if hasattr(category_element, 'expand'):
                                category_element.expand()
                            # Or try clicking if it's clickable
                            elif hasattr(category_element, 'click_input'):
                                category_element.click_input()
                                time.sleep(0.5)  # Wait for expansion
                        except:
                            pass
                        
                        # Look for child elements that contain issue details
                        children = category_element.children()
                        for child in children:
                            try:
                                child_text = child.window_text().strip()
                                if child_text and len(child_text) > 5:
                                    items.append(child_text)
                                    
                                # Also check grandchildren
                                grandchildren = child.children()
                                for grandchild in grandchildren:
                                    try:
                                        gc_text = grandchild.window_text().strip()
                                        if gc_text and len(gc_text) > 5 and gc_text != child_text:
                                            items.append(gc_text)
                                    except:
                                        pass
                            except:
                                pass
                                
                    except Exception as e:
                        print(f"Error getting {category_type} items: {e}")
                    
                    return items
                
                # Start exploration
                explore_element(element)
                return category_data
            
            # Extract detailed category data
            print("Extracting detailed accessibility categories...")
            category_data = extract_accessibility_categories(self.accessibility_pane)
            
            # Update results with categorized data
            results['errors'] = category_data['errors']
            results['warnings'] = category_data['warnings'] 
            results['tips'] = category_data['tips']
            results['all_found_text'] = category_data['all_items']
            
            # Also try alternative approaches to get accessibility data
            print("\nTrying alternative extraction methods...")
            
            # Method 1: Look for specific UI automation patterns
            try:
                # Look for TreeView or List controls that might contain the issues
                tree_views = self.accessibility_pane.children(control_type="TreeView")
                list_views = self.accessibility_pane.children(control_type="ListView")
                lists = self.accessibility_pane.children(control_type="List")
                
                for control_list in [tree_views, list_views, lists]:
                    for control in control_list:
                        if control.is_visible():
                            print(f"Found {control.element_info.control_type}: {control.window_text()}")
                            # Get items from the control
                            try:
                                items = control.children()
                                for item in items:
                                    item_text = item.window_text().strip()
                                    if item_text:
                                        print(f"  Item: {item_text}")
                                        # Categorize the item
                                        lower_text = item_text.lower()
                                        if any(word in lower_text for word in ['error', 'critical']):
                                            results['errors'].append(item_text)
                                        elif any(word in lower_text for word in ['warning', 'caution']):
                                            results['warnings'].append(item_text)
                                        elif any(word in lower_text for word in ['tip', 'suggestion']):
                                            results['tips'].append(item_text)
                            except Exception as e:
                                print(f"Error getting items from {control.element_info.control_type}: {e}")
                                
            except Exception as e:
                print(f"Error with alternative extraction: {e}")
            
            # Method 2: Look for buttons or links that might represent issues
            try:
                buttons = self.accessibility_pane.children(control_type="Button")
                hyperlinks = self.accessibility_pane.children(control_type="Hyperlink")
                
                for button_list in [buttons, hyperlinks]:
                    for btn in button_list:
                        if btn.is_visible():
                            btn_text = btn.window_text().strip()
                            if btn_text and len(btn_text) > 3:
                                print(f"Found clickable item: {btn_text}")
                                
                                # Try clicking to get more details
                                try:
                                    btn.click_input()
                                    time.sleep(0.5)  # Wait for details to load
                                    
                                    # Look for detail panels or expanded content
                                    detail_text = self.get_expanded_details()
                                    if detail_text:
                                        full_issue = f"{btn_text}: {detail_text}"
                                    else:
                                        full_issue = btn_text
                                    
                                    # Categorize based on button text
                                    lower_text = btn_text.lower()
                                    if any(word in lower_text for word in ['error', 'critical']):
                                        results['errors'].append(full_issue)
                                    elif any(word in lower_text for word in ['warning', 'caution']):
                                        results['warnings'].append(full_issue)
                                    elif any(word in lower_text for word in ['tip', 'suggestion']):
                                        results['tips'].append(full_issue)
                                    
                                except Exception as e:
                                    print(f"Error clicking button {btn_text}: {e}")
                                    
            except Exception as e:
                print(f"Error with button extraction: {e}")
            
            # Method 3: Try to find text blocks or paragraphs
            try:
                text_blocks = self.accessibility_pane.children(control_type="Text")
                documents = self.accessibility_pane.children(control_type="Document")
                panes = self.accessibility_pane.children(control_type="Pane")
                
                for text_list in [text_blocks, documents, panes]:
                    for text_elem in text_list:
                        if text_elem.is_visible():
                            text_content = text_elem.window_text().strip()
                            if text_content and len(text_content) > 10:
                                print(f"Found text block: {text_content[:100]}...")
                                
                                # Split into lines and categorize
                                lines = text_content.split('\n')
                                for line in lines:
                                    line = line.strip()
                                    if line and len(line) > 5:
                                        lower_line = line.lower()
                                        if any(word in lower_line for word in ['error', 'critical']):
                                            results['errors'].append(line)
                                        elif any(word in lower_line for word in ['warning', 'caution']):
                                            results['warnings'].append(line)
                                        elif any(word in lower_line for word in ['tip', 'suggestion']):
                                            results['tips'].append(line)
                                            
            except Exception as e:
                print(f"Error with text block extraction: {e}")
            
            # Get all text content for combined analysis
            combined_text = ' '.join([
                item['text'] for item in results['all_found_text'] if isinstance(item, dict)
            ])
            
            # Add any direct text content
            if results['raw_text']:
                combined_text += ' ' + results['raw_text']
                        
        except Exception as e:
            print(f"Error with scrape accessibility: {e}")
               
    def get_expanded_details(self):
        """Get details from expanded accessibility issue"""
        try:
            # Look for detail panes or description areas that might have appeared
            detail_controls = self.accessibility_pane.children(control_type="Text")
            details = []
            
            for control in detail_controls:
                if control.is_visible():
                    text = control.window_text().strip()
                    if text and len(text) > 10:  # Filter out short/empty text
                        details.append(text)
            
            return " | ".join(details) if details else None
            
        except Exception as e:
            print(f"Error getting expanded details: {str(e)}")
            return None
            
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