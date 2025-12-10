import pandas as pd
import time
import re
from datetime import datetime
from typing import Callable, Optional, Dict, List, Any
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager


# ================= DEFAULT CONFIG =================
DEFAULT_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLScuV1zxbgFNmedJjKqrWdmsN_f7O7H2ij439HlfVy_GJXJHYA/viewform?usp=header"
DEFAULT_EXCEL_PATH = "EOI_Generation_Panipat_100_Records.xlsx"
# ==================================================


class GoogleFormSubmitter:
    """
    Class to handle automated Google Form submissions from Excel data.
    Can be used standalone or integrated with FastAPI.
    """
    
    def __init__(self, form_url: str, excel_path: str, headless: bool = False):
        """
        Initialize the Google Form Submitter.
        
        Args:
            form_url: The Google Form URL
            excel_path: Path to the Excel file with form data
            headless: Whether to run browser in headless mode (default: False)
        """
        self.form_url = form_url
        self.excel_path = excel_path
        self.headless = headless
        self.driver = None
        self.form_map = {}
        self.errors: List[str] = []
    
    def _create_driver(self) -> webdriver.Chrome:
        """Create and configure Chrome WebDriver."""
        import os
        options = Options()
        
        # Always use headless mode in production/Docker
        if self.headless or os.environ.get('DOCKER_ENV'):
            options.add_argument("--headless=new")
        
        # Essential for Docker/containerized environments
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-software-rasterizer")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-infobars")
        options.add_argument("--disable-setuid-sandbox")
        options.add_argument("--single-process")
        options.add_argument("--ignore-certificate-errors")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-features=VizDisplayCompositor")
        
        # Add user agent to appear more like a real browser
        options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        
        # For Docker/Linux environments
        if os.environ.get('CHROME_BIN'):
            options.binary_location = os.environ.get('CHROME_BIN')

        return webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options
        )
    
    @staticmethod
    def normalize(text: Any) -> str:
        """
        Normalize text by removing whitespace, newlines, asterisks, and converting to lowercase.
        Google Forms adds newlines and asterisks (*) to required field labels.
        """
        if not isinstance(text, str):
            text = str(text)
        # Remove newlines, asterisks, extra spaces, then strip and lowercase
        text = text.replace('\n', '').replace('*', '').strip().lower()
        # Remove multiple spaces
        text = re.sub(r'\s+', ' ', text)
        return text
    
    def _extract_form_mapping(self) -> Dict[str, Dict]:
        """
        Dynamically builds mapping using the question's container element.
        Returns: { normalized_question_text : question_element_info }
        """
        mapping = {}
        wait = WebDriverWait(self.driver, 30)  # Increased timeout
        
        # First, check if the page loaded at all
        try:
            # Wait for body to be present
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        except TimeoutException:
            print("❌ Page failed to load at all")
            raise Exception("Page failed to load")
        
        # Take a screenshot for debugging
        self.driver.save_screenshot("debug_page_load.png")
        print(f"📸 Screenshot saved: debug_page_load.png")
        print(f"📄 Current URL: {self.driver.current_url}")
        print(f"📄 Page title: {self.driver.title}")
        
        # Check for common Google Forms issues
        page_source = self.driver.page_source.lower()
        
        if "sign in" in page_source or "login" in page_source:
            print("⚠️ Google Forms may require sign-in")
        
        if "sorry" in page_source and "access" in page_source:
            raise Exception("Form access denied - the form may be restricted")
        
        # Try multiple selectors for form questions
        selectors = [
            "//div[@role='listitem']",
            "//div[contains(@class, 'freebirdFormviewerComponentsQuestionBaseRoot')]",
            "//div[contains(@class, 'freebirdFormviewerViewNumberedItemContainer')]",
            "//div[@data-params]"
        ]
        
        questions = []
        for selector in selectors:
            try:
                # Wait a bit for each selector
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, selector))
                )
                questions = self.driver.find_elements(By.XPATH, selector)
                if questions:
                    print(f"✅ Found {len(questions)} questions using selector: {selector}")
                    break
            except TimeoutException:
                print(f"  Selector '{selector}' not found, trying next...")
                continue
        
        if not questions:
            # Last resort: print page structure for debugging
            print("\n❌ Could not find form questions. Page structure:")
            try:
                forms = self.driver.find_elements(By.TAG_NAME, "form")
                print(f"  Found {len(forms)} form elements")
                divs = self.driver.find_elements(By.XPATH, "//div[@role]")
                roles = set([d.get_attribute('role') for d in divs[:20]])
                print(f"  Found div roles: {roles}")
            except Exception as e:
                print(f"  Error inspecting page: {e}")
            
            raise Exception("No form questions found. The form may not have loaded correctly.")
        
        print(f"Found {len(questions)} form questions")

        for idx, q in enumerate(questions):
            try:
                # Get the question label
                label_elem = q.find_element(By.XPATH, ".//div[@role='heading']")
                label = label_elem.text
                
                # Try to find input field (any type)
                try:
                    input_box = q.find_element(By.XPATH, ".//input | .//textarea")
                    field_type = input_box.get_attribute("type") or input_box.tag_name
                    
                    # Store the xpath to the question container
                    xpath = f"(//div[@role='listitem'])[{idx + 1}]//input | (//div[@role='listitem'])[{idx + 1}]//textarea"
                    mapping[self.normalize(label)] = {
                        'xpath': xpath,
                        'type': field_type,
                        'label': label,
                        'index': idx + 1
                    }
                    print(f"  Mapped: '{label}' -> Question #{idx+1} (type: {field_type})")
                except Exception:
                    # Other field types (dropdown, radio, etc.)
                    continue
                        
            except Exception:
                continue

        return mapping
    
    @staticmethod
    def _parse_date_value(value: Any) -> tuple:
        """
        Parse date value from Excel (Timestamp, datetime, or string) and return day, month, year.
        """
        day, month, year = None, None, None
        
        try:
            # Handle pandas Timestamp
            if isinstance(value, pd.Timestamp):
                day = value.day
                month = value.month
                year = value.year
            # Handle datetime objects
            elif hasattr(value, 'day') and hasattr(value, 'month') and hasattr(value, 'year'):
                day = value.day
                month = value.month
                year = value.year
            # Handle string dates
            elif isinstance(value, str):
                value = value.strip()
                # Try to parse common date formats
                for fmt in ['%d-%m-%Y', '%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%d-%m-%y', '%Y/%m/%d']:
                    try:
                        dt = datetime.strptime(value, fmt)
                        day = dt.day
                        month = dt.month
                        year = dt.year
                        break
                    except ValueError:
                        continue
            # Handle numeric (Excel serial date)
            elif isinstance(value, (int, float)):
                # Excel serial date conversion
                dt = pd.to_datetime(value, unit='D', origin='1899-12-30')
                day = dt.day
                month = dt.month
                year = dt.year
        except Exception as e:
            print(f"    Warning: Could not parse date '{value}': {e}")
        
        return day, month, year
    
    def _fill_date_field(self, question_container_xpath: str, value: Any, max_retries: int = 3) -> bool:
        """
        Special handling for Google Forms date fields which use separate inputs for day, month, year.
        """
        day, month, year = self._parse_date_value(value)
        
        if day is None or month is None or year is None:
            print(f"    ❌ Could not parse date: {value}")
            return False
        
        print(f"    Parsed date: Day={day}, Month={month}, Year={year}")
        
        # Extract the question index from xpath
        match = re.search(r'\[(\d+)\]', question_container_xpath)
        if not match:
            print("    ❌ Could not extract question index from xpath")
            return False
        
        q_index = match.group(1)
        base_xpath = f"(//div[@role='listitem'])[{q_index}]"
        
        for attempt in range(max_retries):
            try:
                # Try to find date input fields by their placeholder or aria-label
                date_inputs = self.driver.find_elements(
                    By.XPATH, f"{base_xpath}//input[@type='text' or @type='tel']"
                )
                
                if len(date_inputs) >= 3:
                    # Google Forms typically has Day, Month, Year inputs
                    for inp in date_inputs:
                        aria_label = (inp.get_attribute('aria-label') or '').lower()
                        placeholder = (inp.get_attribute('placeholder') or '').lower()
                        label_text = aria_label + placeholder
                        
                        self.driver.execute_script(
                            "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", inp
                        )
                        time.sleep(0.2)
                        
                        inp.click()
                        time.sleep(0.1)
                        inp.clear()
                        
                        if 'day' in label_text or 'dd' in label_text:
                            inp.send_keys(str(day).zfill(2))
                            print(f"    Entered day: {day}")
                        elif 'month' in label_text or 'mm' in label_text:
                            inp.send_keys(str(month).zfill(2))
                            print(f"    Entered month: {month}")
                        elif 'year' in label_text or 'yyyy' in label_text:
                            inp.send_keys(str(year))
                            print(f"    Entered year: {year}")
                        
                        time.sleep(0.2)
                    
                    return True
                
                # Alternative: If there's a single date input field (HTML5 date type)
                date_input = self.driver.find_elements(By.XPATH, f"{base_xpath}//input[@type='date']")
                if date_input:
                    field = date_input[0]
                    self.driver.execute_script(
                        "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", field
                    )
                    time.sleep(0.3)
                    
                    # HTML5 date input expects YYYY-MM-DD format
                    date_str = f"{year}-{str(month).zfill(2)}-{str(day).zfill(2)}"
                    
                    # Use JavaScript to set the value
                    self.driver.execute_script("""
                        arguments[0].value = arguments[1];
                        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                    """, field, date_str)
                    
                    time.sleep(0.3)
                    print(f"    Date entered via HTML5 input: {date_str}")
                    return True
                
                # Fallback: Try typing the date in DD-MM-YYYY format
                text_inputs = self.driver.find_elements(By.XPATH, f"{base_xpath}//input")
                if text_inputs:
                    field = text_inputs[0]
                    self.driver.execute_script(
                        "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", field
                    )
                    time.sleep(0.3)
                    field.click()
                    time.sleep(0.2)
                    field.clear()
                    
                    # Try typing date as DD-MM-YYYY
                    date_str = f"{str(day).zfill(2)}-{str(month).zfill(2)}-{year}"
                    field.send_keys(date_str)
                    field.send_keys(Keys.TAB)
                    time.sleep(0.3)
                    print(f"    Date typed: {date_str}")
                    return True
                    
            except Exception as e:
                print(f"    Attempt {attempt + 1} failed for date field: {str(e)[:100]}")
                time.sleep(0.5)
        
        return False
    
    def _fill_field_with_retry(self, field_info: Dict, value: Any, max_retries: int = 3) -> bool:
        """
        Fill a form field with retry logic and automatic field type detection.
        """
        wait = WebDriverWait(self.driver, 10)
        xpath = field_info['xpath']
        field_type = field_info['type']
        
        # Special handling for date fields
        if field_type == 'date':
            return self._fill_date_field(xpath, value, max_retries)
        
        for attempt in range(max_retries):
            try:
                # Wait for field to be present
                field = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                
                # Scroll to field
                self.driver.execute_script(
                    "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", field
                )
                time.sleep(0.3)
                
                # Click to focus
                field.click()
                time.sleep(0.3)
                
                # Clear existing value
                field.clear()
                time.sleep(0.2)
                
                # Send keys
                field.send_keys(str(value))
                time.sleep(0.3)
                
                # For email fields, trigger validation
                if field_type == 'email':
                    field.send_keys(Keys.TAB)
                    time.sleep(0.4)
                
                # Verify the value was entered
                entered_value = field.get_attribute('value')
                if entered_value:
                    return True
                else:
                    print(f"    Attempt {attempt + 1}: No value entered (expected '{value}')")
                    
            except Exception as e:
                print(f"    Attempt {attempt + 1} failed: {str(e)[:100]}")
                time.sleep(0.5)
        
        return False
    
    def _wait_for_submission_confirmation(self, timeout: int = 10) -> bool:
        """Wait for Google Forms confirmation page/message."""
        try:
            wait = WebDriverWait(self.driver, timeout)
            wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(text(), 'Your response has been recorded') or contains(text(), 'submitted')]")
                )
            )
            return True
        except TimeoutException:
            print("⚠️ WARNING: Submission confirmation not detected")
            return False
    
    def _submit_row(self, row: pd.Series, row_index: int) -> bool:
        """Submit a single row with proper error handling and confirmation."""
        try:
            self.driver.get(self.form_url)
            wait = WebDriverWait(self.driver, 15)
            
            # Wait for form to be interactive
            wait.until(EC.presence_of_element_located((By.XPATH, "//div[@role='listitem']")))
            time.sleep(1.5)

            # Fill in the form fields
            filled_count = 0
            failed_fields = []
            
            for col, value in row.items():
                key = self.normalize(col)
                if key in self.form_map:
                    field_info = self.form_map[key]
                    print(f"  Filling '{col}' ({field_info['type']}): {value}")
                    
                    if self._fill_field_with_retry(field_info, value):
                        filled_count += 1
                        print("    ✅ Successfully filled")
                    else:
                        print(f"    ❌ Failed to fill field '{col}'")
                        failed_fields.append(col)
                else:
                    print(f"  ⚠️ Field '{col}' not found in form mapping")

            total_fields = len([k for k in row.keys() if self.normalize(k) in self.form_map])
            print(f"  Filled {filled_count}/{total_fields} fields")
            
            if failed_fields:
                error_msg = f"Row {row_index + 1}: Failed fields - {', '.join(failed_fields)}"
                self.errors.append(error_msg)
                print(f"  ⚠️ {error_msg}")

            # Find and click submit button
            try:
                time.sleep(1)
                
                submit_btn = wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//div[@role='button' and @aria-label='Submit']")
                    )
                )
                
                self.driver.execute_script(
                    "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", submit_btn
                )
                time.sleep(0.5)
                
                self.driver.execute_script("arguments[0].click();", submit_btn)
                print("  Clicked submit button")
                
                if self._wait_for_submission_confirmation():
                    print("  ✅ Submission confirmed!")
                    return True
                else:
                    error_msg = f"Row {row_index + 1}: Submission confirmation not received"
                    self.errors.append(error_msg)
                    print(f"  ❌ {error_msg}")
                    return False
                    
            except TimeoutException:
                error_msg = f"Row {row_index + 1}: Submit button not found"
                self.errors.append(error_msg)
                print(f"  ❌ {error_msg}")
                return False

        except Exception as e:
            error_msg = f"Row {row_index + 1}: {str(e)}"
            self.errors.append(error_msg)
            print(f"  ❌ Error during submission: {e}")
            return False
    
    def run(self, progress_callback: Optional[Callable] = None) -> Dict[str, Any]:
        """
        Run the form submission process.
        
        Args:
            progress_callback: Optional callback function(current_row, total_rows, success, fail, message)
        
        Returns:
            Dictionary with results: {total_rows, success_count, fail_count, errors}
        """
        df = pd.read_excel(self.excel_path)
        total_rows = len(df)
        print(f"Loaded {total_rows} rows from Excel")
        print(f"Columns: {list(df.columns)}\n")
        
        self.driver = self._create_driver()
        self.errors = []
        
        try:
            print(f"🌐 Loading form URL: {self.form_url}")
            self.driver.get(self.form_url)
            
            # Wait longer for initial page load
            time.sleep(5)
            
            # Handle potential cookie consent or other popups
            try:
                # Try to dismiss any consent dialogs (common on Google services)
                consent_buttons = self.driver.find_elements(
                    By.XPATH, 
                    "//button[contains(text(), 'Accept') or contains(text(), 'I agree') or contains(text(), 'Continue')]"
                )
                for btn in consent_buttons:
                    try:
                        btn.click()
                        print("  Clicked consent/continue button")
                        time.sleep(2)
                    except Exception:
                        pass
            except Exception:
                pass

            self.form_map = self._extract_form_mapping()
            print(f"\n✅ Dynamic form mapping complete: {len(self.form_map)} fields mapped\n")

            if not self.form_map:
                raise Exception("No form fields detected! Check the form URL.")

            # Show which Excel columns will be used
            print("Excel columns that will be filled:")
            for col in df.columns:
                key = self.normalize(col)
                if key in self.form_map:
                    print(f"  ✅ '{col}' -> Form field")
                else:
                    print(f"  ⚠️ '{col}' -> No matching form field")
            print()

            success_count = 0
            fail_count = 0

            for index, row in df.iterrows():
                print(f"\n{'='*60}")
                print(f"📝 Processing row {index + 1}/{total_rows}...")
                print(f"{'='*60}")
                
                if self._submit_row(row, index):
                    success_count += 1
                    time.sleep(2)
                else:
                    fail_count += 1
                    time.sleep(1)
                
                # Call progress callback if provided
                if progress_callback:
                    progress_callback(
                        index + 1, total_rows, success_count, fail_count,
                        f"Processing row {index + 1}/{total_rows}"
                    )

            print(f"\n{'='*60}")
            print(f"✅ Successfully submitted: {success_count}")
            print(f"❌ Failed submissions: {fail_count}")
            print(f"{'='*60}")
            
            return {
                "total_rows": total_rows,
                "success_count": success_count,
                "fail_count": fail_count,
                "errors": self.errors
            }

        except Exception as e:
            print(f"❌ Fatal error: {e}")
            import traceback
            traceback.print_exc()
            raise
        finally:
            if self.driver:
                self.driver.quit()


def main():
    """Main function for standalone execution."""
    submitter = GoogleFormSubmitter(
        form_url=DEFAULT_FORM_URL,
        excel_path=DEFAULT_EXCEL_PATH,
        headless=False
    )
    result = submitter.run()
    print(f"\nFinal Result: {result}")


if __name__ == "__main__":
    main()
