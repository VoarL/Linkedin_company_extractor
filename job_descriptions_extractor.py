"""
Job Descriptions Extractor

This script reads the Job Tracker Excel file, opens each LinkedIn job link,
extracts the full job description, and writes them to text files organized by category.

Each category (Digital, Analog, etc.) gets its own text file with format:
    Company: <company name>
    Job Title: <job title>
    URL: <link>

    <job description>

    ----------------------------------------
"""

import re
import time
import random
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from datetime import datetime


def setup_driver():
    """Set up Chrome driver with options to avoid detection."""
    chrome_options = Options()
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    driver = webdriver.Chrome(options=chrome_options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver


def is_linkedin_job_url(url):
    """Check if the URL is a valid LinkedIn job URL."""
    if not url or not isinstance(url, str):
        return False

    linkedin_patterns = [
        r'linkedin\.com/jobs/view/\d+',
        r'linkedin\.com/jobs/search.*currentJobId=\d+',
    ]

    return any(re.search(pattern, url) for pattern in linkedin_patterns)


def get_url_from_cell(cell):
    """Extract URL from cell - either from hyperlink or cell value."""
    if cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    if cell.value and isinstance(cell.value, str):
        return cell.value
    return None


def format_description_from_html(html):
    """Convert HTML to plain text while preserving bullet points."""
    import re

    # Replace <li> tags with bullet points
    html = re.sub(r'<li[^>]*>', '\n  • ', html)
    html = re.sub(r'</li>', '', html)

    # Replace <br> tags with newlines
    html = re.sub(r'<br\s*/?>', '\n', html)

    # Replace paragraph/div closings with double newlines
    html = re.sub(r'</p>', '\n\n', html)
    html = re.sub(r'</div>', '\n', html)

    # Replace headers with newlines
    html = re.sub(r'</(h[1-6])>', '\n\n', html)

    # Remove all remaining HTML tags
    html = re.sub(r'<[^>]+>', '', html)

    # Decode HTML entities
    html = html.replace('&nbsp;', ' ')
    html = html.replace('&amp;', '&')
    html = html.replace('&lt;', '<')
    html = html.replace('&gt;', '>')
    html = html.replace('&quot;', '"')
    html = html.replace('&#39;', "'")

    # Clean up excessive whitespace
    lines = html.split('\n')
    cleaned_lines = []
    prev_empty = False

    for line in lines:
        line = line.strip()
        if line:
            cleaned_lines.append(line)
            prev_empty = False
        elif not prev_empty:
            cleaned_lines.append('')
            prev_empty = True

    return '\n'.join(cleaned_lines).strip()


def extract_job_info(driver, url):
    """Extract job title, company, and description from a LinkedIn job page."""
    try:
        driver.get(url)
        time.sleep(random.uniform(2, 4))

        company = None
        job_title = None
        description = None

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )

        # Extract job title
        for selector in ['h1.top-card-layout__title', 'h1.topcard__title', 'h1']:
            try:
                element = driver.find_element(By.CSS_SELECTOR, selector)
                if element and element.text.strip():
                    job_title = element.text.strip()
                    break
            except NoSuchElementException:
                continue

        # Extract company name
        for selector in ['a.topcard__org-name-link', '.topcard__flavor a', 'a[href*="/company/"]']:
            try:
                element = driver.find_element(By.CSS_SELECTOR, selector)
                if element and element.text.strip():
                    company = element.text.strip()
                    break
            except NoSuchElementException:
                continue

        # Try to click "Show more" button to expand description
        try:
            show_more_btn = driver.find_element(By.CSS_SELECTOR, '.show-more-less-html__button--more')
            show_more_btn.click()
            time.sleep(0.5)
        except:
            pass  # Button may not exist or already expanded

        # Extract job description with bullet points preserved
        description_selectors = [
            '.show-more-less-html__markup',
            '.description__text',
            '.jobs-description__content',
            '.jobs-box__html-content',
            'div[class*="description"]',
            '.job-details',
        ]

        for selector in description_selectors:
            try:
                element = driver.find_element(By.CSS_SELECTOR, selector)
                if element:
                    # Get inner HTML and convert to text with bullets preserved
                    html = element.get_attribute('innerHTML')
                    if html:
                        description = format_description_from_html(html)
                        if description:
                            break
            except NoSuchElementException:
                continue

        return {
            'company': company,
            'job_title': job_title,
            'description': description
        }

    except TimeoutException:
        print(f"    Timeout loading page")
        return None
    except Exception as e:
        print(f"    Error: {str(e)}")
        return None


def sanitize_filename(name):
    """Remove invalid characters from filename."""
    return re.sub(r'[<>:"/\\|?*]', '', name)


def main():
    excel_path = "Job Tracker.xlsx"

    print(f"Reading {excel_path}...")
    wb = load_workbook(excel_path)
    ws = wb.active

    # Find column indices
    headers = {}
    for col in range(1, ws.max_column + 1):
        header = ws.cell(row=1, column=col).value
        if header:
            headers[header] = col

    print(f"Headers found: {headers}")

    category_col = headers.get('Category', 1)
    company_col = headers.get('Company', 2)
    job_title_col = headers.get('Job Title', 3)  # This column has hyperlinks to jobs

    # Organize jobs by category
    jobs_by_category = {}
    errors = []

    for row in range(2, ws.max_row + 1):
        category = ws.cell(row=row, column=category_col).value
        if not category:
            continue

        category = category.strip()
        if category not in jobs_by_category:
            jobs_by_category[category] = []

        # Get URL from Job Title column hyperlink
        job_title_cell = ws.cell(row=row, column=job_title_col)
        url = get_url_from_cell(job_title_cell)

        # Get existing company/title from spreadsheet as fallback
        existing_company = ws.cell(row=row, column=company_col).value
        existing_title = job_title_cell.value

        jobs_by_category[category].append({
            'row': row,
            'url': url,
            'existing_company': existing_company,
            'existing_title': existing_title,
            'is_linkedin': is_linkedin_job_url(url) if url else False
        })

    # Count jobs
    total_jobs = sum(len(jobs) for jobs in jobs_by_category.values())
    linkedin_jobs = sum(1 for jobs in jobs_by_category.values() for job in jobs if job['is_linkedin'])

    print(f"\nFound {total_jobs} jobs across {len(jobs_by_category)} categories")
    print(f"LinkedIn jobs: {linkedin_jobs}")
    print(f"Non-LinkedIn/No URL: {total_jobs - linkedin_jobs}")
    print(f"\nCategories: {list(jobs_by_category.keys())}")

    # Set up browser
    print("\nSetting up browser...")
    driver = setup_driver()

    try:
        processed = 0
        for category, jobs in jobs_by_category.items():
            print(f"\n{'='*60}")
            print(f"Processing category: {category} ({len(jobs)} jobs)")
            print('='*60)

            filename = f"{sanitize_filename(category)}_jobs.txt"

            with open(filename, 'w', encoding='utf-8') as f:
                f.write(f"{'='*60}\n")
                f.write(f"Category: {category}\n")
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"{'='*60}\n\n")

                for i, job in enumerate(jobs):
                    processed += 1
                    row = job['row']
                    url = job['url']

                    print(f"\n[{processed}/{total_jobs}] Row {row}...")

                    if not job['is_linkedin']:
                        # Non-LinkedIn or no URL
                        error_msg = f"Row {row}: Not a LinkedIn URL"
                        if url:
                            error_msg += f" - {url[:50]}..."
                        else:
                            error_msg += " - No URL"
                        errors.append(error_msg)
                        print(f"  Skipping: {error_msg}")

                        # Still write entry with available info
                        f.write(f"Company: {job['existing_company'] or 'N/A'}\n")
                        f.write(f"Job Title: {job['existing_title'] or 'N/A'}\n")
                        f.write(f"URL: {url or 'N/A'}\n")
                        f.write(f"Status: SKIPPED - Not a LinkedIn URL\n")
                        f.write(f"\n{'-'*40}\n\n")
                        continue

                    print(f"  URL: {url[:50]}...")
                    info = extract_job_info(driver, url)

                    if info:
                        company = info['company'] or job['existing_company'] or 'N/A'
                        title = info['job_title'] or job['existing_title'] or 'N/A'
                        description = info['description'] or 'No description available'

                        print(f"  Company: {company}")
                        print(f"  Title: {title}")
                        print(f"  Description: {len(description)} chars")

                        f.write(f"Company: {company}\n")
                        f.write(f"Job Title: {title}\n")
                        f.write(f"URL: {url}\n")
                        f.write(f"\n{description}\n")
                        f.write(f"\n{'-'*40}\n\n")
                    else:
                        error_msg = f"Row {row}: Failed to extract - {url[:50]}..."
                        errors.append(error_msg)
                        print(f"  ERROR: Failed to extract")

                        f.write(f"Company: {job['existing_company'] or 'N/A'}\n")
                        f.write(f"Job Title: {job['existing_title'] or 'N/A'}\n")
                        f.write(f"URL: {url}\n")
                        f.write(f"Status: ERROR - Failed to extract description\n")
                        f.write(f"\n{'-'*40}\n\n")

                    # Random delay
                    time.sleep(random.uniform(2, 4))

            print(f"\nSaved: {filename}")

    finally:
        driver.quit()

    # Write error log
    if errors:
        with open('extraction_errors.txt', 'w', encoding='utf-8') as f:
            f.write(f"Extraction Errors - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("="*60 + "\n\n")
            for error in errors:
                f.write(f"{error}\n")
        print(f"\n{len(errors)} errors logged to extraction_errors.txt")

    print(f"\nDone! Processed {processed} jobs.")
    print(f"Output files created for categories: {list(jobs_by_category.keys())}")


if __name__ == "__main__":
    main()
