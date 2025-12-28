# LinkedIn Job Extractor

Python scripts to extract job details from LinkedIn job postings and organize them in Excel spreadsheets and text files.

## Scripts

### 1. `linkedin_scraper.py`
Extracts job details from LinkedIn URLs in an Excel file and fills in missing information.

**Features:**
- Reads LinkedIn job URLs from Excel (supports both plain URLs and hyperlinked cells)
- Extracts: Company name, Job title, Days since posted
- Adds clickable hyperlinks to job titles (blue underlined style)
- Saves progress after each extraction (crash-resistant)
- Skips rows that already have complete data

**Usage:**
```bash
python linkedin_scraper.py
```

### 2. `job_descriptions_extractor.py`
Extracts full job descriptions from LinkedIn and organizes them by category into text files.

**Features:**
- Groups jobs by category (Digital, Analog, etc.)
- Extracts complete job descriptions with bullet points preserved
- Creates separate text files for each category (e.g., `Digital_jobs.txt`)
- Logs errors for non-LinkedIn URLs or failed extractions
- Handles "Show more" button to get full descriptions

**Usage:**
```bash
python job_descriptions_extractor.py
```

**Output format:**
```
Company: <company name>
Job Title: <job title>
URL: <linkedin url>

<full job description with bullet points>

------------------------------------------------------------
```

## Requirements

- Python 3.x
- Chrome browser installed

### Python packages:
```bash
pip install openpyxl selenium
```

## Excel File Structure

The scripts expect an Excel file named `Job Tracker.xlsx` with the following columns:
- **Column A**: Category (e.g., Digital, Analog)
- **Column B**: Company
- **Column C**: Job Title (with hyperlinks to LinkedIn job URLs)
- **Column D**: How long ago (Days)
- **Column E**: Notes

## Notes

- LinkedIn may require you to be logged in for some job pages
- Random delays are included between requests to avoid rate limiting
- Close the Excel file before running scripts to allow saving
