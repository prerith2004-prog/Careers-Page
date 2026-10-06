# Career Site Job Scraper

A Python script that reads a list of company career sites from an Excel file, opens each site in a browser, searches for jobs, keeps only **Bangalore** and **IT-related** roles, and saves the results to a new Excel file.

The script was written with AI assistance as part of a task to automate job collection from company career pages.

## How it works

1. **Input:** `career_sites.xlsx` holds the list of career site URLs (columns: `Company`, `Career URL`).
2. **Open and read:** The script opens each site in a headless Chromium browser using [Playwright](https://playwright.dev/python/). It accepts cookie pop-ups, types the location and keyword into the site's search boxes if present, scrolls, and reads up to 3 result pages. If the landing page shows no jobs, it follows "Search jobs" style links and tries again.
3. **Filter:** A job is kept only if all three are true:
   - the link looks like a job link,
   - the job card mentions Bangalore / Bengaluru,
   - the title or card contains an IT keyword (software, developer, java, python, devops, cloud, data, QA, SAP, security, etc.).
4. **Output:** `jobs_output.xlsx` is created with two sheets:
   - **Jobs:** Company, Job Title, Location, Posted, Job Link, Source Page, Fetched On
   - **Run Log:** status and job count for each company, so you can see which sites worked

A screenshot of each site is saved in the `screenshots/` folder for debugging.

## Project structure

```
.
├── career_sites.xlsx    # input: list of career sites
├── job_scraper.py       # the scraper
├── jobs_output.xlsx     # output (created when you run the script)
├── screenshots/         # one screenshot per site (created when you run the script)
└── README.md
```

## Requirements

- Python 3.9 or newer
- Windows, macOS or Linux

## Setup

```bash
pip install playwright pandas openpyxl
playwright install chromium
```

## Usage

1. Put your career site URLs in `career_sites.xlsx`, one company per row.
2. Run the script from the folder containing both files:

   ```bash
   python job_scraper.py
   ```

3. Open `jobs_output.xlsx` when it finishes.

## Settings

Edit these values at the top of `job_scraper.py`:

| Setting | Default | Meaning |
|---|---|---|
| `SEARCH_LOCATION` | `Bangalore` | Typed into the site's location box |
| `SEARCH_KEYWORD` | `software` | Typed into the site's keyword box |
| `MAX_PAGES` | `3` | Result pages to read per site |
| `HEADLESS` | `True` | Set to `False` to watch the browser while it runs |
| `LOCATION_KEYWORDS` | Bangalore, Bengaluru, Bengalooru | Location words a job must contain |
| `IT_KEYWORDS` | software, developer, engineer, ... | IT words a job must contain |

## Author

[Prerith]
