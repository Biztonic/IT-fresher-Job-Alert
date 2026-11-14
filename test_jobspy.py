import sys
from jobspy import scrape_jobs

try:
    jobs = scrape_jobs(
        site_name=["indeed", "linkedin"],
        search_term="QA Automation Engineer",
        location="Pune, India",
        results_wanted=3,
        hours_old=72,
        country_indeed="India",
    )
    print(f"Success! Scraped {len(jobs)} jobs:")
    for idx, row in jobs.iterrows():
        title = row.get('title', 'N/A')
        company = row.get('company', 'N/A')
        loc = row.get('location', 'N/A')
        url = row.get('job_url', 'N/A')
        print(f"- {title} at {company} ({loc}) -> {url[:60]}...")
except Exception as e:
    print(f"Error scraping: {e}")
