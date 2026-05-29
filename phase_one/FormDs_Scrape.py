# formds code to collect data in current date 
# made by : Eslam Ayman 
# Date : 9 / 9 / 2025

import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
from datetime import datetime

company_names = []
company_fields = []
reported_funding_values = []
incremental_cash_values = []
filing_dates = []
new_or_amended_flags = []
filing_links = []

TARGET_DATE = datetime.today().strftime("%Y-%m-%d")
print(f" Collecting data for {TARGET_DATE}...")

stop_scraping = False
page = 1

while not stop_scraping:
    url = f"https://www.formds.com/filings/newest?page={page}"
    print(f"\n📄 Loading page {page} ...")

    try:
        response = requests.get(url)
        if response.status_code != 200:
            print(f" Failed to load page {page}, Status Code: {response.status_code}")
            break

        soup = BeautifulSoup(response.text, 'lxml')
        rows = soup.find_all("tr")

        if not rows or len(rows) == 1:
            print(" No more data found. Stopping...")
            break

        for row in rows[1:]:
            cols = row.find_all("td")
            if len(cols) < 6:
                continue

            filing_date = cols[4].text.strip()
            if filing_date > TARGET_DATE:
                continue
            if filing_date < TARGET_DATE:
                stop_scraping = True
                break

            company_tag = cols[0].find("a")
            company_name = company_tag.text.strip() if company_tag else cols[0].text.strip()
            company_link = "https://www.formds.com" + company_tag['href'] if company_tag and company_tag.has_attr("href") else "N/A"

            field_tag = cols[0].find("span")
            company_field = field_tag.text.strip("() ") if field_tag else ""

            reported_funding = cols[2].text.strip()
            incremental_cash = cols[3].text.strip()
            new_or_amended = cols[5].text.strip()

            company_names.append(company_name)
            company_fields.append(company_field)
            reported_funding_values.append(reported_funding)
            incremental_cash_values.append(incremental_cash)
            filing_dates.append(filing_date)
            new_or_amended_flags.append(new_or_amended)
            filing_links.append(company_link)

        page += 1
        time.sleep(1)

    except Exception as e:
        print(f" Error on page {page}: {e}")
        break

df = pd.DataFrame({
    "Company": company_names,
    "Field": company_fields,
    "Reported Funding": reported_funding_values,
    "Incremental Cash": incremental_cash_values,
    "Date": filing_dates,
    "New or Amended": new_or_amended_flags,
    "Link": filing_links
})

output_file = f"formds_{TARGET_DATE}.xlsx"
df.to_excel(output_file, index=False)

print(f"\n {len(df)} records saved to {output_file} successfully ")
