import warnings

warnings.filterwarnings("ignore")

import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import time


def extract_info_from_homepage(url):
    headers = {"User-Agent": "Mozilla/5.0"}
    email, address, ceo = "", "", ""
    try:
        resp = requests.get(url, headers=headers, timeout=10)
        if not resp.ok:
            return address, email, ceo
        soup = BeautifulSoup(resp.text, "lxml")
        text = soup.get_text(separator="\n")

        # Find all emails
        emails = re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.\w+", text)
        email = emails[0] if emails else ""

        # Try to find CEO pattern
        ceo_patterns = [
            r"CEO[:\-–\s]+([A-Z][A-Za-z.\s]+)",
            r"Chief Executive Officer[:\-–\s]+([A-Z][A-Za-z.\s]+)",
        ]
        for pat in ceo_patterns:
            ceo_match = re.search(pat, text, re.IGNORECASE)
            if ceo_match:
                ceo = ceo_match.group(1).strip()
                break

        # Try to find an address (looks for lines containing "address", "headquarters", US address format, etc.)
        lines = text.splitlines()
        for line in lines:
            if len(address) == 0:
                if "address" in line.lower() or "headquarters" in line.lower():
                    address = line.strip()
            # Try to catch US-style addresses
            addr_match = re.search(r"\d{1,5}\s+\w.*,\s*[A-Za-z\s]+,\s*[A-Z]{2}\s+\d{5}", line)
            if addr_match and len(address) == 0:
                address = addr_match.group().strip()
    except Exception as e:
        pass  # Could log errors here if you want
    return address, email, ceo


# Load data
df = pd.read_csv("fortune1000_2024.csv")
results = []

for idx, row in df.iterrows():
    company = row.get('Company')
    website = row.get('Website')
    if not isinstance(website, str) or not website.startswith(('http', 'www')):
        continue
    if not website.startswith("http"):
        url = "http://" + website
    else:
        url = website
    print(f"Scraping {company} ({url}) ...")
    address, email, ceo = extract_info_from_homepage(url)
    results.append({
        "Company": company,
        "Address": address,
        "Email-ID": email,
        "CEO": ceo
    })
    time.sleep(2)  # avoid overwhelming servers and getting blocked

# Export to Excel
outdf = pd.DataFrame(results)
outdf.to_excel("scraped_company_websites.xlsx", index=False)
print("DONE! Data saved to scraped_company_websites.xlsx")
