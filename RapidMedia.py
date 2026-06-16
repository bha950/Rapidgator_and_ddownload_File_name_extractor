import pandas as pd
import requests
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
import os
import time

INPUT_FILE = "urls.xlsx"     # column A should contain URLs
OUTPUT_FILE = "results.xlsx"

MAX_WORKERS = 90             # increase/decrease depending on internet speed
SAVE_EVERY = 500             # autosave progress

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0 Safari/537.36"
    )
}


# ---------------------------------------------------
# Extract filename/title from HTML
# ---------------------------------------------------
def extract_title(url, html):
    soup = BeautifulSoup(html, "lxml")

    # Rapidgator
    if "rapidgator.net" in url:
        if soup.title:
            return soup.title.text.replace("Download file ", "").strip()

    # DDownload
    elif "ddownload.com" in url:
        h1 = soup.find("h1")
        if h1 and h1.text.strip():
            return h1.text.strip()

        if soup.title:
            return soup.title.text.replace("DDownload.com - ", "").strip()

        og = soup.find("meta", property="og:title")
        if og and og.get("content"):
            return og["content"].strip()

    # Generic fallback
    og = soup.find("meta", property="og:title")
    if og and og.get("content"):
        return og["content"].strip()

    if soup.title:
        return soup.title.text.strip()

    return "Not found"


# ---------------------------------------------------
# Fetch one URL
# ---------------------------------------------------
def process_url(index, url):
    try:
        response = requests.get(
            url,
            headers=HEADERS,
            timeout=20,
            allow_redirects=True
        )

        html = response.text
        result = extract_title(url, html)

        return index, result

    except Exception as e:
        return index, f"Error: {str(e)}"


# ---------------------------------------------------
# Main
# ---------------------------------------------------
def main():

    # Load Excel
    df = pd.read_excel(INPUT_FILE)

    # Assume first column contains URLs
    url_column = df.columns[0]

    # Create result column if not exists
    if "Filename" not in df.columns:
        df["Filename"] = ""

    # Resume support
    pending = []

    for idx, row in df.iterrows():
        if pd.isna(row["Filename"]) or row["Filename"] == "":
            pending.append((idx, str(row[url_column]).strip()))

    total = len(pending)

    print(f"Remaining URLs: {total}")

    completed = 0

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:

        futures = {
            executor.submit(process_url, idx, url): (idx, url)
            for idx, url in pending
            if url and url != "nan"
        }

        for future in as_completed(futures):

            idx, result = future.result()

            df.at[idx, "Filename"] = result

            completed += 1

            print(f"[{completed}/{total}] {result}")

            # Autosave
            if completed % SAVE_EVERY == 0:
                df.to_excel(OUTPUT_FILE, index=False)
                print(f"Autosaved at {completed}")

    # Final save
    df.to_excel(OUTPUT_FILE, index=False)

    print("Done.")


if __name__ == "__main__":
    start = time.time()

    main()

    end = time.time()

    print(f"Finished in {(end - start)/60:.2f} minutes")