# google_contact_scraper_app.py

import streamlit as st
import pandas as pd
import tempfile
import os
from playwright.sync_api import sync_playwright
import openpyxl
import time
import re

# --- SCRAPE GOOGLE PAGE ---
def scrape_google_for_contact(firm_name, page):
    print(f"Searching Google for: {firm_name}")

    query = f"{firm_name} email phone site:.co.uk OR site:.com"
    page.goto(f"https://www.google.com/search?q={query}")

    page.wait_for_timeout(3000)

    try:
        # Loop through top 10 results
        for i in range(10):
            result_link = page.locator("div.MjjYud div > a").nth(i)
            link_href = result_link.get_attribute("href")

            if not link_href:
                continue

            print(f"Checking result: {link_href}")

            skip_keywords = ["facebook", "yell", "trustpilot", "linkedin", "192.com", "unbiased"]
            if any(keyword in link_href for keyword in skip_keywords):
                print(f"Skipping directory result: {link_href}")
                continue

            # Go to result
            page.goto(link_href)
            page.wait_for_timeout(3000)

            # Try clicking Contact link if present
            try:
                contact_link = page.locator("a:has-text('Contact')").first
                if contact_link.count() > 0:
                    contact_link.click()
                    page.wait_for_timeout(3000)
            except Exception as e:
                print(f"Contact link not found or click error: {e}")

            # Extract visible mailto links
            email_elements = page.locator("a[href^='mailto:']")
            emails = email_elements.all_inner_texts()
            email = emails[0] if emails else ""
            print(f"Email found (Google): {email}")

            # Extract visible tel links
            phone_elements = page.locator("a[href^='tel:']")
            phones = phone_elements.all_inner_texts()
            phone = phones[0] if phones else ""
            print(f"Phone found (Google): {phone}")

            # If we found either email or phone, return it
            if email or phone:
                return phone, email

            # If no Contact link or no data found, try About page
            try:
                about_link = page.locator("a:has-text('About')").first
                if about_link.count() > 0:
                    about_link.click()
                    page.wait_for_timeout(3000)

                    # Retry extracting mailto and tel on About page
                    email_elements = page.locator("a[href^='mailto:']")
                    emails = email_elements.all_inner_texts()
                    email = emails[0] if emails else email

                    phone_elements = page.locator("a[href^='tel:']")
                    phones = phone_elements.all_inner_texts()
                    phone = phones[0] if phones else phone

                    print(f"Email found (About page): {email}")
                    print(f"Phone found (About page): {phone}")

                    if email or phone:
                        return phone, email

            except Exception as e:
                print(f"About link not found or click error: {e}")

            # Also check footer
            try:
                footer = page.locator("footer").first
                footer_html = footer.inner_html()

                email_matches = re.findall(r'mailto:([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)', footer_html)
                email = email_matches[0] if email_matches else email

                phone_matches = re.findall(r'(?:\+?\d{1,3}[\s.-]?)?(?:\(?\d{2,4}\)?[\s.-]?)?\d{3,4}[\s.-]?\d{3,4}', footer_html)
                phone_matches = [p for p in phone_matches if len(p.strip()) >= 9]
                phone = phone_matches[0] if phone_matches else phone

                print(f"Email found (Footer): {email}")
                print(f"Phone found (Footer): {phone}")

                if email or phone:
                    return phone, email

            except Exception as e:
                print(f"Footer not found or scrape error: {e}")

        print("No suitable site found in top 10 Google results.")
        return "", ""

    except Exception as e:
        print(f"Google fallback scrape error: {e}")
        return "", ""

# --- MAIN APP ---
def main():
    st.title("Firm Email & Phone Scraper (Google Search)")

    uploaded_file = st.file_uploader("Upload Firm List Excel file", type=["xlsx"])

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, engine="openpyxl")

        df["Email"] = ""
        df["Phone"] = ""

        if st.button("Generate"):
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                page = browser.new_page()

                progress_bar = st.progress(0)
                status_text = st.empty()
                progress_text = st.empty()

                start_time = time.time()

                for idx, firm_name in enumerate(df.iloc[:, 0]):
                    print(f"\n--- Firm {idx+1}: {firm_name} ---")

                    phone, email = scrape_google_for_contact(firm_name, page)

                    df.at[idx, "Email"] = email
                    df.at[idx, "Phone"] = phone

                    time.sleep(2)

                    progress = (idx + 1) / len(df)
                    progress_bar.progress(progress)

                    elapsed = time.time() - start_time
                    avg_time_per_firm = elapsed / (idx + 1)
                    remaining_firms = len(df) - (idx + 1)
                    eta_seconds = remaining_firms * avg_time_per_firm

                    def format_seconds(seconds):
                        mins, secs = divmod(int(seconds), 60)
                        return f"{mins:02}:{secs:02}"

                    status_text.text(
                        f"Processing {idx + 1} of {len(df)} firms...\n"
                        f"Elapsed: {format_seconds(elapsed)}, "
                        f"ETA: {format_seconds(eta_seconds)}"
                    )

                    progress_text.text(f"Elapsed time: {format_seconds(elapsed)} seconds")

                browser.close()

                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    df.to_excel(tmp.name, index=False)
                    tmp_path = tmp.name

                st.success("Done! Download your enriched file below.")
                with open(tmp_path, "rb") as f:
                    st.download_button("Download Enriched Excel", f, file_name="enriched_firm_contacts.xlsx")

                os.unlink(tmp_path)

if __name__ == "__main__":
    main()
