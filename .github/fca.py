# fca_streamlit_app.py

import streamlit as st
import pandas as pd
import tempfile
import os
from playwright.sync_api import sync_playwright
import openpyxl
import time

# --- SCRAPE FUNCTION ---
def scrape_fca_page(url, page):
    try:
        page.goto(url, wait_until="domcontentloaded")

        # Accept cookies popup if shown
        try:
            if page.locator("text=Accept all cookies").is_visible():
                page.locator("text=Accept all cookies").click()
                page.wait_for_timeout(1000)
        except:
            pass

        # --- Scrape Address ---
        try:
            address_section = page.locator("text=Address").locator("xpath=..").locator("xpath=following-sibling::*[1]")
            address_text = address_section.inner_text().strip()
            address_lines = address_text.split("\n")
        except:
            address_lines = []

        # --- Scrape Phone ---
        try:
            phone_section = page.locator("text=Phone").locator("xpath=..").locator("xpath=following-sibling::*[1]")
            phone = phone_section.inner_text().strip()
        except:
            phone = ""

        # --- Scrape Email ---
        try:
            email_section = page.locator("text=Email").locator("xpath=..").locator("xpath=following-sibling::*[1]")
            email = email_section.inner_text().strip()
        except:
            email = ""

        return address_lines, phone, email

    except Exception as e:
        print(f"Error scraping {url}: {e}")
        return [], "", ""

# --- MAIN APP ---
def main():
    st.title("FCA Register Data Enrichment Tool")

    uploaded_file = st.file_uploader("Upload FCA Register Excel file", type=["xlsx"])

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, engine="openpyxl")

        # Extract hyperlinks from column A
        wb = openpyxl.load_workbook(uploaded_file)
        sheet = wb.active
        urls = []
        for cell in sheet["A"][1:]:  # skip header row
            if cell.hyperlink:
                urls.append(cell.hyperlink.target)
            else:
                urls.append("")

        df["FCA URL"] = urls

        if st.button("Generate"):
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                page = browser.new_page()  # Reuse single page → faster

                # Prepare columns BEFORE the loop
                df["Address Line 1"] = ""
                df["Address Line 2"] = ""
                df["Address Line 3"] = ""
                df["Address Line 4"] = ""
                df["Postcode"] = ""
                df["Email"] = ""
                df["Phone"] = ""

                # Progress bar + Timer
                progress_bar = st.progress(0)
                status_text = st.empty()
                start_time = time.time()

                for idx, url in enumerate(df["FCA URL"]):
                    if url:
                        address_lines, phone, email = scrape_fca_page(url, page)

                        # Fill address columns
                        for i in range(5):
                            column_name = f"Address Line {i+1}" if i < 4 else "Postcode"
                            value = address_lines[i] if i < len(address_lines) else ""
                            df.iloc[idx, df.columns.get_loc(column_name)] = value

                        df.at[idx, "Email"] = email
                        df.at[idx, "Phone"] = phone

                    # Progress + Timer
                    progress = (idx + 1) / len(df)
                    progress_bar.progress(progress)

                    elapsed = time.time() - start_time
                    avg_time_per_firm = elapsed / (idx + 1)
                    remaining_firms = len(df) - (idx + 1)
                    eta_seconds = remaining_firms * avg_time_per_firm

                    # Format time nicely
                    def format_seconds(seconds):
                        mins, secs = divmod(int(seconds), 60)
                        return f"{mins:02}:{secs:02}"

                    status_text.text(
                        f"Processing {idx + 1} of {len(df)} firms...\n"
                        f"Elapsed: {format_seconds(elapsed)}, "
                        f"ETA: {format_seconds(eta_seconds)}"
                    )

                browser.close()

                # Save to temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    df.to_excel(tmp.name, index=False)
                    tmp_path = tmp.name

                st.success("Done! Download your enriched file below.")
                with open(tmp_path, "rb") as f:
                    st.download_button("Download Enriched Excel", f, file_name="enriched_fca_register.xlsx")

                os.unlink(tmp_path)

if __name__ == "__main__":
    main()
