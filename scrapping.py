import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options as FirefoxOptions
import time
import pandas as pd
import PyPDF2
import requests
import os 
import re
from concurrent.futures import ThreadPoolExecutor
import calendar 
import schedule
from datetime import datetime
import threading
import openpyxl
import json


# Global DataFrame variable
df = pd.DataFrame()

def send_email_notification(subject, body, attachment_paths=None):
    from_address = "maru.dev.purpose@gmail.com"
    to_address = "b.sachio88@gmail.com"
    password = "ymkm xqeh iklc dgxy"

    # Create the email message
    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = to_address
    msg['Subject'] = subject

    # Attach the body of the email
    msg.attach(MIMEText(body, 'plain'))

    if attachment_paths:
        for attachment_path in attachment_paths:
            if os.path.exists(attachment_path):
                # Attach the file
                attachment = MIMEBase('application', 'octet-stream')
                with open(attachment_path, 'rb') as file:
                    attachment.set_payload(file.read())
                encoders.encode_base64(attachment)
                attachment.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
                msg.attach(attachment)
            else:
                print(f"Warning: {attachment_path} not found, skipping attachment.")

    # Send the email via SMTP
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_address, password)
        server.sendmail(from_address, to_address, msg.as_string())
        server.quit()
        print(f"Email sent to {to_address}")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Create a new driver with headless Firefox
def create_driver():
    firefox_options = FirefoxOptions()
    firefox_options.add_argument("--headless")
    driver = webdriver.Firefox(options=firefox_options)
    return driver

# Helper function to safely extract text from elements
def safe_extract(driver, by, class_name):
    try:
        return driver.find_element(by, class_name).text.strip()
    except Exception:
        return 'Not Found'

# Scrape data from the main table
def scrape_table(driver):
    try:
        # Wait for the table to be visible (this ensures the page has loaded)
        WebDriverWait(driver, 50).until(
            EC.presence_of_element_located((By.XPATH, '//table'))  # Adjust XPath to the table
        )

        # Locate the table and rows
        table = driver.find_element(By.XPATH, '//table')  # Adjust the XPath if necessary
        rows = table.find_elements(By.XPATH, './/tr')

        # Initialize lists for storing data
        all_rows = []

        # Iterate through each row
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, 'td')  # Find all cells in the row
            row_data = []

            for cell in cells:
                # Extract text from the cell
                child_elements = cell.find_elements(By.XPATH, './*')  # Check for child elements
                if child_elements:
                    for child in child_elements:
                        row_data.append(child.text.strip())  # Append the text to row_data
                        # Check if the child is an <strong> tag (link)
                        if child.tag_name == 'strong':
                            link = child.find_element(By.TAG_NAME, 'a')  # Find the <a> tag
                            link_url = link.get_attribute('href')
                            row_data.append(link_url)
                else:
                    row_data.append(cell.text.strip())  # Append text if no child elements

            # Append the row data to the main list if not empty
            if row_data:
                all_rows.append(row_data)

        # Return the collected data
        return all_rows

    except Exception as e:
        print("Error while scraping the table:", e)
        return []

# Handle pagination and scrape across multiple pages
def handle_paging(driver):
    current_page = 1
    total_pages = 5
    all_data = []

    while current_page <= total_pages:
        print(f"Scraping page {current_page}...")
        page_data = scrape_table(driver)
        all_data.extend(page_data)
        
        try:
            next_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, '//a[@class="pagingButton fa fa-arrow-right"]'))
            )
            next_button.click()
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//table')))
            current_page += 1
            time.sleep(3)
        except Exception:
            break
    return all_data

# Sanitize folder name for downloads
def sanitize_title(title):
    return re.sub(r'[<>:"/\\|?*]', '_', re.sub(r'\s+', ' ', title).strip())

# Download and merge PDFs concurrently using ThreadPoolExecutor
def download_pdfs(pdf_urls, download_folder):
    with ThreadPoolExecutor() as executor:
        return list(executor.map(lambda url: download_pdf(url, download_folder), pdf_urls))

# Helper function to download individual PDFs
def download_pdf(file_url, download_folder):
    response = requests.get(file_url)
    file_name = os.path.join(download_folder, file_url.split("/")[-1])
    with open(file_name, 'wb') as file:
        file.write(response.content)
    return file_name

def load_existing_data(excel_file):
    """Load existing data from Excel and return a DataFrame."""
    try:
        return pd.read_excel(excel_file)
    except FileNotFoundError:
        return pd.DataFrame()  # Return an empty DataFrame if the file doesn't exist

def check_for_new_entries_and_updates(new_data, existing_data):
    """Check for new entries or updates based on 'Nomor Ketentuan'."""
    new_entries = []
    updated_entries = []
    existing_ids = existing_data['Nomor Ketentuan'].values if not existing_data.empty else []

    for index, row in new_data.iterrows():
        if row['Nomor Ketentuan'] not in existing_ids:
            # New entry
            new_entries.append(row)
        else:
            # Check for updates (comparing specific fields like 'Tentang' or 'Tanggal Berlaku')
            existing_row = existing_data[existing_data['Nomor Ketentuan'] == row['Nomor Ketentuan']].iloc[0]
            if existing_row['Tentang'] != row['Tentang'] or existing_row['Tanggal Regulasi Efektif'] != row['Tanggal Regulasi Efektif']:
                updated_entries.append(row)
    
    return pd.DataFrame(new_entries), pd.DataFrame(updated_entries)


def save_to_excel_rewrite(df, file_name):
    """Rewrite the Excel file with new data, creating it if it doesn't exist."""
    try:
        # Ensure the file path is valid
        if not file_name:
            raise ValueError("File path is empty or invalid")

        # Check if the directory exists, if not create it
        dir_name = os.path.dirname(file_name)
        if dir_name and not os.path.exists(dir_name):
            os.makedirs(dir_name)
            print(f"Directory {dir_name} created.")
        
        # Write the DataFrame to the Excel file, replacing any existing content
        df.to_excel(file_name, index=False, sheet_name='Sheet1')

        print(f"Data written to {file_name} successfully.")

    except Exception as e:
        print(f"Error writing to {file_name}: {e}")

def download_and_merge_pdfs(pdf_urls, regulation_id, regulation_title):
    """
    Downloads PDFs from a list of URLs, merges them if there are multiple files,
    and returns the path to the merged PDF.

    Args:
        pdf_urls (list): List of PDF URLs to download.
        regulation_id (str): Unique identifier for the regulation (used for folder naming).
        regulation_title (str): Title of the regulation (used for file naming).

    Returns:
        str: Path to the merged PDF file, or None if no PDFs are downloaded.
    """
    sanitized_title = sanitize_title(regulation_title)
    download_folder = os.path.join("downloads", sanitize_title(regulation_id))
    os.makedirs(download_folder, exist_ok=True)

    pdf_files = []
    for i, url in enumerate(pdf_urls):
        try:
            pdf_file_path = os.path.join(download_folder, f"{sanitized_title}_{i + 1}.pdf")
            # Download the PDF from the URL
            response = requests.get(url, stream=True)
            response.raise_for_status()
            with open(pdf_file_path, 'wb') as pdf_file:
                pdf_file.write(response.content)
            pdf_files.append(pdf_file_path)
            print(f"Downloaded: {pdf_file_path}")
        except Exception as e:
            print(f"Failed to download PDF from {url}. Error: {e}")

    if not pdf_files:
        print("No PDFs were downloaded.")
        return None

    # If there are multiple PDFs, merge them
    if len(pdf_files) > 1:
        merged_pdf_path = os.path.join(download_folder, f"merged_{sanitized_title}.pdf")
        try:
            merger = PyPDF2.PdfMerger()
            for pdf_file in pdf_files:
                merger.append(pdf_file)
            with open(merged_pdf_path, 'wb') as merged_pdf:
                merger.write(merged_pdf)
            print(f"Merged PDF saved to: {merged_pdf_path}")
            return merged_pdf_path
        except Exception as e:
            print(f"Failed to merge PDFs. Error: {e}")
            return None
    else:
        # If only one PDF was downloaded, return its path
        return pdf_files[0]

def OJK_regulation_scraper():
    global df

    # Load existing data and tracking metadata
    existing_data = load_existing_data('regulasi_report.xlsx')
    pdf_tracking_file = "pdf_tracking.json"

    if not os.path.exists(pdf_tracking_file) or os.stat(pdf_tracking_file).st_size == 0:
        # If file doesn't exist or is empty, initialize it
        with open(pdf_tracking_file, 'w') as f:
            json.dump({}, f)

    # Load the JSON file
    with open(pdf_tracking_file, 'r') as f:
        pdf_tracking = json.load(f)

    driver = create_driver()
    driver.get("https://www.ojk.go.id/id/Regulasi/Default.aspx")
    WebDriverWait(driver, 20).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//table')))

    scraped_data = handle_paging(driver)
    columns = ['Nomor Regulasi', 'Judul Regulasi', 'URL Regulasi', 'Informasi Singkat Regulasi']
    df = pd.DataFrame(scraped_data, columns=columns)

    # Add 'Regulator' column filled with 'OJK'
    df['Regulator'] = 'OJK'

    new_regulations = []

    for index, row in df.iterrows():
        regulation_url = row['URL Regulasi']
        regulation_id = row['Nomor Regulasi']

        if pd.notnull(regulation_url):
            try:
                driver.get(regulation_url)
                WebDriverWait(driver, 15).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')

                # Extract regulation details
                details = {key: safe_extract(driver, By.CLASS_NAME, cls) for key, cls in {
                    'Sektor': 'sektor-regulasi-display',
                    'Subsektor': 'subsektor-regulasi-display',
                    'Jenis': 'jenis-regulasi-display',
                    'Nomor': 'nomor-regulasi-display',
                    'Tanggal Berlaku': 'display-date-text.tanggal-2'
                }.items()}

                # Scrape PDF URLs
                div = driver.find_element(By.ID, 'ctl00_PlaceHolderMain_ctl04__ControlWrapper_RichHtmlField')
                pdf_links = div.find_elements(By.TAG_NAME, 'a')
                pdf_urls = [link.get_attribute('href') for link in pdf_links if link.get_attribute('href').endswith('.pdf')]

                # Detect new or updated regulations
                if regulation_id in pdf_tracking:
                    tracked_entry = pdf_tracking[regulation_id]
                    # Check if details or PDF URLs have changed
                    if tracked_entry['details'] != details or tracked_entry['pdf_urls'] != pdf_urls:
                        print(f"Update detected for {regulation_id}. Re-downloading PDFs.")
                        updated_files = download_and_merge_pdfs(pdf_urls, regulation_id, row['Judul Regulasi'])
                        pdf_tracking[regulation_id] = {
                            'details': details,
                            'pdf_urls': pdf_urls,
                            'merged_pdf': updated_files
                        }
                    else:
                        print(f"No updates detected for {regulation_id}. Skipping PDF download.")
                else:
                    print(f"New regulation detected: {regulation_id}. Downloading PDFs.")
                    merged_pdf = download_and_merge_pdfs(pdf_urls, regulation_id, row['Judul Regulasi'])
                    pdf_tracking[regulation_id] = {
                        'details': details,
                        'pdf_urls': pdf_urls,
                        'merged_pdf': merged_pdf
                    }
                    new_regulations.append(row['Nomor Regulasi'])

                # Renaming and filling columns as per the new requirements
                df.loc[index, 'Nomor Ketentuan'] = row['Nomor Regulasi']
                df.loc[index, 'Tanggal Regulasi Efektif'] = details.get('Tanggal Berlaku')
                df.loc[index, 'Tentang'] = row['Judul Regulasi']
                df.loc[index, 'Topic'] = details.get('Sektor')
                df.loc[index, 'Jenis'] = details.get('Jenis')
                
            except Exception as e:
                print(f"Failed to scrape details for {regulation_url}. Error: {e}")
                for key in ['Sektor', 'Subsektor', 'Jenis', 'Nomor', 'Tanggal', 'PDF Files']:
                    df.loc[index, key] = 'Failed to retrieve'
    
    driver.quit()

    # Save tracking metadata
    with open('pdf_tracking.json', 'w') as f:
        json.dump(pdf_tracking, f)

    # Reorder the columns to match the desired format
    df = df[['Regulator', 'Nomor Ketentuan', 'Tanggal Regulasi Efektif', 'Tentang', 'Topic', 'Jenis']]

    # Check for new entries and updates
    new_entries, updated_entries = check_for_new_entries_and_updates(df, existing_data)

    if not new_entries.empty:
        print(f"Ada {len(new_entries)} entri baru.")
        save_to_excel_rewrite(df, 'regulasi_report.xlsx')
        
        # Send email notification with PDFs for new regulations
        subject = "New OJK Regulations Detected"
        body = f"The following new regulations have been detected: {', '.join(len(new_regulations))}."
        print(new_regulations)

        # Create a list of attachment paths for new regulations
        attachment_paths = ['regulasi_report.xlsx']
        for regulation_id in new_regulations:
            # Path to the folder containing PDFs for the regulation
            folder_path = f'downloads/{regulation_id}/'
            if os.path.exists(folder_path):
                # Get all files inside the folder
                for file_name in os.listdir(folder_path):
                    file_path = os.path.join(folder_path, file_name)
                    attachment_paths.append(file_path)

        # Send email with attachments
        send_email_notification(subject, body, attachment_paths)
    
    if not updated_entries.empty:
        print(f"Ada {len(updated_entries)} pembaruan.")
        for _, updated_row in updated_entries.iterrows():
            existing_data.loc[existing_data['Nomor Ketentuan'] == updated_row['Nomor Ketentuan']] = updated_row
        existing_data.to_excel('regulasi_report.xlsx', index=False)

    if new_entries.empty and updated_entries.empty:
        print("Tidak ada entri baru atau pembaruan.")

    return df


def is_last_working_day_of_week(date):
    # Check if it's the last working day of the week (Friday)
    return date.weekday() == 4  # 4 corresponds to Friday

def is_last_working_day_of_month(date):
    # Check if it's the last weekday of the month
    last_day_of_month = calendar.monthrange(date.year, date.month)[1]
    last_date = datetime.date(date.year, date.month, last_day_of_month)
    return date == last_date or last_date.weekday() in [5, 6] and date.weekday() == 4  # If last day is weekend, check Friday

def is_last_working_day_of_year(date):
    # Check if it's the last weekday of the year
    last_day_of_year = datetime.date(date.year, 12, 31)
    return date == last_day_of_year or last_day_of_year.weekday() in [5, 6] and date.weekday() == 4

def create_excel_report(df, file_name, report_type="daily"):
    today = datetime.today()
    
    # Save DataFrame to Excel using xlsxwriter engine
    with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Regulasi Laporan')
        
        # Access workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Regulasi Laporan']
        
        # Set column widths and wrap text
        wrap_format = workbook.add_format({'text_wrap': True, 'align': 'center', 'valign': 'vcenter'})
        
        worksheet.set_column('A:A', 15, wrap_format)  # Regulator
        worksheet.set_column('B:B', 20, wrap_format)  # Nomor Ketentuan
        worksheet.set_column('C:C', 25, wrap_format)  # Tanggal Regulasi Terbit
        worksheet.set_column('D:D', 50, wrap_format)  # Hal yang diatur oleh ketentuan (tentang)
        worksheet.set_column('E:E', 20, wrap_format)  # Jenis Regulasi

        # Add header format
        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter'})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

    print(f"Report {report_type} Created and Saved As {file_name}")

def export_report(period):
    global df
    today = datetime.today()

    # Determine the filename based on the report type
    if period == "daily" and today.weekday() < 5:  # Monday to Friday
        file_name = f"report_daily_{today.strftime('%Y-%m-%d')}.xlsx"
        create_excel_report(df, file_name, period)
        print(f"Daily report exported to {file_name}")
        
        # Send email notification with the daily report
        subject = f"Daily OJK Regulations Report - {today.strftime('%Y-%m-%d')}"
        body = "The daily OJK regulations report has been generated and is attached."
        send_email_notification(subject, body, attachment_path=file_name)

    elif period == "weekly" and is_last_working_day_of_week(today):
        file_name = f"report_weekly_{today.strftime('%Y-%m-%d')}.xlsx"
        create_excel_report(df, file_name, period)
        print(f"Weekly report exported to {file_name}")
        
        # Send email notification with the weekly report
        subject = f"Weekly OJK Regulations Report - {today.strftime('%Y-%m-%d')}"
        body = "The weekly OJK regulations report has been generated and is attached."
        send_email_notification(subject, body, attachment_path=file_name)
    
    elif period == "monthly" and is_last_working_day_of_month(today):
        file_name = f"report_monthly_{today.strftime('%Y-%m-%d')}.xlsx"
        create_excel_report(df, file_name, period)
        print(f"Monthly report exported to {file_name}")
        
        # Send email notification with the monthly report
        subject = f"Monthly OJK Regulations Report - {today.strftime('%Y-%m-%d')}"
        body = "The monthly OJK regulations report has been generated and is attached."
        send_email_notification(subject, body, attachment_path=file_name)
    
    elif period == "yearly" and is_last_working_day_of_year(today):
        file_name = f"report_yearly_{today.strftime('%Y-%m-%d')}.xlsx"
        create_excel_report(df, file_name, period)
        print(f"Yearly report exported to {file_name}")
        
        # Send email notification with the yearly report
        subject = f"Yearly OJK Regulations Report - {today.strftime('%Y-%m-%d')}"
        body = "The yearly OJK regulations report has been generated and is attached."
        send_email_notification(subject, body, attachment_path=file_name)
    
    else:
        print(f"Not the last working day for {period}. No report exported.")


def schedule_reports():
    global df
    # Schedule daily report
    schedule.every().day.at("17:00").do(export_report, df, "daily")
    
    # Schedule weekly report (last Friday of the week)
    schedule.every().friday.at("17:00").do(export_report, df, "weekly")
    
    # Schedule monthly report (check if it's the last weekday of the month)
    schedule.every().day.at("17:00").do(
        lambda: export_report(df, "monthly") if is_last_working_day_of_month() else None
    )

    # Schedule yearly report (check if it's the last weekday of the year)
    schedule.every().day.at("17:00").do(
        lambda: export_report(df, "yearly") if is_last_working_day_of_year() else None
    )
    while True:
        schedule.run_pending()
        time.sleep(60)  # Check every minute if it's time to run any task


# Main function to scrape every X minutes
def scrape_every_x_minutes(wait_minutes):
    global df
    wait_seconds = wait_minutes * 60

    while True:
        print("Starting scraping...")
        df = OJK_regulation_scraper()  # Your scraping function
        
        # Export the scraped data (for demonstration, exporting to CSV)
        csv_file = 'scraped_data.csv'
        df.to_csv(csv_file, index=False, encoding='utf-8')
        print(f"Completed scraping. Data exported to {csv_file}.")

        # Export the scraped data (for demonstration, exporting to Excel with formatting)
        excel_file = 'scraped_data.xlsx'
        create_excel_report(df, excel_file, "scraped")
        print(f"Data exported to {excel_file}.")

        # Wait for the next scraping cycle
        print(f"Waiting for {wait_minutes} minutes...")
        time.sleep(wait_seconds)
        print("Restarting the scraping process...")

# Run the scraper every X minutes (e.g., 30 minutes)
# Main function to run everything
def run():
    # Start report scheduling in a separate thread
    report_thread = threading.Thread(target=schedule_reports)
    report_thread.daemon = True  # This makes the thread exit when the main program exits
    report_thread.start()
    
    # Start scraping every X minutes (e.g., every 30 minutes)
    scrape_every_x_minutes(30)

# Run the script
run()