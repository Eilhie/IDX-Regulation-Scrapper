import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd

# Email notification function
def send_email_notification(subject, body):
    sender_email = "maru.dev.purpose@gmail.com"  # Your email
    receiver_email = "b.sachio88@gmail.com"  # Receiver's email
    password = "ymkm xqeh iklc dgxy"  # Your app password (generated in Google Account)

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    try:
        # Set up the server
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, msg.as_string())  # Send the email
        server.close()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email. Error: {str(e)}")

# Function to create a new driver
def create_driver():
    return webdriver.Firefox()  # Provide the path to geckodriver if necessary

# Function to scrape the table data
def scrape_table(driver):
    try:
        # Wait for the table to be visible (this ensures the page has loaded)
        WebDriverWait(driver, 10).until(
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
                        # Check if the child is an <a> tag (link)
                        if child.tag_name == 'a':
                            link_url = child.get_attribute('href')
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

# Function to handle paging and scrape across pages
def handle_paging(driver):
    current_page = 1
    total_pages = 5  # Set the number of pages to scrape, or loop until no more pages
    all_data = []

    while current_page <= total_pages:
        print(f"Scraping page {current_page}...")

        # Scrape the table from the current page
        page_data = scrape_table(driver)
        all_data.extend(page_data)  # Add the page data to the main list

        try:
            # Wait until the 'Next' button is clickable
            next_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//a[@class="pagingButton fa fa-arrow-right"]'))
            )
            next_button.click()  # Simulate clicking the "Next" button

            # Wait for the next page to load fully
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//table'))  # Wait for table to appear on the next page
            )

            current_page += 1
            time.sleep(3)  # Optional, but it helps ensure the next page is ready

        except Exception as e:
            print("No more pages or error:", e)
            break

    return all_data

# Function to run the scraping every X minutes
def scrape_every_x_minutes(wait_minutes):
    wait_seconds = wait_minutes * 60  # Convert minutes to seconds

    while True:
        print("Starting scraping...")

        # Create a new driver instance
        driver = create_driver()
        driver.get("https://www.ojk.go.id/id/Regulasi/Default.aspx")  # Open the website

        # Ensure the page is fully loaded by checking document ready state
        WebDriverWait(driver, 20).until(
            lambda driver: driver.execute_script('return document.readyState') == 'complete'
        )

        # Wait for the table to fully appear
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//table'))  # Wait for the table to appear
        )

        # Scrape data and handle paging
        scraped_data = handle_paging(driver)

        # Convert the scraped data to a DataFrame
        columns = ['Nomor Regulasi', 'Judul Regulasi', 'Informasi Singkat Regulasi']  # Adjust column names if necessary
        df = pd.DataFrame(scraped_data, columns=columns)

        # Save the DataFrame to a CSV file
        csv_file = 'scraped_data.csv'
        df.to_csv(csv_file, index=False, encoding='utf-8')

        print(f"Completed scraping. Data exported to {csv_file}.")

        # Send email notification after scraping is done
        send_email_notification(
            subject="Scraping Completed",
            body=f"Your scraping job has finished. Data has been saved to {csv_file}."
        )

        # Close the driver after the scrape is finished
        driver.quit()

        print(f"Waiting for {wait_minutes} minutes...")
        time.sleep(wait_seconds)  # Wait for the specified number of minutes before starting again

        print("Restarting the scraping process...")

# Run the scraper every X minutes (e.g., 30 minutes)
wait_time_in_minutes = 30
scrape_every_x_minutes(wait_time_in_minutes)
