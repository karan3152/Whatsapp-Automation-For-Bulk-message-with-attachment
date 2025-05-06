import os
import time
import random
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
import urllib.parse
import sys

# Add the parent directory to sys.path to import file_manager
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
import file_manager


def setup_driver():
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_argument('--disable-notifications')
    chrome_options.add_argument('--start-maximized')
    chrome_options.add_argument('--disable-popup-blocking')
    driver = uc.Chrome(options=chrome_options)
    return driver


def load_contacts(file_path):
    try:
        # Load contacts from Excel file
        df = pd.read_excel(file_path, engine='openpyxl')
        df.columns = df.columns.str.strip().str.upper()  # Normalize column names
        required_columns = {'MOBILE'}
        if required_columns.issubset(df.columns):
            return df[['MOBILE']].to_dict(orient='records')
        else:
            print("Error: Required columns ( 'MOBILE') are missing.")
            return []
    except Exception as e:
        print(f"Error loading contacts: {e}")
        return []


def load_message_template(file_path):
    """Read the message template from a text file."""
    try:
        with open(file_path, 'r') as file:
            message_template = file.read()
        return message_template
    except Exception as e:
        print(f"Error reading message template: {e}")
        return None


def format_message(contact, message_template):
    """Generate personalized message."""
    return message_template.format(
        # name=contact['NAME'],
        # uan=contact['UAN'],
        mobile=contact['MOBILE']
    )


def send_message(driver, contact, message):
    try:
        phone_number = str(contact['MOBILE']).strip().replace(" ", "").replace("-", "").replace("+", "")
        encoded_message = urllib.parse.quote(message)

        # Navigate to the WhatsApp Web URL for the contact
        url = f"https://web.whatsapp.com/send?phone={phone_number}&text={encoded_message}"
        driver.get(url)
        time.sleep(random.uniform(1,4))  # Wait for the page to load

        # Wait for the message box and send button
        message_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
        )
        send_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//span[@data-icon="send"]'))
        )

        # Click on the send button
        send_button.click()
        time.sleep(random.uniform(1,4))  # Short delay after sending

        # Verify if the message was sent (a simple confirmation can be added here)
        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.XPATH, '//span[@data-icon="send"]'))
        )
        print(f"Message successfully sent to {contact['MOBILE']}")
        return True

    except TimeoutException:
        print(f"Timeout occurred for contact: {contact['MOBILE']}")
        return False
    except WebDriverException as e:
        print(f"Error sending message to {contact['MOBILE']}: {e}")
        return False

import os

def retry_failed_contacts(driver, contacts, message, max_retries=3):
    failed_contacts = []

    for attempt in range(max_retries):
        current_failed_contacts = []
        for contact in contacts:
            print(f"Retrying ({attempt + 1}/{max_retries}): {contact['MOBILE']}")
            success = send_message(driver, contact, message)
            if not success:
                current_failed_contacts.append(contact)

        if not current_failed_contacts:
            print("All messages sent successfully!")
            return True  # All messages sent successfully
        else:
            print(f"Retrying failed contacts ({len(current_failed_contacts)})...")

        contacts = current_failed_contacts  # Update contacts to failed ones for the next attempt
        failed_contacts.extend(current_failed_contacts)  # Add to the overall failed list

    # Log failed contacts after all retries
    if failed_contacts:
        print("Failed to send messages to the following contacts after retries:")
        for contact in failed_contacts:
            print(contact['MOBILE'])

        # Check if the file exists
        base_path = os.path.dirname(os.path.abspath(__file__))
        failed_contacts_file = os.path.join(base_path, "Failed_Contacts.xlsx")

        if os.path.exists(failed_contacts_file):
            # Read existing file
            existing_df = pd.read_excel(failed_contacts_file, engine='openpyxl')
            new_df = pd.DataFrame(failed_contacts)

            # Concatenate the old and new dataframes
            updated_df = pd.concat([existing_df, new_df], ignore_index=True)

            # Backup the existing file before overwriting
            file_manager.delete_excel_file(failed_contacts_file, backup=True)

            # Save the updated dataframe
            updated_df.to_excel(failed_contacts_file, index=False)
        else:
            # Create a new file if it doesn't exist
            file_manager.create_empty_excel(failed_contacts_file, columns=['MOBILE'])
            failed_contacts_df = pd.DataFrame(failed_contacts)
            failed_contacts_df.to_excel(failed_contacts_file, index=False)

        print(f"Failed contacts have been logged into '{failed_contacts_file}'.")

    return False

def main():
    # File paths
    base_path = os.path.dirname(os.path.abspath(__file__))
    excel_file = os.path.join(base_path, "UAN.xlsx")
    message_template_file = os.path.join(base_path, "message.txt")
    failed_contacts_file = os.path.join(base_path, "Failed_Contacts.xlsx")

    # Create backup and delete old Excel files
    if os.path.exists(excel_file):
        file_manager.delete_excel_file(excel_file, backup=True)

    # Create new empty Excel file with table structure
    file_manager.create_empty_excel(excel_file, columns=['MOBILE'])

    # Create backup and delete old text files
    if os.path.exists(message_template_file):
        file_manager.delete_text_file(message_template_file, backup=True)

    # Create new empty text file
    default_message = "Hello,\n\nThis is a message for {mobile}.\n\nRegards,\nHR Team"
    file_manager.create_empty_text_file(message_template_file, content=default_message)

    # Delete old failed contacts file if it exists
    if os.path.exists(failed_contacts_file):
        file_manager.delete_excel_file(failed_contacts_file, backup=True)
        file_manager.create_empty_excel(failed_contacts_file, columns=['MOBILE'])

    print("All files have been reset. New empty files have been created.")

    # Load contacts (will be empty since we just created a new file)
    contacts = load_contacts(excel_file)

    if not contacts:
        print("No contacts found in the newly created Excel file.")
        print("Please add contacts to the Excel file and run the program again.")
        return

    # Load message template
    message_template = load_message_template(message_template_file)
    if not message_template:
        print("Error loading message template. Exiting.")
        return

    # Setup browser driver
    driver = setup_driver()
    try:
        driver.get("https://web.whatsapp.com")
        input("Scan the QR code and press Enter to continue...")
        time.sleep(10)  # Wait for user to log in

        for i, contact in enumerate(contacts, start=1):
            print(f"Sending message to ({i}/{len(contacts)}): {contact['MOBILE']}")
            message = format_message(contact, message_template)
            success = send_message(driver, contact, message)
            if not success:
                print(f"Adding {contact['MOBILE']} to retry list")
                retry_failed_contacts(driver, [contact], message)  # Retry for failed contact
            time.sleep(random.uniform(1,2))  # Short delay between messages
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
