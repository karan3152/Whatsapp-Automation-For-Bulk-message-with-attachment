# import os
# import time
# import random
# import pandas as pd
# import undetected_chromedriver as uc
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.common.exceptions import TimeoutException, WebDriverException
# import urllib.parse


# def setup_driver():
#     chrome_options = uc.ChromeOptions()
#     chrome_options.add_argument('--disable-blink-features=AutomationControlled')
#     chrome_options.add_argument('--disable-notifications')
#     chrome_options.add_argument('--start-maximized')
#     chrome_options.add_argument('--disable-popup-blocking')
#     driver = uc.Chrome(options=chrome_options)
#     return driver


# def load_contacts(file_path):
#     try:
#         df = pd.read_excel(file_path, engine='openpyxl')
#         df.columns = df.columns.str.strip().str.upper()
#         required_columns = {'MOBILE'}
#         if required_columns.issubset(df.columns):
#             return df[['MOBILE']].to_dict(orient='records')
#         else:
#             print("Error: Required columns (MOBILE) are missing.")
#             return []
#     except Exception as e:
#         print(f"Error loading contacts: {e}")
#         return []


# def load_message_template(file_path):
#     try:
#         with open(file_path, 'r') as file:
#             message_template = file.read()
#         return message_template
#     except Exception as e:
#         print(f"Error reading message template: {e}")
#         return None


# def format_message(contact, message_template):
#     return message_template.format(
#         mobile=contact['MOBILE']
#     )


# def send_message(driver, contact, message):
#     try:
#         phone_number = str(contact['MOBILE']).strip().replace(" ", "").replace("-", "").replace("+", "")
#         encoded_message = urllib.parse.quote(message)

#         # Navigate to the WhatsApp Web URL for the contact
#         url = f"https://web.whatsapp.com/send?phone={phone_number}&text={encoded_message}"
#         driver.get(url)
#         time.sleep(random.uniform(3,10))

#         # Wait for the message box and send button
#         WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
#         )
#         send_button = WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.XPATH, '//span[@data-icon="send"]'))
#         )
#         send_button.click()
#         time.sleep(random.uniform(1,5))

#         print(f"Message sent to {contact['MOBILE']}")
#         return True
#     except TimeoutException:
#         print(f"Timeout occurred for contact: {contact['MOBILE']}")
#         return False
#     except WebDriverException as e:
#         print(f"Error sending message to {contact['MOBILE']}: {e}")
#         return False


# def send_photo(driver, contact, attachment_path):
#     try:
#         # Click on the attachment button (paperclip icon)
#         attachment_button = WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.XPATH, '//button[@aria-label="Attach"]'))
#         )
#         attachment_button.click()
#         time.sleep(random.uniform(1, 3))  # Short delay to allow dropdown menu to appear

#         # Locate the file input for attaching photos
#         file_input = WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.XPATH, '//input[@type="file"]'))
#         )

#         file_input.send_keys(attachment_path)  # Upload the photo
#         time.sleep(random.uniform(2, 5))  # Wait for the file to be uploaded

#         # Click the send button
#         send_button = WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.XPATH, '//span[@data-icon="send"]'))
#         )
#         send_button.click()
#         time.sleep(random.uniform(1, 3))  # Short delay after sending the photo

#         print(f"Photo sent to {contact['MOBILE']}")
#         return True
#     except TimeoutException:
#         print(f"Timeout occurred while sending photo to {contact['MOBILE']}")
#         return False
#     except WebDriverException as e:
#         print(f"Error sending photo to {contact['MOBILE']}: {e}")
#         return False
# def main():
#     # File paths
#     base_path = r"E:\HRTS_Application\Whatsapp_Sender Application\4.Whatsapp_message  version 3\4.Project_WP"
#     excel_file = os.path.join(base_path, "contacts.xlsx")
#     message_template_file = os.path.join(base_path, "Message.txt")
#     attachment_paths = [
#         os.path.join(base_path, "1 New Joining Application.pdf"),
#         os.path.join(base_path, "2 PF From 2 Revised.pdf"),
#         os.path.join(base_path, "3 Form 1 Nomination & Declaration Form.pdf"),
#         os.path.join(base_path, "4 Form11Revised.pdf")
#     ]  # Replace with the actual file path for the attachment

#     # Load contacts
#     contacts = load_contacts(excel_file)
#     if not contacts:
#         print("No contacts found. Exiting.")
#         return

#     # Load message template
#     message_template = load_message_template(message_template_file)
#     if not message_template:
#         print("Error loading message template. Exiting.")
#         return

#     # Setup browser driver
#     driver = setup_driver()
#     try:
#         driver.get("https://web.whatsapp.com")
#         input("Scan the QR code and press Enter to continue...")
#         time.sleep(10)  # Wait for user to log in

#         for i, contact in enumerate(contacts, start=1):
#             print(f"Sending message to ({i}/{len(contacts)}): {contact['MOBILE']}")
#             message = format_message(contact, message_template)

#             # Send text message first
#             message_sent = send_message(driver, contact, message)

#             # Send photos with message if text message was sent successfully
#             if message_sent and attachment_paths:
#                 for attachment_path in attachment_paths:
#                     photo_sent = send_photo(driver, contact, attachment_path)
#                     if not photo_sent:
#                         print(f"Failed to send photo to {contact['MOBILE']}")
#             elif not message_sent:
#                 print(f"Skipping photo upload for {contact['MOBILE']} due to text message failure.")

#             time.sleep(random.uniform(3, 5))  # Short delay between messages
#     finally:
#         try:
#             driver.quit()
#         except Exception as e:
#             print(f"Error during driver quit: {e}")


# if __name__ == "__main__":
#     main()

# # This is the main code which send photo and text message to the contact using WhatsApp web so don't dare to touch






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
        # Reading the contacts file using pandas
        df = pd.read_excel(file_path, engine='openpyxl')
        df.columns = df.columns.str.strip().str.upper()  # Clean up column names
        required_columns = {'MOBILE'}
        if required_columns.issubset(df.columns):
            return df[['MOBILE']].to_dict(orient='records')
        else:
            print("Error: Required columns (MOBILE) are missing.")
            return []
    except Exception as e:
        print(f"Error loading contacts: {e}")
        return []

def load_message_template(file_path):
    try:
        # Open the message template with utf-8 encoding
        with open(file_path, 'r', encoding='utf-8') as file:
            message_template = file.read()
        return message_template
    except UnicodeDecodeError as e:
        print(f"Error: Unable to decode the message template file. {e}")
        return None
    except Exception as e:
        print(f"Error reading message template: {e}")
        return None

def format_message(contact, message_template):
    return message_template.format(
        mobile=contact['MOBILE']
    )

def send_message(driver, contact, message):
    try:
        phone_number = str(contact['MOBILE']).strip().replace(" ", "").replace("-", "").replace("+", "")
        encoded_message = urllib.parse.quote(message)

        # Navigate to the WhatsApp Web URL for the contact
        url = f"https://web.whatsapp.com/send?phone={phone_number}&text={encoded_message}"
        driver.get(url)
        time.sleep(random.uniform(3,10))

        # Wait for the message box and send button
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
        )
        send_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//span[@data-icon="send"]'))
        )
        send_button.click()
        time.sleep(random.uniform(1,5))

        print(f"Message sent to {contact['MOBILE']}")
        return True
    except TimeoutException:
        print(f"Timeout occurred for contact: {contact['MOBILE']}")
        return False
    except WebDriverException as e:
        print(f"Error sending message to {contact['MOBILE']}: {e}")
        return False

def send_photo(driver, contact, attachment_path):
    try:
        # Click on the attachment button (paperclip icon)
        attachment_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//button[@aria-label="Attach"]'))
        )
        attachment_button.click()
        time.sleep(random.uniform(1, 3))  # Short delay to allow dropdown menu to appear

        # Locate the file input for attaching photos
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@type="file"]'))
        )

        file_input.send_keys(attachment_path)  # Upload the photo
        time.sleep(random.uniform(2, 5))  # Wait for the file to be uploaded

        # Click the send button
        send_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//span[@data-icon="send"]'))
        )
        send_button.click()
        time.sleep(random.uniform(1, 3))  # Short delay after sending the photo

        print(f"Photo sent to {contact['MOBILE']}")
        return True
    except TimeoutException:
        print(f"Timeout occurred while sending photo to {contact['MOBILE']}")
        return False
    except WebDriverException as e:
        print(f"Error sending photo to {contact['MOBILE']}: {e}")
        return False

def main():
    # File paths
    base_path = os.path.dirname(os.path.abspath(__file__))
    excel_file = os.path.join(base_path, "contacts.xlsx")
    message_template_file = os.path.join(base_path, "Message.txt")
    failed_contacts_file = os.path.join(base_path, "Failed_Contacts.xlsx")

    # PDF files
    pdf_files = [
        os.path.join(base_path, "1 New Joining Application.pdf"),
        os.path.join(base_path, "2 PF From 2 Revised.pdf"),
        os.path.join(base_path, "3 Form 1 Nomination & Declaration Form.pdf"),
        os.path.join(base_path, "4 Form11Revised.pdf")
    ]

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

    # Handle PDF files - backup existing ones
    for pdf_file in pdf_files:
        if os.path.exists(pdf_file):
            file_manager.handle_pdf_file(pdf_file, action="backup")

    print("All files have been reset. New empty files have been created.")

    # Set the attachment paths for sending
    attachment_paths = [pdf_file for pdf_file in pdf_files if os.path.exists(pdf_file)]

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

            # Send text message first
            message_sent = send_message(driver, contact, message)

            # Send photos with message if text message was sent successfully
            if message_sent and attachment_paths:
                for attachment_path in attachment_paths:
                    photo_sent = send_photo(driver, contact, attachment_path)
                    if not photo_sent:
                        print(f"Failed to send photo to {contact['MOBILE']}")
            elif not message_sent:
                print(f"Skipping photo upload for {contact['MOBILE']} due to text message failure.")

            time.sleep(random.uniform(3, 5))  # Short delay between messages
    finally:
        try:
            driver.quit()
        except Exception as e:
            print(f"Error during driver quit: {e}")

if __name__ == "__main__":
    main()
