# WhatsApp Sender Application

An automated tool for sending WhatsApp messages and attachments to multiple contacts.

## Features

- Send personalized text messages to multiple contacts
- Attach and send PDF files along with messages
- Automatically handle failed message attempts
- User-friendly interface with WhatsApp Web integration

## Requirements

- Python 3.6+
- Chrome browser
- Internet connection
- WhatsApp account

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/whatsapp-sender-application.git
   cd whatsapp-sender-application
   ```

2. Create a virtual environment:
   ```
   python -m venv .venv
   ```

3. Activate the virtual environment:
   - Windows:
     ```
     .venv\Scripts\activate
     ```
   - macOS/Linux:
     ```
     source .venv/bin/activate
     ```

4. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Prepare your contacts:
   - Add phone numbers to the `contacts.xlsx` file
   - Make sure the file has a column named "MOBILE"

2. Customize your message:
   - Edit the `Message.txt` file with your desired message
   - You can use `{mobile}` as a placeholder for the recipient's number

3. Add attachments (optional):
   - Place PDF files in the project directory
   - Update the file paths in `main.py` if needed

4. Run the application:
   ```
   python main.py
   ```

5. When prompted, scan the QR code with your WhatsApp mobile app to log in to WhatsApp Web

## Project Structure

- `main.py`: Main application script
- `file_manager.py`: Utility for file operations
- `contacts.xlsx`: Excel file containing contact numbers
- `Message.txt`: Template for the message to be sent
- `Failed_Contacts.xlsx`: Records of failed message attempts

## Notes

- The application uses WhatsApp Web, so your phone must be connected to the internet
- Rate limiting may apply based on WhatsApp's policies
- Use responsibly and respect privacy laws and regulations

## License

MIT License

## Disclaimer

This tool is for educational purposes only. The developers are not responsible for any misuse or violation of WhatsApp's terms of service.
