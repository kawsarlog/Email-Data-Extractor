# Email Data Extractor 📧

The Email Data Extractor is a Python program designed to gather relevant information from email bodies and store it in an Excel spreadsheet.

[![Python Version](https://img.shields.io/badge/python-3.9.7-blue.svg)](https://www.python.org/downloads/release/python-397/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Openpyxl](https://img.shields.io/badge/openpyxl-3.0.12-blue.svg)](https://openpyxl.readthedocs.io/en/stable/)
[![BeautifulSoup4](https://img.shields.io/badge/BeautifulSoup4-4.10.0-blue.svg)](https://www.crummy.com/software/BeautifulSoup/bs4/doc/)

## Features

- **IMAP Connection:** Establish a secure connection to the specified IMAP server using provided credentials.
- **Email Retrieval:** Search and retrieve all emails from the specified mailbox.
- **Information Extraction:** Extract relevant information such as subject, sender, date, and content from each email.
- **Duplicate Fixing:** Avoid duplicate entries by checking existing data in the Excel file.
- **Excel Saving:** Save the extracted information into an Excel spreadsheet.

## Dependencies

- `imaplib` for IMAP communication.
- `email` library for parsing email messages.
- `BeautifulSoup` library for parsing HTML content.

## Installation

```bash
pip install openpyxl bs4
```

## Usage

1. **Input Configuration:** Provide the IMAP server, username, and password in the program.
    ```python
    imap_server = "imaps.udag.de"
    username = "example@email.com"
    password = "YourPassword123"
    ```

2. **Run the Program:** Execute the program to connect to the email account, extract information, and save data into an Excel file.

3. **Avoiding Duplicates:** The program checks for duplicate entries, ensuring only unique information is stored.

4. **Output Excel File:** The Excel file serves as a structured dataset, containing non-duplicated information from email bodies.

## Contact
For any questions or feedback, feel free to reach out:

- Email: kawsar@kawsarlog.com
- Website: kawsarlog.com

---
