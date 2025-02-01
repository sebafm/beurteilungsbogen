# Excel to Word and Email Automation

## 📌 Overview
This project automates the process of generating Word documents from an Excel file and sending them via email to supervisors. The script:

1. Reads an Excel file containing employee data.
2. Groups employees based on their supervisor.
3. Generates Word documents for each employee using a template.
4. Sends the generated documents as email attachments to the respective supervisors.

## 🚀 Features
- **Excel Parsing:** Reads employee data from an Excel file.
- **Dynamic Word Document Creation:** Replaces placeholders in a Word template with actual employee data.
- **Automated Email Sending:** Sends emails with the generated documents as attachments.
- **Multi-threaded Processing:** Enhances efficiency by processing multiple supervisors in parallel.
- **Unit Testing:** Ensures correctness with built-in tests.

## 📂 Project Structure
```
├── email_utils.py          # Handles email creation and sending
├── excel_to_word_email.py  # Main script executing the workflow
├── word_utils.py           # Functions to generate Word documents from a template
├── tests/                  # Unit tests for the functionalities
├── Mitarbeiterdaten.xlsx   # Example employee data in Excel format
├── README.md               # Project documentation
```

## 🛠️ Installation
### Prerequisites
- Python 3.7+
- Required libraries: `pandas`, `python-docx`, `smtplib`
- An SMTP server for sending emails

### Install Dependencies
```sh
pip install pandas python-docx
```

## 📖 Usage
1. **Prepare the Excel file**
   - Ensure the Excel file (`Mitarbeiterdaten.xlsx`) contains the necessary columns:
     - `Vorgesetzter_Email`, `Name`, `Personalnummer`, `Dienststelle`, etc.

2. **Modify the Word template**
   - Use placeholders such as `[Name]`, `[Personalnummer]` in the Word document.

3. **Run the script**
```sh
python excel_to_word_email.py
```

## 📝 Configuration
### SMTP Email Settings
Modify `email_utils.py` to include your SMTP settings:
```python
smtp_server = "smtp.example.com"
smtp_port = 587
sender_email = "your-email@example.com"
sender_password = "your-password"
```

## ✅ Running Unit Tests
To verify the script’s correctness, run:
```sh
python -m unittest discover tests
```

## 🤝 Contribution
Feel free to fork this repository, make improvements, and submit a pull request.

## 📜 License
This project is licensed under the MIT License.
