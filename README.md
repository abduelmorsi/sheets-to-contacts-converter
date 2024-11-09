# Contact Converter

A simple web application with a desktop application version to convert sheet files to contacts. Import contacts from Google Sheets or CSV files and export them as CSV (Google Contacts format) or VCF (vCard).

## 🌐 Live Demo
Access the web version at: [Contact Converter](https://aelmorsi.pythonanywhere.com/)

## ✨ Features
- Import contacts from:
  - Google Sheets
  - CSV files
- Export formats:
  - CSV (Google Contacts compatible)
  - VCF (vCard format)
- Add optional name prefix
- Simple column selection using letters

## 📝 How to Use
1. Choose your input method:
   - Paste a Google Sheets URL (must be published to web)
   - Upload a CSV file
2. (Optional) Add a name prefix
3. Specify column letters for:
   - Name column (e.g., A)
   - Phone column (e.g., B)
4. Select output format:
   - CSV for Google Contacts
   - VCF for direct import to most phones
5. Click Convert

## 🖥️ Desktop Version
Download the latest desktop version from the [Releases](../../releases) page.

## 🛠️ Running Locally
```bash
# Clone the repository
git clone https://github.com/abduelmorsi/sheets-to-contacts-converter.git

# Install dependencies
pip install -r requirements.txt

# Run the application
python app.py
```

## 📄 License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
```