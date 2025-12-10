# Bulk Mail GUI – Premium Tkinter Edition

A modern, responsive, and easy-to-use **Bulk Email Sender** built with Python and Tkinter.  
This tool allows you to send personalized emails to multiple recipients using an Excel file, with full preview support, attachments, and a polished professional UI.

---

## 🚀 Features

### ✔ Premium GUI (Tkinter)
- Fully responsive layout  
- Auto-adjusts in full-screen mode  
- Clean styling with modern components  
- Organized panels: Compose, Preview, Attachments, Logs

### ✔ Excel-Based Bulk Emailing
- Load any `.xlsx` file  
- Select Email + Name columns  
- Preview first 10 recipients  
- Automatically detects common column names

### ✔ Personalized Messages
Supports placeholders inside email body:
Hi {Name},

Automatically replaced per row.

### ✔ Attachments Support
- Add multiple files  
- Preview attachments  
- Remove individually

### ✔ Preview Mode (Dry Run)
Test everything safely **without sending** actual emails.

### ✔ Real Send Mode
Send real emails using Gmail SMTP with App Password.

### ✔ Progress & Logging
- Live progress bar  
- Detailed log output  
- Error reporting  
- Summary after completion  

---

## 📦 Requirements

Install dependencies:

```bash
pip install pandas openpyxl
```

### ✔ This tool works on:
- Windows  
- macOS  
- Linux
(Only Python 3.9+ required)

## ▶️ How to Run
```bash
python bulk_mail_gui.py
```


The GUI will open automatically.

## 🔐 Gmail SMTP Setup
- To send emails using Gmail:
- Go to: https://myaccount.google.com/security
- Enable 2-Step Verification
- Go to App Passwords
- Create a password for "Mail"
- Use that password in your script (already integrated)

## 📁 Project Structure
```bash
bulk_mail_gui.py     # Main GUI application
README.md            # Project documentation
```
