# Automated HR Inbox Scanning and Applicant Rating

##  Problem Statement
Recruitment teams often receive candidate applications directly through email.  
Manually opening each email, downloading attachments, and tracking applicant information in spreadsheets is **time-consuming**.  

This project automates the process of:
- Fetching applications from an HR email inbox
- Extracting structured applicant details
- Saving data into Excel for HR review
- Classifying applicants into **Junior, Mid-level, or Senior** categories based on experience

---

## Features
- Connects to Gmail/Outlook inbox via **IMAP**
- Downloads unread resumes (PDF/DOCX)
- Extracts key details:
  - Name  
  - Email  
  - Phone  
  - Years of Experience  
  - Location  
  - Applied Position  
- Stores details in an **Excel file**
- Applies **experience-based rating**:
  - Entry Level (<1 year)  
  - Junior (0–2 years)  
  - Mid-level (2–5 years)  
  - Senior (5+ years)  

---

##  Tech Stack
- **Python 3.9+**
- Libraries:  
  - `imaplib`, `email` → Fetch emails  
  - `pdfplumber`, `python-docx` → Parse resumes  
  - `pandas`, `openpyxl` → Excel storage  
  - `dotenv` → Secure environment variables  

---

