import os 
import re
import io
import imaplib
import email
from email.header import decode_header , make_header
from pathlib import Path
from datetime import datetime
from docx import Document
import pandas as pd 
from dotenv import load_dotenv

#Resume Parsing 
import pdfplumber #used to extract text from pdfs
from docx import Document #used to extract text from docx files

load_dotenv()
IMAP_HOST = os.getenv("IMAP_HOST")
IMAP_PORT = int(os.getenv("IMAP_PORT","993"))
IMAP_USER = os.getenv("IMAP_USER")
IMAP_PASSWORD = os.getenv("IMAP_PASSWORD")
IMAP_FOLDER = os.getenv("IMAP_FOLDER","INBOX")
SUBJECT_KEYWORDS = os.getenv("SUBJECT_KEYWORDS","").strip().lower()
FROM_DOMAIN = os.getenv("FROM_DOMAIN","").strip().lower()

ATTACH_DIR = os.getenv("ATTACH_DIR","storage/attachments")
OUTPUT_XLSX = os.getenv("OUTPUT_XLSX","storage/applicants.xlsx") 

Path(ATTACH_DIR).mkdir(parents=True, exist_ok=True) #create attachment directory if not exists
Path(OUTPUT_XLSX).parent.mkdir(parents=True, exist_ok=True) #create output directory if not exists

#Helper functions
EMAIL_RE=re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
PHONE_RE=re.compile(r"\+?\d[\d -]{8,}\d")
YEAR_RE=re.compile(r"(\b\d{1,2}(\.\d+)?\s*(\+)?\s*(years|yrs|year)\b)", re.I)
LOCATION_HITS = ["location", "based in", "current location", "city", "address"]

def rating_from_years(years: float) -> str: #simple rating based on years of experience
    if years is None:
        return "Unknown"
    if years < 1:
        return "Entry Level"
    elif years < 2:
        return "Junior (0-2)"
    elif years < 5:
        return "Mid Level (2-5)"
    else:
        return "Senior (5+)"
    
def safe_decode(s): #decode email subjects properly
    try:
        return str(make_header(decode_header(s)))
    except:
        return s    
    
def connect_imap(): #connect to IMAP server and select folder
    mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    mail.login(IMAP_USER, IMAP_PASSWORD)
    mail.select(IMAP_FOLDER)
    return mail    

def search_unread_ids(mail): #search for unread emails
    status, data = mail.search(None, 'ALL') #search for unread emails
     #data is a list of email ids in bytes
    if status != "OK":
        print("No messages found!")
        return []
    ids=data[0].split()
    return ids  

def passes_filters(msg): #check if email passes subject and from domain filters
    subs = safe_decode(msg.get("Subject","")).lower()
    from_addr = safe_decode(msg.get("From","")).lower()
    if SUBJECT_KEYWORDS and SUBJECT_KEYWORDS not in subs:
        return False
    if FROM_DOMAIN and FROM_DOMAIN not in from_addr:
        return False
    return True

def save_attachment(part, email_id):  # save attachment to disk
    filename = part.get_filename()              # get the attachment's filename
    if filename:
        filename = safe_decode(filename)        # decode it (handles encoded names like =?UTF-8?...?=)
        filepath = os.path.join(ATTACH_DIR, f"{email_id}_{filename}")  
        # prefix with email ID to avoid duplicate names
        with open(filepath, "wb") as f:         # open file in binary write mode
            f.write(part.get_payload(decode=True))  # decode and write the actual file content
        return filepath
    return None

def text_from_pdf(path):
    text=[]
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text.append(page.extract_text())
    return "\n".join(text)

def text_from_docx(path):
    doc = Document(path)
    return "\n".join([para.text for para in doc.paragraphs])

def text_from_attachment(path): #extract text based on file type
    ext = Path(path).suffix.lower()
    try:
        if ext == ".pdf":
            return text_from_pdf(path)
        elif ext in [".docx", ".doc"]:
            return text_from_docx(path)
        else:
            print(f"Unsupported file type: {ext}")
            return ""
    except Exception as e:
        print(f"Error reading {path}: {e}")
        return ""
    
def extract_first(pattern,text): #extract first match of a regex pattern
    match = pattern.search(text)
    return match.group(0) if match else ""

def extract_years(text): #extract years of experience from text
    candidates =  [(m.group(0), m.start()) for m in YEAR_RE.finditer(text)] #find all matches with their positions
    if not candidates:
        return None
    
    values=[]
    for raw,_ in candidates:
        num_part = re.findall(r"\d+(\.\d+)?", raw)
        if num_part:
            try:
                values.append(float(num_part[0]))
            except Exception:
                pass 
    return max(values) if values else None
        
def extract_location(text): #extract location based on keywords
    lines = [ln.strip()  for ln in text.splitlines() if ln.strip()]   

    for ln in lines:
        lower = ln.lower()
        if any(h in lower for h in LOCATION_HITS):
            if ':' in ln:
                return ln.split(':',1)[1].strip()
            return ln   
    return None

def guess_name_from_email(email_addr):
    if not email_addr:
        return None
    local = email_addr.split('@')[0]
    return local.replace('.', ' ').replace('_', ' ').title()

def extract_applied_position(subject , body_text):
    subject = subject or ""
    body_text = body_text or ""

    m = re.search(r"(applying for|application for|applied for|position|role)[:\- ]+(.{3,80})", subject, re.I)
    if m:
        return m.group(2).strip(" .")
    m2 = re.search(r"(applying for|application for|applied for|position|role)[:\- ]+(.{3,80})", body_text, re.I)
    if m2:
        return m2.group(2).strip(" .")
    # Try to extract from subjeect
    if ":" in subject:
        after = subject.split(":",1)[1].strip()
        if 2<=len(after)<=80:
            return after
    return None

def get_email_body_text(msg):
    parts = []
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition") or "")
            if ctype == "text/plain" and "attachment" not in disp.lower():
                try:
                    parts.append(part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", errors="ignore"))
                except Exception:
                    pass
    else:
        try:
            parts.append(msg.get_payload(decode=True).decode(msg.get_content_charset() or "utf-8", errors="ignore"))
        except Exception:
            pass
    return "\n".join(parts)

def append_row_to_excel(row:dict , path:str):
    columns = ["Name","Email","Phone","YearsExperience","Location","AppliedPosition",
               "SourceEmailDate","EmailSubject","AttachmentPaths","Rating"]
    
    df = pd.DataFrame([row], columns=columns)
    if not os.path.exists(path): #if file doesn't exist, create new
        df.to_excel(path, index=False) #write to excel
    else:
        existing_df = pd.read_excel(path) #read existing data
        combined_df = pd.concat([existing_df, df], ignore_index=True) #append new row
        combined_df.to_excel(path, index=False) #write back to excel


def main():
    mail = connect_imap()
    ids = search_unread_ids(mail)
    print(f"Found {len(ids)} unread emails.")

    for eid in ids:
        status, data = mail.fetch(eid, "(RFC822)")
        if status != "OK" or not data or not data[0]:
            continue

        msg = email.message_from_bytes(data[0][1]) #parse email message from bytes
        if not passes_filters(msg): #apply subject and from domain filters
            continue
        subject = safe_decode(msg.get("Subject")) #decode subject , sometimes it is encoded in utf-8 or other formats
        from_hdr = safe_decode(msg.get("From"))
        date_hdr = safe_decode(msg.get("Date")) #decode date header
        msg_date = None

        try:
            msg_date = email.utils.parseaddr_to_datetime(date_hdr) #parse date to datetime object
        except Exception:
            msg_date = datetime.utcnow() 
        body_text = get_email_body_text(msg) #get email body text

        attachment_paths = []
        resume_texts = []

        if msg.is_multipart():
            for part in msg.walk():
                cdisp = str(part.get("Content-Disposition") or "").lower() #get content disposition , what kind of part it is
                if "attachment" in cdisp: #if it is an attachment
                    path = save_attachment(part, eid.decode()) #save attachment to disk
                    attachment_paths.append(path)
                    resume_texts.append(text_from_attachment(path)) #extract text from attachment

        full_text = "\n".join([body_text] + resume_texts) #combine body text and resume texts

        #Field Extraction
        email_found = extract_first(EMAIL_RE, full_text) or extract_first(EMAIL_RE, from_hdr)
        phone_found = extract_first(PHONE_RE, full_text)
        years = extract_years(full_text)
        location = extract_location(full_text)
        applied = extract_applied_position(subject, body_text)

        name = None 
        mname = re.search(r"\b(Name)\s*[:\-]\s*(.+)", full_text, re.I)
        if mname:
            name = mname.group(2).splitlines()[0].strip() #take first line after name:
        if not name:
            name = guess_name_from_email(email_found)

        rating = rating_from_years(years)

        row = {
            "Name": name or "Unknown",
            "Email": email_found or "Unknown",
            "Phone": phone_found or "Unknown",
            "YearsExperience": years if years is not None else "Unknown",
            "Location": location or "Unknown",
            "AppliedPosition": applied or "Unknown",
            "SourceEmailDate": msg_date.strftime("%Y-%m-%d %H:%M:%S") if msg_date else None,
            "EmailSubject": subject,
            "AttachmentPaths": ", ".join(attachment_paths),
            "Rating": rating
        }

        append_row_to_excel(row, OUTPUT_XLSX)

        print(f"Added: {row['Name']} | {row['Email']} | {row['Rating']}")

    mail.close()
    mail.logout()      
    print("Done processing emails.")

if __name__ == "__main__":
    main()      










    

    
