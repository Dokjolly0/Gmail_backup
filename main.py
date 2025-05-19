import imaplib
import email
from email.header import decode_header
import os
from dotenv import load_dotenv
import re
from bs4 import BeautifulSoup

load_dotenv()

# Trova automaticamente la cartella Trash/Cestino
def find_recycle_bin(mail):
    status, mailboxes = mail.list()
    for mbox in mailboxes:
        nome = mbox.decode()
        if any(chiave in nome.lower() for chiave in ["trash", "cestino"]):
            # Estrai il nome effettivo della cartella tra virgolette
            match = re.search(r'"([^"]+)"$', nome)
            if match:
                return match.group(1)
    return None

def find_mailbox_by_keywords(mail, keywords):
    status, mailboxes = mail.list()
    if status != "OK":
        return None

    for mbox in mailboxes:
        nome = mbox.decode()
        if any(kw in nome.lower() for kw in keywords):
            match = re.search(r'"([^"]+)"$', nome)
            if match:
                return match.group(1)
    return None

def find_inbox(mail):
    return find_mailbox_by_keywords(mail, ["inbox", "posta in arrivo"])

def find_sent(mail):
    return find_mailbox_by_keywords(mail, ["sent", "posta inviata", "inviati"])

def find_drafts(mail):
    return find_mailbox_by_keywords(mail, ["drafts", "bozze"])

def find_trash(mail):
    return find_mailbox_by_keywords(mail, ["trash", "cestino", "deleted"])

def find_all_mail(mail):
    return find_mailbox_by_keywords(mail, ["all mail", "[gmail]/all mail", "tutti"])

def find_spam(mail):
    return find_mailbox_by_keywords(mail, ["spam", "posta indesiderata", "junk"])

def find_custom_label(mail, label_name):
    status, mailboxes = mail.list()
    if status != "OK":
        return None

    label_name_lower = label_name.lower()
    for mbox in mailboxes:
        nome = mbox.decode()
        if label_name_lower in nome.lower():
            match = re.search(r'"([^"]+)"$', nome)
            if match:
                return match.group(1)
    return None

# Inserisci qui le tue credenziali
MAIL = os.getenv("MAIL")
PASSWORD = os.getenv("PASSWORD")

# Cartella per salvare gli allegati
ALLEGATI_DIR = "allegati_email"
if not os.path.exists(ALLEGATI_DIR):
    os.makedirs(ALLEGATI_DIR)

# Connessione
mail = imaplib.IMAP4_SSL("imap.gmail.com")
mail.login(MAIL, PASSWORD)

# Trova e seleziona la cartella Cestino/Trash
dir_to_find = find_sent(mail)
if not dir_to_find:
    print("❌ ERRORE: Nessuna cartella Inviati trovata.")
    exit()

status, _ = mail.select(f'"{dir_to_find}"')
if status != "OK":
    print(f"❌ ERRORE: Impossibile selezionare la cartella {dir_to_find}.")
    exit()

# Cerca tutte le email
status, messages = mail.search(None, "ALL")
email_ids = messages[0].split()

# Per salvare output
indirizzi_set = set()
contenuti = []

for email_id in email_ids:
    res, msg_data = mail.fetch(email_id, "(RFC822)")
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])

            # Mittente e destinatari
            from_ = msg.get("From", "")
            to_ = msg.get("To", "")
            cc_ = msg.get("Cc", "")

            indirizzi = re.findall(r'[\w\.-]+@[\w\.-]+', from_ + to_ + cc_)
            indirizzi_set.update(indirizzi)

            # Oggetto
            subject, encoding = decode_header(msg.get("Subject", ""))[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding or "utf-8", errors="ignore")

            # Corpo
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    ctype = part.get_content_type()
                    disp = str(part.get("Content-Disposition"))
                    if ctype == "text/plain" and "attachment" not in disp:
                        charset = part.get_content_charset()
                        body = part.get_payload(decode=True).decode(charset or "utf-8", errors="ignore")
                        break
                    elif ctype == "text/html" and "attachment" not in disp:
                        charset = part.get_content_charset()
                        html = part.get_payload(decode=True).decode(charset or "utf-8", errors="ignore")
                        soup = BeautifulSoup(html, "html.parser")
                        body = soup.get_text()
                        break
            else:
                body = msg.get_payload(decode=True).decode(errors="ignore")

            # Allegati
            allegati_lista = []
            for part in msg.walk():
                if part.get_content_disposition() == "attachment":
                    filename = part.get_filename()
                    if filename:
                        decoded_filename, encoding = decode_header(filename)[0]
                        if isinstance(decoded_filename, bytes):
                            try:
                                filename = decoded_filename.decode(encoding or "utf-8")
                            except UnicodeDecodeError:
                                filename = decoded_filename.decode("latin-1") # Tentativo con latin-1 se utf-8 fallisce
                        filepath = os.path.join(ALLEGATI_DIR, filename)
                        with open(filepath, 'wb') as f:
                            f.write(part.get_payload(decode=True))
                        allegati_lista.append(filename)

            # Costruisci contenuto completo
            email_text = f"From: {from_}\nTo: {to_}\nCc: {cc_}\nSubject: {subject}\n\nBody:\n{body.strip()}\n\nAllegati: {', '.join(allegati_lista)}\n{'='*80}\n"
            contenuti.append(email_text)

# Scrivi i file
with open("indirizzi.txt", "w", encoding="utf-8") as f:
    for email_addr in sorted(indirizzi_set):
        f.write(email_addr + "\n")

with open("contenuti_completi.txt", "w", encoding="utf-8") as f:
    f.writelines(contenuti)

print(f"✅ Completato! File generati: 'indirizzi.txt', 'contenuti_completi.txt' e salvati gli allegati nella cartella '{ALLEGATI_DIR}'")