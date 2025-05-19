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

def sanitize_filename(filename):
    # Rimuove caratteri non validi per Windows
    filename = filename.replace('\r', '').replace('\n', ' ')
    return re.sub(r'[<>:"/\\|?*]', '_', filename)

# Funzione per scrivere i dati delle email e salvare gli allegati
def process_and_save_email(msg, allegati_dir, address_set, contents):
    """
    Processa una singola email, estrae informazioni e salva gli allegati.

    Args:
        msg: L'oggetto email.Message da processare.
        allegati_dir: La directory dove salvare gli allegati.
        address_set: Il set per memorizzare gli indirizzi email unici.
        contents: La lista per memorizzare il contenuto completo delle email.
    """
    # Mittente e destinatari
    from_ = msg.get("From", "")
    to_ = msg.get("To", "")
    cc_ = msg.get("Cc", "")
    indirizzi = re.findall(r'[\w\.-]+@[\w\.-]+', from_ + to_ + cc_)
    address_set.update(indirizzi)

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
    attachments_list = []
    for part in msg.walk():
        if part.get_content_disposition() == "attachment":
            filename = part.get_filename()
            if filename:
                decoded_filename, encoding = decode_header(filename)[0]
                if isinstance(decoded_filename, bytes):
                    try:
                        filename = decoded_filename.decode(encoding or "utf-8")
                    except UnicodeDecodeError:
                        filename = decoded_filename.decode("latin-1")  # Tentativo con latin-1 se utf-8 fallisce
                safe_filename = sanitize_filename(filename)
                filepath = os.path.join(allegati_dir, safe_filename)
                with open(filepath, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                attachments_list.append(filename)

    # Costruisci contenuto completo
    email_text = f"From: {from_}\nTo: {to_}\nCc: {cc_}\nSubject: {subject}\n\nBody:\n{body.strip()}\n\nAllegati: {', '.join(attachments_list)}\n{'=' * 80}\n"
    contents.append(email_text)

def backup_single_folder(callback):
    # Inserisci qui le tue credenziali
    MAIL = os.getenv("MAIL")
    PASSWORD = os.getenv("PASSWORD")

    # Cartella principale per gli allegati
    ALLEGATI_DIR = "attachments"
    ADDRESS_DIR = "address"
    if not os.path.exists(ALLEGATI_DIR):
        os.makedirs(ALLEGATI_DIR)
    if not os.path.exists(ADDRESS_DIR):
        os.makedirs(ADDRESS_DIR)

    # Connessione
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(MAIL, PASSWORD)

    # Trova e seleziona la cartella
    folder_name = callback(mail)
    if not folder_name:
        print(f"❌ ERRORE: Impossibile trovare la cartella per la funzione {callback.__name__}.")
        mail.logout()
        return

    status, _ = mail.select(f'"{folder_name}"')
    if status != "OK":
        print(f"❌ ERRORE: Impossibile selezionare la cartella {folder_name}.")
        mail.logout()
        return

    # Crea una sottocartella specifica per questa cartella di posta
    folder_safe_name = re.sub(r'[^\w\s-]', '', folder_name).strip()  # Rimuove caratteri non alfanumerici
    allegati_dir = os.path.join(ALLEGATI_DIR, folder_safe_name)
    if not os.path.exists(allegati_dir):
        os.makedirs(allegati_dir)

    # Cerca tutte le email
    status, messages = mail.search(None, "ALL")
    email_ids = messages[0].split()

    # Per salvare output
    address_set = set()
    contents = []

    for email_id in email_ids:
        print(f"Processing mail id: {email_id}")
        res, msg_data = mail.fetch(email_id, "(RFC822)")
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                process_and_save_email(msg, allegati_dir, address_set, contents) # Usa la funzione

    # Scrivi i file (ora con nomi più specifici per la cartella)
    with open(f"{ADDRESS_DIR}/{folder_safe_name}.txt", "w", encoding="utf-8") as f:
        for email_addr in sorted(address_set):
            f.write(email_addr + "\n")

    with open(f"{ADDRESS_DIR}/contents_complets_{folder_safe_name}.txt", "w", encoding="utf-8") as f:
        f.writelines(contents)

    print(f"✅ Completato il backup della cartella '{folder_name}'. File generati: 'indirizzi_{folder_safe_name}.txt', 'contents_complets_{folder_safe_name}.txt' e allegati salvati in '{allegati_dir}'")
    mail.logout()

if __name__ == "__main__":
    list_callback_find = [find_inbox, find_sent, find_drafts, find_trash, find_all_mail, find_spam]
    for callback in list_callback_find: 
        backup_single_folder(callback)
