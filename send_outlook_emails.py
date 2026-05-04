import win32com.client
import csv
import os
import re
import time
import datetime

DEFAULT_MAX_EMAILS = 100
DEFAULT_BATCH_SIZE = 6
DEFAULT_BATCH_INTERVAL = 1800
DEFAULT_DELAY_IN_BATCH = 5
DEFAULT_SEND_FROM = "richard.m@ureka.co.uk"

PAUSE_FLAG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pause.flag")

def _is_paused():
    return os.path.exists(PAUSE_FLAG)

def _wait_while_paused():
    if not _is_paused():
        return
    print("\n⏸ PAUSED — delete 'pause.flag' to resume …")
    while _is_paused():
        time.sleep(2)
    print("▶ Resumed!\n")

def _countdown(seconds, label="Next batch"):
    end_time = time.time() + seconds
    while True:
        _wait_while_paused()
        remaining = int(end_time - time.time())
        if remaining <= 0:
            break
        mins, secs = divmod(remaining, 60)
        print(f"\r⏳ {label} in {mins:02d}:{secs:02d} …", end="", flush=True)
        time.sleep(1)
    print("\r" + " " * 50 + "\r", end="", flush=True)

def _update_csv_status(csv_path, email, status="Email Sent", sent_from="", email_column="Institution Email", status_column="Status"):
    today = datetime.datetime.now().strftime("%d %b %Y")
    rows = []
    
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        fieldnames = list(reader.fieldnames) if reader.fieldnames else []
        for row in reader:
            if row.get(email_column, "").strip().lower() == email.strip().lower():
                row[status_column] = status
                row["send Date"] = today
                if "Sent from" in fieldnames:
                    row["Sent from"] = sent_from
            rows.append(row)

    if "send Date" not in fieldnames:
        fieldnames.append("send Date")
    if status_column not in fieldnames:
        fieldnames.append(status_column)

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

_TITLE_RE = re.compile(r'^(mr\.?|ms\.?|mrs\.?|dr\.?|prof\.?|miss|sir|dame|rev\.?)\s+', re.IGNORECASE)

def _first_name(full_name: str) -> str:
    if not full_name or not full_name.strip():
        return ""
    name = _TITLE_RE.sub('', full_name.strip())
    parts = name.split() if name else full_name.strip().split()
    return parts[0] if parts else ""

def _get_account(outlook, email_address):
    try:
        namespace = outlook.GetNamespace("MAPI")
        for i in range(1, namespace.Accounts.Count + 1):
            acct = namespace.Accounts.Item(i)
            if acct.SmtpAddress.lower() == email_address.lower():
                return acct
    except Exception as e:
        print(f"Error reading accounts: {e}")
    return None

def send_emails_from_template(
    template_path,
    recipients,
    csv_path=None,
    max_emails=DEFAULT_MAX_EMAILS,
    batch_size=DEFAULT_BATCH_SIZE,
    batch_interval=DEFAULT_BATCH_INTERVAL,
    delay_in_batch=DEFAULT_DELAY_IN_BATCH,
    dry_run=False,
    send_from=None,
    sender_names=None,
    email_column="Institution Email",
    status_column="Status",
):
    if not os.path.exists(template_path):
        raise FileNotFoundError(template_path)

    if send_from is None:
        send_from = [DEFAULT_SEND_FROM]
    elif isinstance(send_from, str):
        send_from = [send_from]

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        raise RuntimeError(f"Could not connect to Outlook. Ensure Classic Outlook is running. Error: {e}")

    accounts = {}
    for addr in send_from:
        acct = _get_account(outlook, addr)
        if acct is None:
            raise RuntimeError(f"Account not found: {addr}")
        accounts[addr] = acct

    sent_by = {addr: 0 for addr in send_from}
    total_cap = max_emails * len(send_from)

    total_sent = 0
    remaining = list(recipients)
    robin_idx = 0

    while remaining and total_sent < total_cap:
        _wait_while_paused()

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except:
            pass

        batch = remaining[:batch_size]
        remaining = remaining[batch_size:]

        for recipient in batch:
            current_addr = send_from[robin_idx % len(send_from)]
            robin_idx += 1
            account = accounts[current_addr]

            try:
                mail = outlook.CreateItemFromTemplate(template_path)

                # --- SAFELY UPDATE RECIPIENTS ---
                # Remove ANY existing 'To' recipients saved in the template (Type 1 = olTo)
                # But KEEP the CC (Type 2 = olCC) and BCC (Type 3 = olBCC)
                for i in range(mail.Recipients.Count, 0, -1):
                    try:
                        if mail.Recipients.Item(i).Type == 1:
                            mail.Recipients.Remove(i)
                    except Exception:
                        pass

                # Cleanly add the new recipient
                target_email = recipient["email"].strip()
                new_recip = mail.Recipients.Add(target_email)
                new_recip.Type = 1 # Set as 'To'
                new_recip.Resolve() # Force Outlook to accept the email address

                # Optional: Force resolve everything else (like the CCs) to ensure no Outlook errors
                mail.Recipients.ResolveAll()

                # --- FORCE ACCOUNT ---
                try:
                    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                except Exception:
                    try:
                        mail.SendUsingAccount = account 
                    except:
                        pass

                # --- TEXT REPLACEMENTS ---
                try:
                    body = mail.HTMLBody
                    is_html = True
                except:
                    body = mail.Body
                    is_html = False

                if body:
                    preview_first = _first_name(recipient.get('full_name', ''))
                    short_name = len(preview_first) <= 3
                    
                    def _ireplace(text, old, new):
                        return re.compile(re.escape(old), re.IGNORECASE).sub(lambda m: new, text)

                    if short_name:
                        body = re.sub(r'(?i)Dear\s+\[First name\],', lambda m: 'Greetings,', body)
                    else:
                        body = _ireplace(body, "[First name]", preview_first)
                        
                    inst_name = recipient.get("institution_name", "")
                    body = _ireplace(body, "[Institution Name]", inst_name)
                    
                    if sender_names and current_addr in sender_names:
                        body = _ireplace(body, "[sender name]", sender_names[current_addr])
                    body = _ireplace(body, "[sender mail]", current_addr)
                    
                    if is_html:
                        mail.HTMLBody = body
                    else:
                        mail.Body = body

                # --- SAVE AND SEND ---
                mail.Save()

                if dry_run:
                    mail.Display()
                else:
                    mail.Send()

                total_sent += 1
                sent_by[current_addr] += 1

                print(f"[{total_sent}] {recipient['email']} via {current_addr}")

                if csv_path and not dry_run:
                    _update_csv_status(
                        csv_path,
                        recipient["email"],
                        sent_from=current_addr,
                        email_column=email_column,
                        status_column=status_column,
                    )

            except Exception as e:
                print(f"❌ Error processing {recipient.get('email', 'unknown')}: {e}")

            time.sleep(delay_in_batch)

        if remaining and total_sent < total_cap:
            _countdown(batch_interval)

    print(f"\nFinished — {total_sent} emails sent.")