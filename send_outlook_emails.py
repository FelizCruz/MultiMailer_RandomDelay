import csv
import datetime
import os
import random
import re
import time

try:
    import win32com.client
except ImportError:
    win32com = None


DEFAULT_MAX_EMAILS = 96
DEFAULT_BATCH_SIZE = 7
DEFAULT_BATCH_INTERVAL = 1500
DEFAULT_DELAY_IN_BATCH = 5
DEFAULT_SEND_FROM = "EnterMail@outlook.com"
DEFAULT_BCC_BATCH_SIZE = 35
SEND_MODE_INDIVIDUAL = "individual"
SEND_MODE_BCC = "bcc"

PAUSE_FLAG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pause.flag")


def _is_paused():
    return os.path.exists(PAUSE_FLAG)


def _wait_while_paused():
    if not _is_paused():
        return
    print("\nPAUSED - delete 'pause.flag' to resume ...")
    while _is_paused():
        time.sleep(2)
    print("Resumed!\n")


def _countdown(seconds, label="Next batch"):
    end_time = time.time() + seconds
    while True:
        _wait_while_paused()
        remaining = int(end_time - time.time())
        if remaining <= 0:
            break
        mins, secs = divmod(remaining, 60)
        print(f"\r{label} in {mins:02d}:{secs:02d} ...", end="", flush=True)
        time.sleep(1)
    print("\r" + " " * 50 + "\r", end="", flush=True)


def _update_csv_status(
    csv_path,
    email,
    status="Email Sent",
    sent_from="",
    email_column="Institution Email",
    status_column="Status",
):
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


_TITLE_RE = re.compile(
    r"^(mr\.?|ms\.?|mrs\.?|dr\.?|prof\.?|miss|sir|dame|rev\.?)\s+",
    re.IGNORECASE,
)


def _first_name(full_name: str) -> str:
    if not full_name or not full_name.strip():
        return ""
    name = _TITLE_RE.sub("", full_name.strip())
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


def _ireplace(text, old, new):
    return re.compile(re.escape(old), re.IGNORECASE).sub(lambda m: new, text)


def _clear_recipient_types(mail, recipient_types):
    for i in range(mail.Recipients.Count, 0, -1):
        try:
            if mail.Recipients.Item(i).Type in recipient_types:
                mail.Recipients.Remove(i)
        except Exception:
            pass


def _add_recipient(mail, email_address, recipient_type):
    recipient = mail.Recipients.Add(email_address.strip())
    recipient.Type = recipient_type
    recipient.Resolve()
    return recipient


def _force_account(mail, account):
    try:
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
    except Exception:
        try:
            mail.SendUsingAccount = account
        except Exception:
            pass


def _get_body(mail):
    try:
        return mail.HTMLBody, True
    except Exception:
        return mail.Body, False


def _set_body(mail, body, is_html):
    if is_html:
        mail.HTMLBody = body
    else:
        mail.Body = body


def _apply_body_replacements(
    body,
    recipient,
    current_addr,
    sender_names=None,
    filter_short_names=True,
    send_mode=SEND_MODE_INDIVIDUAL,
):
    if not body:
        return body

    if send_mode == SEND_MODE_BCC:
        body = re.sub(r"(?i)Dear\s+\[First name\],", lambda m: "Greetings,", body)
        body = _ireplace(body, "[First name]", "")
        body = _ireplace(body, "[Institution Name]", "")
    else:
        preview_first = _first_name(recipient.get("full_name", ""))
        short_name = len(preview_first) <= 3

        if filter_short_names and short_name:
            body = re.sub(r"(?i)Dear\s+\[First name\],", lambda m: "Greetings,", body)
        else:
            body = _ireplace(body, "[First name]", preview_first)

        inst_name = recipient.get("institution_name", "")
        body = _ireplace(body, "[Institution Name]", inst_name)

    if sender_names and current_addr in sender_names:
        body = _ireplace(body, "[sender name]", sender_names[current_addr])
    body = _ireplace(body, "[sender mail]", current_addr)

    return body


def _validate_send_mode(send_mode):
    send_mode = (send_mode or SEND_MODE_INDIVIDUAL).strip().lower()
    if send_mode not in (SEND_MODE_INDIVIDUAL, SEND_MODE_BCC):
        raise ValueError("send_mode must be 'individual' or 'bcc'")
    return send_mode


def _next_account_with_capacity(send_from, sent_by, max_emails, robin_idx):
    for _ in range(len(send_from)):
        current_addr = send_from[robin_idx % len(send_from)]
        robin_idx += 1
        if sent_by[current_addr] < max_emails:
            return current_addr, robin_idx
    return None, robin_idx


def send_emails_from_template(
    template_path,
    recipients,
    csv_path=None,
    max_emails=DEFAULT_MAX_EMAILS,
    batch_size=DEFAULT_BATCH_SIZE,
    batch_interval=DEFAULT_BATCH_INTERVAL,
    batch_interval_max=None,
    delay_in_batch=DEFAULT_DELAY_IN_BATCH,
    delay_in_batch_max=None,
    dry_run=False,
    send_from=None,
    sender_names=None,
    email_column="Institution Email",
    status_column="Status",
    filter_short_names=True,
    send_mode=SEND_MODE_INDIVIDUAL,
    bcc_batch_size=DEFAULT_BCC_BATCH_SIZE,
):
    if not os.path.exists(template_path):
        raise FileNotFoundError(template_path)

    send_mode = _validate_send_mode(send_mode)
    if send_mode == SEND_MODE_BCC and bcc_batch_size < 1:
        raise ValueError("bcc_batch_size must be at least 1")
    if batch_size < 1:
        raise ValueError("batch_size must be at least 1")
    if max_emails < 1:
        raise ValueError("max_emails must be at least 1")

    if send_from is None:
        send_from = [DEFAULT_SEND_FROM]
    elif isinstance(send_from, str):
        send_from = [send_from]

    try:
        if win32com is None:
            raise ImportError("pywin32 is not installed")
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        raise RuntimeError(
            "Could not connect to Outlook. Ensure Classic Outlook is running. "
            f"Error: {e}"
        )

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
        except Exception:
            pass

        batch_capacity = min(batch_size, total_cap - total_sent)
        batch = remaining[:batch_capacity]
        remaining = remaining[batch_capacity:]

        while batch and total_sent < total_cap:
            current_addr, robin_idx = _next_account_with_capacity(
                send_from,
                sent_by,
                max_emails,
                robin_idx,
            )
            if current_addr is None:
                break

            account = accounts[current_addr]
            available_for_account = max_emails - sent_by[current_addr]
            group_size = 1
            if send_mode == SEND_MODE_BCC:
                group_size = min(bcc_batch_size, available_for_account, len(batch))

            recipient_group = batch[:group_size]
            batch = batch[group_size:]
            representative = recipient_group[0]

            try:
                mail = outlook.CreateItemFromTemplate(template_path)

                if send_mode == SEND_MODE_BCC:
                    _clear_recipient_types(mail, {1, 3})
                    _add_recipient(mail, current_addr, 1)
                    for recipient in recipient_group:
                        _add_recipient(mail, recipient["email"], 3)
                else:
                    _clear_recipient_types(mail, {1})
                    _add_recipient(mail, representative["email"], 1)

                mail.Recipients.ResolveAll()
                _force_account(mail, account)

                body, is_html = _get_body(mail)
                body = _apply_body_replacements(
                    body,
                    representative,
                    current_addr,
                    sender_names=sender_names,
                    filter_short_names=filter_short_names,
                    send_mode=send_mode,
                )
                _set_body(mail, body, is_html)

                mail.Save()

                if dry_run:
                    mail.Display()
                else:
                    mail.Send()

                sent_count = len(recipient_group)
                total_sent += sent_count
                sent_by[current_addr] += sent_count

                if send_mode == SEND_MODE_BCC:
                    print(
                        f"[{total_sent}] BCC batch of {sent_count} recipient(s) "
                        f"via {current_addr}"
                    )
                else:
                    print(f"[{total_sent}] {representative['email']} via {current_addr}")

                if csv_path and not dry_run:
                    for recipient in recipient_group:
                        _update_csv_status(
                            csv_path,
                            recipient["email"],
                            sent_from=current_addr,
                            email_column=email_column,
                            status_column=status_column,
                        )

            except Exception as e:
                if send_mode == SEND_MODE_BCC:
                    target = f"BCC batch starting with {representative.get('email', 'unknown')}"
                else:
                    target = representative.get("email", "unknown")
                print(f"Error processing {target}: {e}")

            if total_sent >= total_cap:
                break

            _delay = delay_in_batch
            if delay_in_batch_max and delay_in_batch_max > delay_in_batch:
                _delay = random.uniform(delay_in_batch, delay_in_batch_max)
            time.sleep(_delay)

        if remaining and total_sent < total_cap:
            _interval = batch_interval
            if batch_interval_max and batch_interval_max > batch_interval:
                _interval = random.uniform(batch_interval, batch_interval_max)
            _countdown(_interval)

    print(f"\nFinished - {total_sent} recipient(s) sent.")
