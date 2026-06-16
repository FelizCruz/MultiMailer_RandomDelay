import csv

from load_recipients import load_from_csv
from send_outlook_emails import _first_name, send_emails_from_template


# Fixed paths
TEMPLATE_PATH = r"C:\Enter_path\abc.oft"
DATABASE_PATH = r"C:\Enter_path - Sheet1.csv"

# Defaults (user can override interactively)
DEFAULT_SEND_FROM = ["EnterMail@outlook.com"]
DEFAULT_MAX_EMAILS = 450  # per account
DEFAULT_BATCH_SIZE = 300
DEFAULT_BATCH_INTERVAL = 1800  # seconds
DEFAULT_BATCH_INTERVAL_MAX = 2000  # seconds
DELAY_IN_BATCH = 2  # seconds between Outlook messages inside a batch (min)
DELAY_IN_BATCH_MAX = 5  # seconds between Outlook messages inside a batch (max)
DRY_RUN = False  # True = open drafts for review
DEFAULT_NAME_COLUMN = "Contacted Person Name"
DEFAULT_EMAIL_COLUMN = "Institution Email"
DEFAULT_STATUS_COLUMN = "Status"
DEFAULT_SHORT_NAME_FILTER = True
DEFAULT_SEND_MODE = "individual"
DEFAULT_BCC_BATCH_SIZE = 35


def _ask(prompt, default, cast=str):
    """Show prompt with [default]; return cast value or default on blank."""
    raw = input(f"{prompt} [{default}]: ").strip()
    if not raw:
        return default
    try:
        return cast(raw)
    except ValueError:
        print(f"  Invalid input, using default ({default})")
        return default


def _ask_range(label, default_min, default_max, cast=int):
    """Prompt for a min/max range. Returns (min_val, max_val)."""
    print(f"\n  {label}")
    while True:
        lo = _ask("    Min (seconds)", default_min, cast)
        hi = _ask("    Max (seconds)", default_max, cast)
        if lo <= hi:
            if lo == hi:
                print(f"    Fixed delay of {lo}s (no randomisation)")
            else:
                print(f"    Will randomise between {lo}s and {hi}s")
            return lo, hi
        print(f"  Min ({lo}) must be <= Max ({hi}). Please try again.")


def _ask_yes_no(prompt, default=True):
    """Prompt for a yes/no value and return True for yes, False for no."""
    default_label = "Y/n" if default else "y/N"
    while True:
        raw = input(f"{prompt} [{default_label}]: ").strip().lower()
        if not raw:
            return default
        if raw in ("y", "yes"):
            return True
        if raw in ("n", "no"):
            return False
        print("  Please enter y or n.")


def _ask_send_mode(default=DEFAULT_SEND_MODE):
    """Prompt for the send mode. Defaults to the existing individual mode."""
    while True:
        raw = input(f"Sending mode: individual or bcc [{default}]: ").strip().lower()
        selected = raw or default
        if selected in ("individual", "bcc"):
            return selected
        print("  Please enter 'individual' or 'bcc'.")


def _get_csv_headers(csv_path):
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        return reader.fieldnames or []


def _ask_column(prompt, default, headers):
    print(f"\nAvailable columns: {', '.join(headers)}")
    while True:
        choice = input(f"{prompt} [{default}]: ").strip()
        selected = choice or default
        if selected in headers:
            return selected
        print("  Column not found. Please enter one of the listed columns exactly.")


def main():
    print("=" * 55)
    print("  Mailer - Interactive Setup")
    print("=" * 55)

    print(f"\nDefault send-from: {', '.join(DEFAULT_SEND_FROM)}")
    raw = input(
        "Enter send-from email(s) separated by commas\n"
        "(leave blank to keep defaults): "
    ).strip()

    if raw:
        send_from_list = [e.strip() for e in raw.split(",") if e.strip()]
    else:
        send_from_list = list(DEFAULT_SEND_FROM)

    print(f"  Will round-robin across {len(send_from_list)} account(s):")
    for addr in send_from_list:
        print(f"      - {addr}")

    print("\nEnter the display name for each sending account:")
    sender_names = {}
    for addr in send_from_list:
        name = input(f"  Name for {addr}: ").strip()
        if not name:
            name = addr.split("@")[0].replace(".", " ").title()
            print(f"    (defaulting to '{name}')")
        sender_names[addr] = name

    print()
    send_mode = _ask_send_mode(DEFAULT_SEND_MODE)
    bcc_batch_size = DEFAULT_BCC_BATCH_SIZE
    if send_mode == "bcc":
        bcc_batch_size = _ask("BCC recipients per message", DEFAULT_BCC_BATCH_SIZE, int)

    max_emails = _ask("Max emails PER ACCOUNT", DEFAULT_MAX_EMAILS, int)
    batch_size = _ask("Batch size", DEFAULT_BATCH_SIZE, int)
    filter_short_names = _ask_yes_no(
        "Replace names of 3 characters or fewer with 'Greetings,'",
        DEFAULT_SHORT_NAME_FILTER,
    )

    delay_min, delay_max = _ask_range(
        "Delay between Outlook messages WITHIN a batch:",
        DELAY_IN_BATCH,
        DELAY_IN_BATCH_MAX,
    )
    interval_min, interval_max = _ask_range(
        "Delay between BATCHES:",
        DEFAULT_BATCH_INTERVAL,
        DEFAULT_BATCH_INTERVAL_MAX,
    )

    total_possible = max_emails * len(send_from_list)
    print(
        f"\n  {max_emails} recipient(s) x {len(send_from_list)} account(s) = "
        f"{total_possible} total possible recipient(s)"
    )
    if send_mode == "bcc":
        print(
            "  BCC mode will send up to "
            f"{bcc_batch_size} hidden recipient(s) per Outlook message"
        )

    headers = _get_csv_headers(DATABASE_PATH)
    if not headers:
        print("Could not read CSV headers. Please verify the database file.")
        return

    name_column = _ask_column("Column for recipient name", DEFAULT_NAME_COLUMN, headers)
    email_column = _ask_column("Column for recipient email", DEFAULT_EMAIL_COLUMN, headers)
    status_column = _ask_column("Column for send status", DEFAULT_STATUS_COLUMN, headers)

    recipients = load_from_csv(
        DATABASE_PATH,
        name_column=name_column,
        email_column=email_column,
        status_column=status_column,
    )
    print(f"\nFound {len(recipients)} unsent recipient(s) in database.")

    if not recipients:
        print("Nothing to send - all rows already have a status.")
        return

    first = recipients[0]
    first_account = send_from_list[0]
    print(f"\n{'-' * 55}")
    print("  PREVIEW - first email that will be sent:")
    print(f"    From : {sender_names[first_account]} <{first_account}>")
    if send_mode == "bcc":
        hidden_count = min(bcc_batch_size, len(recipients), total_possible)
        print(f"    To   : {first_account}")
        print(f"    BCC  : {hidden_count} hidden recipient(s) in first mini-batch")
    else:
        print(f"    To   : {first['email']}")
        print(f"    Name : {first['full_name']}")

    print(f"    Mode : {'DRY RUN (drafts)' if DRY_RUN else 'LIVE (sends immediately)'}")
    print(f"    Sending mode : {send_mode}")
    if send_mode == "individual":
        print(f"    Short-name filter : {'ON' if filter_short_names else 'OFF'}")

    print("    Placeholders replaced:")
    if send_mode == "bcc":
        print("      Dear [First name],  -> Greetings,")
        print("      [Institution Name]  -> (blank)")
    else:
        preview_first = _first_name(first["full_name"])
        short_name = len(preview_first) <= 3
        if filter_short_names and short_name:
            print(
                "      Dear [First name],  -> Greetings,  "
                f"(name '{preview_first}' is <=3 chars)"
            )
        else:
            print(f"      [First name]        -> {preview_first}")
        print(f"      [Institution Name]  -> {first.get('institution_name', '')}")

    print(f"      [sender name]       -> {sender_names[first_account]}")
    print(f"      [sender mail]       -> {first_account}")
    if send_mode == "bcc":
        print("    Quota/status count : by recipient, not by Outlook message")
    print(f"{'-' * 55}")

    confirm = input("\nProceed? (y/n): ").strip().lower()
    if confirm not in ("y", "yes"):
        print("Aborted.")
        return

    send_emails_from_template(
        TEMPLATE_PATH,
        recipients,
        csv_path=DATABASE_PATH,
        max_emails=max_emails,
        batch_size=batch_size,
        batch_interval=interval_min,
        batch_interval_max=interval_max,
        delay_in_batch=delay_min,
        delay_in_batch_max=delay_max,
        dry_run=DRY_RUN,
        send_from=send_from_list,
        sender_names=sender_names,
        email_column=email_column,
        status_column=status_column,
        filter_short_names=filter_short_names,
        send_mode=send_mode,
        bcc_batch_size=bcc_batch_size,
    )


if __name__ == "__main__":
    main()
