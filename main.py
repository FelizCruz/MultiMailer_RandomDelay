from send_outlook_emails import send_emails_from_template, _first_name
from load_recipients import load_from_csv
import csv

# ── Fixed paths ───────────────────────────────────────────
TEMPLATE_PATH = r"C:\template.oft"
DATABASE_PATH = r"C:\db.csv"

# ── Defaults (user can override interactively) ────────────
DEFAULT_SEND_FROM      = ["id@outlook.com"]
DEFAULT_MAX_EMAILS     = 96    # per account
DEFAULT_BATCH_SIZE     = 7
DEFAULT_BATCH_INTERVAL     = 1500   # seconds (25 min)
DEFAULT_BATCH_INTERVAL_MAX = 2700   # seconds (45 min) — upper bound for randomisation
DELAY_IN_BATCH             = 5     # seconds between emails inside a batch (min)
DELAY_IN_BATCH_MAX         = 20     # seconds between emails inside a batch (max)
DRY_RUN                = False  # True = open drafts for review
DEFAULT_NAME_COLUMN    = "Contacted Person Name"
DEFAULT_EMAIL_COLUMN   = "Institution Email"
DEFAULT_STATUS_COLUMN  = "Status"


# ── Helper: prompt with a default ─────────────────────────
def _ask(prompt, default, cast=str):
    """Show *prompt* with [default]; return cast value or default on blank."""
    raw = input(f"{prompt} [{default}]: ").strip()
    if not raw:
        return default
    try:
        return cast(raw)
    except ValueError:
        print(f"  ⚠  Invalid input, using default ({default})")
        return default


def _ask_range(label, default_min, default_max, cast=int):
    """Prompt for a min/max range. Returns (min_val, max_val).
    If the user enters the same value for both, delays become fixed (no randomisation)."""
    print(f"\n  {label}")
    while True:
        lo = _ask(f"    Min (seconds)", default_min, cast)
        hi = _ask(f"    Max (seconds)", default_max, cast)
        if lo <= hi:
            if lo == hi:
                print(f"    → Fixed delay of {lo}s (no randomisation)")
            else:
                print(f"    → Will randomise between {lo}s – {hi}s")
            return lo, hi
        print(f"  ⚠  Min ({lo}) must be ≤ Max ({hi}). Please try again.")


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
        print("  ⚠  Column not found. Please enter one of the listed columns exactly.")


def main():
    print("═" * 55)
    print("  📧  Mailer — Interactive Setup")
    print("═" * 55)

    # ── 1. Send-from accounts ─────────────────────────────
    print(f"\nDefault send-from: {', '.join(DEFAULT_SEND_FROM)}")
    raw = input(
        "Enter send-from email(s) separated by commas\n"
        "(leave blank to keep defaults): "
    ).strip()

    if raw:
        send_from_list = [e.strip() for e in raw.split(",") if e.strip()]
    else:
        send_from_list = list(DEFAULT_SEND_FROM)

    print(f"  → Will round-robin across {len(send_from_list)} account(s):")
    for addr in send_from_list:
        print(f"      • {addr}")

    # ── 1b. Map sender names to each address ──────────────
    print("\nEnter the display name for each sending account:")
    sender_names = {}
    for addr in send_from_list:
        name = input(f"  Name for {addr}: ").strip()
        if not name:
            name = addr.split("@")[0].replace(".", " ").title()
            print(f"    (defaulting to '{name}')")
        sender_names[addr] = name

    # ── 2. Tunables ───────────────────────────────────────
    print()
    max_emails = _ask("Max emails PER ACCOUNT", DEFAULT_MAX_EMAILS, int)
    batch_size = _ask("Batch size", DEFAULT_BATCH_SIZE, int)

    delay_min, delay_max = _ask_range(
        "Delay between emails WITHIN a batch:",
        DELAY_IN_BATCH, DELAY_IN_BATCH_MAX,
    )
    interval_min, interval_max = _ask_range(
        "Delay between BATCHES:",
        DEFAULT_BATCH_INTERVAL, DEFAULT_BATCH_INTERVAL_MAX,
    )

    total_possible = max_emails * len(send_from_list)
    print(f"\n  → {max_emails} mails × {len(send_from_list)} account(s) = "
          f"{total_possible} total possible emails")

    # ── 3. Load recipients ────────────────────────────────
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
        print("Nothing to send — all rows already have a status.")
        return

    # ── 4. Preview first recipient ────────────────────────
    first = recipients[0]
    first_account = send_from_list[0]
    print(f"\n{'─' * 55}")
    print("  PREVIEW — first email that will be sent:")
    print(f"    From : {sender_names[first_account]} <{first_account}>")
    print(f"    To   : {first['email']}")
    print(f"    Name : {first['full_name']}")
    print(f"    Mode : {'DRY RUN (drafts)' if DRY_RUN else 'LIVE (sends immediately)'}")
    preview_first = _first_name(first['full_name'])
    short_name = len(preview_first) <= 3
    print(f"    Placeholders replaced:")
    if short_name:
        print(f"      Dear [First name],  → Greetings,  (name '{preview_first}' is ≤3 chars)")
    else:
        print(f"      [First name]        → {preview_first}")
    print(f"      [Institution Name]  → {first.get('institution_name', '')}")
    print(f"      [sender name]       → {sender_names[first_account]}")
    print(f"      [sender mail]       → {first_account}")
    print(f"{'─' * 55}")

    confirm = input("\nProceed? (y/n): ").strip().lower()
    if confirm not in ("y", "yes"):
        print("Aborted.")
        return

    # ── 5. Launch ─────────────────────────────────────────
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
    )


if __name__ == "__main__":
    main()
