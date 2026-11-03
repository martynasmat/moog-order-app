import re

import win32com.client


def _prioritize_emails(emails, po, line):
    if line is None:
        return emails
    po_pattern = r"PO\s*" + re.escape(po[2:]) if po.startswith("PO") else re.escape(po)
    pattern = re.compile(
        r"(?<![A-Z0-9])" + po_pattern + r"\s+(?:LN|LINE)\s*([A-Z0-9]+)",
        re.IGNORECASE,
    )
    line = str(line).strip().upper()

    def priority(email):
        subject_lines = [value.upper() for value in pattern.findall(email["subject"])]
        body_lines = [value.upper() for value in pattern.findall(email["body"])]
        if line in subject_lines:
            return 0
        if line in body_lines:
            return 1
        return 3 if subject_lines or body_lines else 2

    ranked = [(priority(email), email) for email in emails]
    return [email for rank, email in sorted(ranked, key=lambda item: item[0]) if rank < 3]


def fetch_emails(orders):
    if not orders:
        return []
    if any(not order["PO"] for order in orders):
        raise ValueError("A PO is required to fetch emails; LINE alone is not enough.")

    emails_by_po = {re.sub(r"\s+", "", str(order["PO"])).upper(): [] for order in orders}
    pattern = re.compile(
        r"(?<![A-Z0-9])(" + "|".join(
            r"PO\s*" + re.escape(po[2:]) if po.startswith("PO") else re.escape(po)
            for po in emails_by_po
        )
        + r")(?![A-Z0-9])", re.IGNORECASE,
    )

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
    for item in inbox.Items:
        if item.Class != 43:  # 43 = MailItem
            continue
        subject = item.Subject or ""
        body = item.Body or ""
        matched_pos = {re.sub(r"\s+", "", po).upper()
                       for po in pattern.findall(subject + "\n" + body)}
        if matched_pos:
            email = {
                "entry_id": item.EntryID,
                "subject": subject,
                "sender": item.SenderEmailAddress,
                "received_at": item.ReceivedTime.isoformat(),
                "body": body,
            }
            for po in matched_pos:
                emails_by_po[po].append(email)

    results = []
    for order in orders:
        po = re.sub(r"\s+", "", str(order["PO"])).upper()
        emails = _prioritize_emails(emails_by_po[po], po, order["LINE"])
        results.append({**order, "emails": emails})
    return results
