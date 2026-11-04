import re

import win32com.client


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
        results.append({**order, "emails": emails_by_po[po]})
    return results
