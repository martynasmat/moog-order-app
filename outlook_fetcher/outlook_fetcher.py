"""Find Outlook conversations by PO in email subjects and bodies.

Each order includes a flat email list and the messages grouped by conversation.
Rows for the same PO receive the same conversations. The LINE value is retained
for classification.
"""
import re
import win32com.client

MAIL_FOLDER = 0  # Folder.DefaultItemType for mail folders
SKIP_FOLDERS = {"deleted items", "junk email", "junk e-mail", "sync issues", "conflicts",
                "local failures", "server failures", "outbox",
                "ištrinti elementai", "nepageidaujamas el. paštas"}
TABLE_COLUMNS = ("Subject", "ReceivedTime", "SenderEmailAddress", "MessageClass")
STORE_ENTRY_ID = "https://schemas.microsoft.com/mapi/proptag/0x0FFB0102"


def _norm_po(value):
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    po = re.sub(r"\s+", "", str(value)).upper()
    if po.startswith("PO#"):
        return "PO" + po[3:]
    return "PO" + po if po.isdecimal() else po


def _po_pattern(wanted_pos):
    alternatives = []
    for po in sorted(wanted_pos, key=lambda po: (-len(po), po)):
        if po.startswith("PO") and len(po) > 2:
            suffix = po[2:]
            prefix = r"PO\s*#?\s*"
            if suffix.isdecimal():
                prefix = rf"(?:{prefix})?"
            alternatives.append(prefix + re.escape(suffix))
        else:
            alternatives.append(re.escape(po))
    return re.compile(r"(?<![A-Z0-9])(?:" + "|".join(alternatives) + r")(?![A-Z0-9])", re.I)


# ---------------------------------------------------------------- Outlook access
def _walk(folder):
    yield folder
    try:
        for sub in folder.Folders:
            yield from _walk(sub)
    except Exception:
        return


def _mail_folders(namespace):
    for store in namespace.Stores:
        try:
            root = store.GetRootFolder()
        except Exception:
            continue  # store offline / no permission
        for folder in _walk(root):
            try:
                if folder.DefaultItemType == MAIL_FOLDER and \
                        folder.Name.strip().lower() not in SKIP_FOLDERS:
                    yield folder
            except Exception:
                continue


def _dasl(numbers):
    props = ('"urn:schemas:httpmail:subject"', '"urn:schemas:httpmail:textdescription"')
    escaped = [number.replace("'", "''") for number in numbers]
    return "@SQL=" + " OR ".join(f"{p} LIKE '%{n}%'" for n in escaped for p in props)


def _rows(folder, numbers):
    """Yield candidate mail rows for the subject and body search."""
    try:
        table = folder.GetTable(_dasl(numbers))
    except Exception:
        # Some Outlook stores cannot filter on the body. Check each message in those folders.
        try:
            table = folder.GetTable()
        except Exception:
            return
    if table is None:
        return
    try:
        for col in TABLE_COLUMNS:
            table.Columns.Add(col)
    except Exception:
        return
    while not table.EndOfTable:
        try:
            row = table.GetNextRow()
            if not str(row.Item("MessageClass") or "").startswith("IPM.Note"):
                continue
            yield (row.Item("EntryID"), row.Item("Subject") or "",
                   row.Item("ReceivedTime"), row.Item("SenderEmailAddress") or "")
        except Exception:
            continue


def list_mail_folders():
    """Diagnostics: which mailboxes/folders does COM see? -> [(store, path, count)]"""
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    out = []
    for f in _mail_folders(ns):
        try:
            out.append((f.Store.DisplayName, f.FolderPath, f.Items.Count))
        except Exception:
            pass
    return out


def _conversation_ids(item):
    members = set()
    try:
        conversation = item.GetConversation()
        if conversation is None:
            return members
        table = conversation.GetTable()
        table.Columns.Add(STORE_ENTRY_ID)
        while not table.EndOfTable:
            row = table.GetNextRow()
            try:
                if not str(row.Item("MessageClass") or "").startswith("IPM.Note"):
                    continue
                store_id = row.BinaryToString(STORE_ENTRY_ID)
                entry_id = row.Item("EntryID")
                if store_id and entry_id:
                    members.add((store_id, entry_id))
            except Exception:
                continue  # A bad row should not hide the rest of the conversation.
    except Exception:
        pass  # Keep the original match if Outlook cannot enumerate its conversation.
    return members


def _message_body(item):
    try:
        return item.Body or ""
    except Exception:
        return ""


def _email_data(message_id, subject, received, sender, body):
    store_id, entry_id = message_id
    try:
        received_at = received.isoformat()
    except (AttributeError, ValueError):
        received_at = str(received)
    return {"entry_id": entry_id, "store_id": store_id, "subject": subject,
            "sender": sender, "received_at": received_at, "body": body}


def _read_conversation_email(namespace, message_id):
    store_id, entry_id = message_id
    try:
        item = namespace.GetItemFromID(entry_id, store_id)
        return _email_data(message_id, item.Subject or "", item.ReceivedTime,
                           item.SenderEmailAddress or "", _message_body(item))
    except Exception:
        return None


def fetch_emails(orders):
    if not orders:
        return []
    if any(not str(o.get("PO") or "").strip() for o in orders):
        raise ValueError("A PO is required to fetch emails; LINE alone is not enough.")

    wanted_pos = {_norm_po(o["PO"]) for o in orders}
    po_pattern = _po_pattern(wanted_pos)
    numbers = sorted({po.removeprefix("PO") or po for po in wanted_pos})

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    emails = {}  # (store_id, entry_id) -> email dict
    seeds = {}  # Matching messages and the PO references they contain.
    conversations = {}  # Message ID -> sorted tuple of conversation member IDs.

    for folder in _mail_folders(ns):
        try:
            store_id = folder.StoreID
        except Exception:
            continue
        for entry_id, subject, received, sender in _rows(folder, numbers):
            message_id = (store_id, entry_id)
            if message_id in seeds:
                continue
            try:
                item = ns.GetItemFromID(entry_id, store_id)
            except Exception:
                item = None
            body = _message_body(item) if item is not None else ""
            found = {
                _norm_po(match.group())
                for text in (subject, body)
                for match in po_pattern.finditer(text)
            }
            if not found:
                continue
            seeds[message_id] = found
            emails[message_id] = _email_data(message_id, subject, received, sender, body)
            if item is not None and message_id not in conversations:
                members = _conversation_ids(item)
                if members:
                    members.add(message_id)
                    for member_id in tuple(members):
                        members.update(conversations.get(member_id, ()))
                    member_ids = tuple(sorted(members))
                    for member_id in member_ids:
                        conversations[member_id] = member_ids

    results = []
    for order in orders:
        po = _norm_po(order["PO"])
        matched_groups = {
            conversations.get(message_id, (message_id,))
            for message_id, found in seeds.items()
            if po in found
        }
        matched = {}
        chains = []
        for member_ids in sorted(matched_groups):
            chain = []
            for message_id in member_ids:
                if message_id not in emails:
                    emails[message_id] = _read_conversation_email(ns, message_id)
                email = emails[message_id]
                if email is not None:
                    chain.append(email)
                    matched[message_id] = email
            if chain:
                chain.sort(key=lambda email: email["received_at"])
                chains.append(chain)
        chains.sort(key=lambda chain: chain[0]["received_at"])
        ordered_emails = [matched[message_id] for message_id in sorted(matched)]
        ordered_emails.sort(key=lambda email: email["received_at"])
        results.append({**order, "emails": ordered_emails, "conversations": chains})
    return results


if __name__ == "__main__":
    # On the Moog PC:  python outlook_fetcher.py   -> lists what COM can see
    for store, path, count in list_mail_folders():
        print(f"{store} | {path} | {count}")
