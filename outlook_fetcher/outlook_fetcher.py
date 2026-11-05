"""Outlook (classic) email fetcher for Moog orders  --  GetTable version.

Drop-in replacement for outlook_fetcher/outlook_fetcher.py:
same function fetch_emails(orders), same input (order dicts with "PO" and optional "LINE"),
same output (each order dict + "emails": [{entry_id, subject, sender, received_at, body}]).

How it is faster than the original:
- each folder is searched by Outlook itself (DASL filter), not email-by-email in Python
- the matching emails come back as one Table (EntryID, Subject, ReceivedTime, Sender)
  in a single COM call instead of opening every email object
- the full body is loaded only for emails that actually mention a PO from the Excel file

What else changed vs. the original:
- searches ALL mail folders in ALL mailboxes in the profile (Inbox subfolders, Sent Items,
  archives, PSTs, shared mailboxes like supply.lt), not only the top-level Inbox
- LN-aware: "PO260378 LN3,6,9,12" goes to rows LN3/6/9/12 only, not to LN4;
  an email that mentions a PO without any LN goes to every row of that PO
- one broken/protected email does not crash the scan; duplicates are counted once
"""
import re
import win32com.client

MAIL_FOLDER = 0  # Folder.DefaultItemType for mail folders
SKIP_FOLDERS = {"deleted items", "junk email", "junk e-mail", "sync issues", "conflicts",
                "local failures", "server failures", "outbox",
                "ištrinti elementai", "nepageidaujamas el. paštas"}
TABLE_COLUMNS = ("Subject", "ReceivedTime", "SenderEmailAddress", "MessageClass")


# ---------------------------------------------------------------- PO / LN parsing
# a line list: numbers joined by "," "&" "+" "/" "and" "ir", or ranges "1-4" / "1 to 4".
# plain spaces are NOT separators, so "LN2 5000vnt" gives {2}, not {2, 5000}
_NUM_LIST = r"\d{1,3}(?:\s*(?:,|&|\+|/|-|–|\band\b|\bir\b|\bto\b)\s*\d{1,3})*"
PO_RE = re.compile(r"(?<![A-Z0-9])PO\s*#?\s*(\d{6})(?!\d)", re.I)
LN_RE = re.compile(r"^\s*[,;:/-]?\s*(?:LN|LINE|LINES|EIL\.?)\s*[:#]?\s*(" + _NUM_LIST + ")", re.I)
LN_EXTRA_RE = re.compile(r"^\s*(?:,|&|\+|\band\b|\bir\b)\s*(?:LN|LINE)\s*(" + _NUM_LIST + ")", re.I)


def _expand(num_list):
    out = set()
    for t in re.split(r"\s*(?:,|&|\+|/|\band\b|\bir\b)\s*", num_list, flags=re.I):
        rng = re.match(r"(\d+)\s*(?:-|–|\bto\b)\s*(\d+)", t, re.I)
        if rng:
            a, b = int(rng.group(1)), int(rng.group(2))
            if a <= b and b - a < 100:
                out.update(range(a, b + 1))
        elif t.strip().isdigit():
            out.add(int(t))
    return out


def parse_po_lines(text):
    """'PO260378 LN3,6,9,12 + PO260535' -> {'PO260378': {3,6,9,12}, 'PO260535': None}
    None = PO mentioned without lines, i.e. the email is about the whole PO."""
    lines_by_po = {}
    for m in PO_RE.finditer(text or ""):
        po = "PO" + m.group(1)
        rest = text[m.end():m.end() + 80]
        lines = set()
        ln = LN_RE.match(rest)
        while ln:  # also handles "LN1, LN2 & LN5"
            lines |= _expand(ln.group(1))
            rest = rest[ln.end():]
            ln = LN_EXTRA_RE.match(rest)
        lines_by_po.setdefault(po, set()).update(lines)
    return {po: (lines or None) for po, lines in lines_by_po.items()}


def email_matches(order_po, order_line, found):
    po = _norm_po(order_po)
    if po not in found:
        return False
    lines = found[po]
    if order_line in (None, "") or lines is None:
        return True
    try:
        return int(float(order_line)) in lines
    except (TypeError, ValueError):
        return True


def _norm_po(value):
    po = re.sub(r"\s+", "", str(value)).upper()
    return po if po.startswith("PO") else "PO" + po


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


def _dasl(numbers, with_body=True):
    props = ['"urn:schemas:httpmail:subject"']
    if with_body:
        props.append('"urn:schemas:httpmail:textdescription"')
    return "@SQL=" + " OR ".join(f"{p} LIKE '%{n}%'" for n in numbers for p in props)


def _rows(folder, numbers):
    """Yield (entry_id, subject, received, sender) for emails in folder that mention a number."""
    table = None
    for flt in (_dasl(numbers, True), _dasl(numbers, False)):  # body search may be unsupported
        try:
            table = folder.GetTable(flt)
            break
        except Exception:
            table = None
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


def fetch_emails(orders):
    if not orders:
        return []
    if any(not str(o.get("PO") or "").strip() for o in orders):
        raise ValueError("A PO is required to fetch emails; LINE alone is not enough.")

    wanted_pos = {_norm_po(o["PO"]) for o in orders}
    numbers = sorted(po[2:] for po in wanted_pos)

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    emails = {}  # entry_id -> (email dict, parsed {PO: lines})

    for folder in _mail_folders(ns):
        for entry_id, subject, received, sender in _rows(folder, numbers):
            if entry_id in emails:
                continue
            # cheap check on the subject first; open the body only when needed
            found = parse_po_lines(subject)
            try:
                body = ns.GetItemFromID(entry_id).Body or ""
            except Exception:
                body = ""
            for po, lines in parse_po_lines(body).items():
                if po not in found:
                    found[po] = lines
                elif found[po] is not None and lines:
                    found[po] |= lines
            found = {po: ln for po, ln in found.items() if po in wanted_pos}
            if not found:
                continue
            try:
                received_at = received.isoformat()
            except Exception:
                received_at = str(received)
            emails[entry_id] = ({"entry_id": entry_id, "subject": subject, "sender": sender,
                                 "received_at": received_at, "body": body}, found)

    results = []
    for order in orders:
        matched = [e for e, found in emails.values()
                   if email_matches(order["PO"], order.get("LINE"), found)]
        matched.sort(key=lambda e: e["received_at"])
        results.append({**order, "emails": matched})
    return results


if __name__ == "__main__":
    # On the Moog PC:  python outlook_fetcher.py   -> lists what COM can see
    for store, path, count in list_mail_folders():
        print(f"{store} | {path} | {count}")
