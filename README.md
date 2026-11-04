# moog-order-app

Run the desktop app from the project folder:

```powershell
.\.venv\Scripts\python.exe -m pip install openpyxl pywin32 PySide6
.\.venv\Scripts\python.exe -m gui
```

Open classic Outlook with your mailbox configured. Drag an `.xlsx` file onto the
upload box or select **Browse files**. Scanning starts automatically and runs in
the background. The table shows matching emails for each PO and optional line.
**View chain** opens the matched messages and their quoted replies in a reader;
**Open selected email in Outlook** opens the original message.

The scan searches the default Inbox by PO only, accepting both `PO123` and `PO 123`.
LINE is displayed but does not affect matching; rows with the same PO receive the same emails.
It does not fetch other folders or additional conversation messages outside that scope.
Use **Refresh** to check Outlook's connection again.
