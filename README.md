# moog-order-app

Run the desktop app from the project folder:

```powershell
.\.venv\Scripts\python.exe -m pip install openpyxl pywin32 PySide6
.\.venv\Scripts\python.exe -m gui
```

## Local email classification

Install the local inference runtime in the app's Python environment. This command uses the Windows x64 CPU wheel.

```powershell
.\.venv\Scripts\python.exe -m pip install "https://github.com/abetlen/llama-cpp-python/releases/download/v0.3.36/llama_cpp_python-0.3.36-py3-none-win_amd64.whl"
```

Place an instruction-tuned GGUF model in the `models` folder. Load an Excel file, choose the model, and click **Classify emails**. The Outlook scan finds related conversations by PO in email subjects and bodies. The model reviews those conversations and returns a status for each requested PO and line. Open **View chain** to read the original emails. Use **Edit system prompt** to save updates to `system_prompt.txt`.

Run `build.bat` after installing the runtime to include classification in the Windows executable. Place the `models` folder beside `MoogOrders.exe`. Prompt edits are saved beside the executable.
