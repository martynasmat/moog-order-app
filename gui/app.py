import math
import sys
from pathlib import Path

import pythoncom
import win32com.client
from PySide6.QtCore import QPointF, Qt, QThread, QTimer, Signal
from PySide6.QtGui import QFont, QPainter
from PySide6.QtWidgets import (
    QApplication, QComboBox, QDialog, QFileDialog, QFrame, QHBoxLayout, QHeaderView,
    QLabel, QListWidget, QMainWindow, QMessageBox, QPushButton, QSplitter,
    QTableWidget, QTableWidgetItem, QTextBrowser, QVBoxLayout, QWidget,
)

from gui.classification import Classification, STATUS_LABELS, load_system_prompt, order_conversations
from gui.classification_widgets import ClassificationTask, ModelSelector, PromptEditor
from outlook_fetcher import fetch_emails
from xlsx_parser import parse_xlsx


class OutlookTask(QThread):
    connected = Signal(bool)
    results_ready = Signal(list)
    failed = Signal(str)

    def __init__(self, path=None, parent=None):
        super().__init__(parent)
        self.path = path

    def run(self):
        pythoncom.CoInitialize()
        try:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
                outlook.GetDefaultFolder(6).Name
                self.connected.emit(True)
            except Exception as error:
                self.connected.emit(False)
                if self.path:
                    raise RuntimeError("Cannot connect to Outlook. Open classic Outlook and try again.") from error
                return
            if self.path:
                orders = parse_xlsx(self.path)
                if not orders:
                    raise ValueError("No PO/LINE rows were found in this file.")
                self.results_ready.emit(fetch_emails(orders))
        except Exception as error:
            self.failed.emit(str(error))
        finally:
            pythoncom.CoUninitialize()


class Spinner(QWidget):
    def __init__(self):
        super().__init__()
        self.setFixedSize(28, 28)
        self.angle = 0
        self.timer = QTimer(self)
        self.timer.setInterval(70)
        self.timer.timeout.connect(self.advance)

    def advance(self):
        self.angle = (self.angle + 30) % 360
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(Qt.PenStyle.NoPen)
        for index in range(12):
            angle = math.radians(self.angle + index * 30)
            color = self.palette().text().color()
            color.setAlpha(40 + index * 19)
            painter.setBrush(color)
            painter.drawEllipse(QPointF(14 + 9 * math.cos(angle), 14 + 9 * math.sin(angle)), 2, 2)


class FileDrop(QFrame):
    file_selected = Signal(str)

    def __init__(self):
        super().__init__()
        self.setObjectName("fileDrop")
        self.setStyleSheet("QFrame#fileDrop { border: 1px dashed palette(mid); }")
        self.setAcceptDrops(True)
        self.setMinimumHeight(110)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        title = QLabel("Drop Excel file here")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        self.filename = QLabel(".xlsx files")
        self.filename.setTextFormat(Qt.TextFormat.PlainText)
        self.filename.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.filename.setWordWrap(True)
        layout.addWidget(self.filename)
        browse = QPushButton("Browse…")
        browse.setFixedWidth(100)
        browse.clicked.connect(self.browse)
        layout.addWidget(browse, alignment=Qt.AlignmentFlag.AlignHCenter)

    def browse(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select an Excel file", "", "Excel files (*.xlsx)")
        if path:
            self.file_selected.emit(path)

    def dragEnterEvent(self, event):
        urls = event.mimeData().urls()
        if (self.isEnabled() and len(urls) == 1 and urls[0].isLocalFile()
                and Path(urls[0].toLocalFile()).suffix.lower() == ".xlsx"):
            event.acceptProposedAction()

    def dropEvent(self, event):
        if self.isEnabled():
            self.file_selected.emit(event.mimeData().urls()[0].toLocalFile())
            event.acceptProposedAction()


class EmailViewer(QDialog):
    def __init__(self, order, parent=None, *, classification=None):
        super().__init__(parent)
        title = str(order["PO"])
        if order["LINE"] is not None:
            title += f" · LN{order['LINE']}"
        self.setWindowTitle(f"Email chain — {title}")
        self.resize(1050, 650)
        self.chains = order_conversations(order)
        self.emails = []
        layout = QVBoxLayout(self)
        heading = QLabel(title)
        layout.addWidget(heading)

        self.classification = QLabel()
        self.classification.setTextFormat(Qt.TextFormat.PlainText)
        self.classification.setWordWrap(True)
        self.classification.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.show_classification(classification)
        layout.addWidget(self.classification)

        if len(self.chains) > 1:
            chain_selector = QComboBox()
            chain_selector.setAccessibleName("Outlook conversation")
            for index, chain in enumerate(self.chains, 1):
                chain_selector.addItem(f"Conversation {index} · {chain[0]['subject']}")
            chain_selector.currentIndexChanged.connect(self.select_conversation)
            layout.addWidget(chain_selector)

        splitter = QSplitter()
        self.messages = QListWidget()
        self.messages.setMinimumWidth(260)
        self.body = QTextBrowser()
        self.body.setOpenExternalLinks(False)
        splitter.addWidget(self.messages)
        splitter.addWidget(self.body)
        splitter.setSizes([320, 690])
        layout.addWidget(splitter, 1)
        buttons = QHBoxLayout()
        open_outlook = QPushButton("Open selected email in Outlook")
        open_outlook.clicked.connect(self.open_in_outlook)
        buttons.addWidget(open_outlook)
        buttons.addStretch()
        close = QPushButton("Close")
        close.clicked.connect(self.accept)
        buttons.addWidget(close)
        layout.addLayout(buttons)
        self.messages.currentRowChanged.connect(self.show_email)
        if self.chains:
            self.select_conversation(0)

    def select_conversation(self, index):
        self.emails = self.chains[index]
        self.messages.clear()
        self.body.clear()
        for email in self.emails:
            self.messages.addItem(f"{email['subject']}\n{email['received_at'][:16].replace('T', ' ')}")
        if self.emails:
            self.messages.setCurrentRow(0)

    def show_classification(self, outcome):
        if isinstance(outcome, Classification):
            self.classification.setText(STATUS_LABELS[outcome.status])
        elif outcome is not None:
            self.classification.setText(f"Classification failed\n{outcome}")
        self.classification.setVisible(outcome is not None)

    def show_email(self, index):
        if index < 0:
            return
        email = self.emails[index]
        self.body.setPlainText(
            f"{email['subject']}\nFrom: {email['sender']}\n"
            f"Received: {email['received_at']}\n\n{email['body']}"
        )

    def open_in_outlook(self):
        index = self.messages.currentRow()
        if index < 0:
            return
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            email = self.emails[index]
            outlook.GetItemFromID(email["entry_id"], email["store_id"]).Display()
        except Exception as error:
            QMessageBox.warning(self, "Could not open email", str(error))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Moog orders")
        self.resize(880, 640)
        self.setMinimumSize(700, 500)
        self.worker = None
        self.classification_worker = None
        self.results = []
        self.classifications = {}
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        top = QHBoxLayout()
        self.outlook_status = QLabel("Checking Outlook…")
        top.addWidget(self.outlook_status)
        top.addStretch()
        self.refresh = QPushButton("Refresh")
        self.refresh.clicked.connect(lambda: self.start_task())
        top.addWidget(self.refresh)
        layout.addLayout(top)

        self.drop = FileDrop()
        self.drop.file_selected.connect(self.load_file)
        layout.addWidget(self.drop)

        classification_controls = QHBoxLayout()
        model_label = QLabel("Model")
        self.model_selector = ModelSelector()
        model_label.setBuddy(self.model_selector)
        self.edit_prompt_button = QPushButton("Edit system prompt")
        self.edit_prompt_button.clicked.connect(self.edit_prompt)
        self.classify_button = QPushButton("Classify emails")
        self.classify_button.clicked.connect(self.start_classification)
        self.cancel_classification_button = QPushButton("Cancel")
        self.cancel_classification_button.clicked.connect(self.cancel_classification)
        classification_controls.addWidget(model_label)
        classification_controls.addWidget(self.model_selector, 1)
        classification_controls.addWidget(self.edit_prompt_button)
        classification_controls.addWidget(self.classify_button)
        classification_controls.addWidget(self.cancel_classification_button)
        layout.addLayout(classification_controls)
        self.classification_status = QLabel()
        self.classification_status.setTextFormat(Qt.TextFormat.PlainText)
        self.classification_status.setWordWrap(True)
        self.classification_status.hide()
        layout.addWidget(self.classification_status)

        self.scanning = QWidget()
        scan_layout = QHBoxLayout(self.scanning)
        scan_layout.setContentsMargins(0, 0, 0, 0)
        self.spinner = Spinner()
        scan_layout.addWidget(self.spinner)
        scan_layout.addWidget(QLabel("Scanning emails…"))
        scan_layout.addStretch()
        self.scanning.hide()
        layout.addWidget(self.scanning)

        heading = QLabel("Results")
        layout.addWidget(heading)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["PO / LN", "Emails found", "Classification", "Email chain"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(3, 160)
        self.table.verticalHeader().hide()
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        layout.addWidget(self.table, 1)
        self.model_selector.currentIndexChanged.connect(self.update_controls)
        self.update_controls()
        QTimer.singleShot(0, self.start_task)

    def set_connected(self, connected):
        self.outlook_status.setText("Outlook connected" if connected else "Outlook not connected")
        self.outlook_status.setStyleSheet(
            "color: #15803d;" if connected else "color: #b91c1c;"
        )

    def start_task(self, path=None):
        if self.worker is not None or self.classification_worker is not None:
            return
        self.worker = OutlookTask(path, self)
        self.worker.connected.connect(self.set_connected)
        self.worker.results_ready.connect(self.show_results)
        self.worker.failed.connect(self.show_error)
        self.worker.finished.connect(self.finish_task)
        self.worker.finished.connect(self.worker.deleteLater)
        if path:
            self.results = []
            self.classifications = {}
            self.classification_status.hide()
            self.table.setRowCount(0)
            self.scanning.show()
            self.spinner.timer.start()
        self.update_controls()
        self.worker.start()

    def load_file(self, path):
        if self.worker is not None or self.classification_worker is not None:
            return
        if Path(path).suffix.lower() != ".xlsx" or not Path(path).is_file():
            QMessageBox.warning(self, "Choose an Excel file", "Please select an existing .xlsx file.")
            return
        self.drop.filename.setText(Path(path).name)
        self.start_task(path)

    def finish_task(self):
        self.worker = None
        self.spinner.timer.stop()
        self.scanning.hide()
        self.update_controls()

    def update_controls(self):
        task = self.classification_worker
        busy = self.worker is not None or task is not None
        self.drop.setEnabled(not busy)
        self.refresh.setEnabled(not busy)
        self.model_selector.setEnabled(not busy)
        self.edit_prompt_button.setEnabled(not busy)
        self.classify_button.setEnabled(
            not busy and bool(self.model_selector.currentData()) and any(row["emails"] for row in self.results)
        )
        self.cancel_classification_button.setVisible(task is not None)
        self.cancel_classification_button.setEnabled(task is not None and not task.cancel_requested.is_set())

    def edit_prompt(self):
        try:
            prompt = load_system_prompt()
        except (OSError, UnicodeError) as error:
            QMessageBox.warning(self, "Could not open system prompt", str(error))
            return
        PromptEditor(prompt, self).exec()

    def start_classification(self):
        if self.worker is not None or self.classification_worker is not None:
            return
        model_path = self.model_selector.currentData()
        jobs = [(row, order) for row, order in enumerate(self.results) if order["emails"]]
        if not model_path or not jobs:
            return
        try:
            prompt = load_system_prompt()
        except (OSError, UnicodeError) as error:
            QMessageBox.warning(self, "Could not open system prompt", str(error))
            return
        if not prompt:
            QMessageBox.warning(self, "System prompt is empty", "Use Edit system prompt before classifying emails.")
            return

        self.classification_worker = ClassificationTask(Path(model_path), prompt, jobs, self)
        self.classification_worker.progress.connect(self.show_classification_progress)
        self.classification_worker.classified.connect(self.show_classification)
        self.classification_worker.failed.connect(self.classification_failed)
        self.classification_worker.finished.connect(self.finish_classification)
        self.classification_worker.finished.connect(self.classification_worker.deleteLater)
        self.classifications = {}
        for row in range(self.table.rowCount()):
            self.table.item(row, 2).setText("—")
        self.show_classification_status(f"Loading {Path(model_path).name}…")
        self.update_controls()
        self.classification_worker.start()

    def cancel_classification(self):
        task = self.classification_worker
        if task is not None and task.isRunning():
            task.cancel_requested.set()
            self.update_controls()
            self.show_classification_status("Cancelling classification…")

    def show_classification_progress(self, status):
        task = self.classification_worker
        if task is not None and not task.cancel_requested.is_set():
            self.show_classification_status(status)

    def show_classification_status(self, status):
        self.classification_status.setText(status)
        self.classification_status.show()

    def show_classification(self, row, outcome):
        self.classifications[row] = outcome
        item = self.table.item(row, 2)
        item.setText(STATUS_LABELS[outcome.status] if isinstance(outcome, Classification) else "Failed")
        item.setToolTip("" if isinstance(outcome, Classification) else outcome)

    def classification_failed(self, message):
        self.show_classification_status("Classification stopped.")
        QMessageBox.warning(self, "Classification failed", message)

    def finish_classification(self):
        if self.classification_worker.summary:
            self.show_classification_status(self.classification_worker.summary)
        self.classification_worker = None
        self.update_controls()

    def show_error(self, message):
        QMessageBox.warning(self, "Scan failed", message)

    def show_results(self, results):
        self.results = results
        self.classifications = {}
        self.table.setRowCount(len(results))
        for index, order in enumerate(results):
            name = str(order["PO"])
            if order["LINE"] is not None:
                name += f"  ·  LN{order['LINE']}"
            self.table.setItem(index, 0, QTableWidgetItem(name))
            count = len(order["emails"])
            count_item = QTableWidgetItem(str(count))
            count_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(index, 1, count_item)
            self.table.setItem(index, 2, QTableWidgetItem("—"))
            button = QPushButton("View chain" if count else "No emails found")
            button.setEnabled(count > 0)
            button.clicked.connect(lambda checked=False, row=index: self.view_chain(row))
            self.table.setCellWidget(index, 3, button)
            self.table.setRowHeight(index, 32)
        self.update_controls()

    def view_chain(self, row):
        EmailViewer(self.results[row], self, classification=self.classifications.get(row)).exec()

    def closeEvent(self, event):
        if self.worker is not None or self.classification_worker is not None:
            QMessageBox.information(self, "App is busy", "Please wait for the current operation to finish before closing.")
            event.ignore()
        else:
            event.accept()


def main():
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 9))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
