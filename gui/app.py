import math
import sys
from pathlib import Path

import pythoncom
import win32com.client
from PySide6.QtCore import QPointF, Qt, QThread, QTimer, Signal
from PySide6.QtGui import QFont, QPainter
from PySide6.QtWidgets import (
    QApplication, QDialog, QFileDialog, QFrame, QHBoxLayout, QHeaderView,
    QLabel, QListWidget, QMainWindow, QMessageBox, QPushButton, QSplitter,
    QTableWidget, QTableWidgetItem, QTextBrowser, QVBoxLayout, QWidget,
)

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
    def __init__(self, order, parent=None):
        super().__init__(parent)
        title = str(order["PO"])
        if order["LINE"] is not None:
            title += f" · LN{order['LINE']}"
        self.setWindowTitle(f"Email chain — {title}")
        self.resize(1050, 650)
        self.emails = order["emails"]
        layout = QVBoxLayout(self)
        heading = QLabel(title)
        layout.addWidget(heading)
        splitter = QSplitter()
        self.messages = QListWidget()
        self.messages.setMinimumWidth(260)
        for email in self.emails:
            self.messages.addItem(f"{email['subject']}\n{email['received_at'][:16].replace('T', ' ')}")
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
        if self.emails:
            self.messages.setCurrentRow(0)

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
            outlook.GetItemFromID(self.emails[index]["entry_id"]).Display()
        except Exception as error:
            QMessageBox.warning(self, "Could not open email", str(error))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Moog orders")
        self.resize(880, 640)
        self.setMinimumSize(700, 500)
        self.worker = None
        self.results = []
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
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["PO / LN", "Emails found", "Email chain"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(2, 160)
        self.table.verticalHeader().hide()
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        layout.addWidget(self.table, 1)
        QTimer.singleShot(0, self.start_task)

    def set_connected(self, connected):
        self.outlook_status.setText("Outlook connected" if connected else "Outlook not connected")
        self.outlook_status.setStyleSheet(
            "color: #15803d;" if connected else "color: #b91c1c;"
        )

    def start_task(self, path=None):
        if self.worker is not None:
            return
        self.drop.setEnabled(False)
        self.refresh.setEnabled(False)
        self.worker = OutlookTask(path, self)
        self.worker.connected.connect(self.set_connected)
        self.worker.results_ready.connect(self.show_results)
        self.worker.failed.connect(self.show_error)
        self.worker.finished.connect(self.finish_task)
        self.worker.finished.connect(self.worker.deleteLater)
        if path:
            self.results = []
            self.table.setRowCount(0)
            self.scanning.show()
            self.spinner.timer.start()
        self.worker.start()

    def load_file(self, path):
        if self.worker is not None:
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
        self.drop.setEnabled(True)
        self.refresh.setEnabled(True)

    def show_error(self, message):
        QMessageBox.warning(self, "Scan failed", message)

    def show_results(self, results):
        self.results = results
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
            button = QPushButton("View chain" if count else "No emails found")
            button.setEnabled(count > 0)
            button.clicked.connect(lambda checked=False, row=order: self.view_chain(row))
            self.table.setCellWidget(index, 2, button)
            self.table.setRowHeight(index, 32)

    def view_chain(self, order):
        EmailViewer(order, self).exec()

    def closeEvent(self, event):
        if self.worker is not None and self.worker.isRunning():
            QMessageBox.information(self, "Outlook is busy", "Please wait for the Outlook operation to finish before closing.")
            event.ignore()
        else:
            event.accept()


def main():
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 9))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
