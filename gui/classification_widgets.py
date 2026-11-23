from threading import Event

from PySide6.QtCore import QIODevice, QSaveFile, QThread, Signal
from PySide6.QtWidgets import (
    QComboBox, QDialog, QDialogButtonBox, QMessageBox, QPlainTextEdit, QVBoxLayout,
)

from gui.classification import MODELS_DIR, PROMPT_PATH, classify_order, load_model


class ModelSelector(QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumContentsLength(24)
        self.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.setAccessibleName("GGUF model")
        self.refresh_models()

    def refresh_models(self):
        current = self.currentData()
        self.setToolTip(f"Place GGUF models in {MODELS_DIR}")
        try:
            MODELS_DIR.mkdir(exist_ok=True)
            paths = sorted(
                (path for path in MODELS_DIR.iterdir() if path.suffix.lower() == ".gguf" and path.is_file()),
                key=lambda path: path.name.casefold(),
            )
        except OSError as error:
            self.clear()
            self.addItem("Model folder unavailable")
            self.setToolTip(str(error))
            return

        available = [str(path) for path in paths]
        if paths and available == [self.itemData(index) for index in range(self.count())]:
            return
        if not paths and self.currentText() == "Add a GGUF to models/":
            return

        self.clear()
        for path in paths:
            self.addItem(path.name, str(path))
        if not paths:
            self.addItem("Add a GGUF to models/")
        selected = self.findData(current)
        if selected >= 0:
            self.setCurrentIndex(selected)

    def showPopup(self):
        self.refresh_models()
        super().showPopup()


class PromptEditor(QDialog):
    def __init__(self, prompt, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit system prompt")
        self.resize(760, 560)
        layout = QVBoxLayout(self)
        self.editor = QPlainTextEdit(prompt)
        self.editor.setAccessibleName("System prompt")
        layout.addWidget(self.editor)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.save_prompt)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def save_prompt(self):
        text = self.editor.toPlainText()
        if not text.strip():
            QMessageBox.warning(self, "System prompt is empty", "Enter a system prompt before saving.")
            return

        destination = QSaveFile(str(PROMPT_PATH))
        data = text.encode("utf-8")
        if (not destination.open(QIODevice.OpenModeFlag.WriteOnly)
                or destination.write(data) != len(data)
                or not destination.commit()):
            QMessageBox.warning(self, "Could not save system prompt", destination.errorString())
            return
        self.accept()


class ClassificationTask(QThread):
    progress = Signal(str)
    classified = Signal(int, object)
    failed = Signal(str)

    def __init__(self, model_path, prompt, jobs, parent=None):
        super().__init__(parent)
        self.model_path = model_path
        self.prompt = prompt
        self.jobs = jobs
        self.cancel_requested = Event()
        self.summary = ""

    def run(self):
        is_cancelled = self.cancel_requested.is_set
        completed = 0
        failures = 0
        try:
            if not is_cancelled():
                with load_model(self.model_path, is_cancelled) as model:
                    for index, (row, order) in enumerate(self.jobs, 1):
                        if is_cancelled():
                            break
                        self.progress.emit(f"Classifying order row {index} of {len(self.jobs)}…")
                        try:
                            outcome = classify_order(model, self.prompt, order)
                        except Exception as error:
                            outcome = str(error) or type(error).__name__
                        if is_cancelled():
                            break
                        if isinstance(outcome, str):
                            failures += 1
                        else:
                            completed += 1
                        self.classified.emit(row, outcome)
        except Exception as error:
            if not is_cancelled():
                self.failed.emit(str(error) or type(error).__name__)
                return

        if is_cancelled() and completed + failures < len(self.jobs):
            status = f"Classification cancelled. Classified {completed} of {len(self.jobs)} order rows."
        else:
            status = f"Classified {completed} of {len(self.jobs)} order rows with {self.model_path.name}."
        if failures:
            status += f" {failures} failed."
        self.summary = status
