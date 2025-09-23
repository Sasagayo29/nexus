# Arquivo: ui/worker.py
from PySide6.QtCore import QThread, Signal

class Worker(QThread):
    progress = Signal(str)
    finished = Signal(bool)

    def __init__(self, file_path, target_function):
        super().__init__()
        self.file_path = file_path
        self.target_function = target_function # Armazena a função que deve ser executada

    def run(self):
        # Executa a função que foi passada durante a criação
        success = self.target_function(self.file_path, self.progress)
        self.finished.emit(success)