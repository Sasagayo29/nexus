# Arquivo: ui/main_window.py
# VERSÃO COM ÍCONE DA APLICAÇÃO

from PySide6.QtWidgets import QMainWindow, QPushButton, QVBoxLayout, QWidget, QTextEdit, QLabel, QFileDialog, QMessageBox, QHBoxLayout
from PySide6.QtGui import QIcon ### NOVA LINHA ###

# Importa a nova função que acabamos de criar
from core.atualizacao_massiva import run_update_process 
# Importa a função antiga
from core.processamento_rateio import run_full_process

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # --- Configurações da Janela ---
        self.setWindowTitle("NEXUS - Automação de Planilhas")
        self.setGeometry(100, 100, 600, 500)
        
        # --- DEFINIÇÃO DO ÍCONE DA APLICAÇÃO --- ### NOVA LINHA ###
        self.setWindowIcon(QIcon("assets/app_icon.png")) ### NOVA LINHA ###

        # Atributos
        self.file_path = None
        self.worker = None

        # Widgets
        self.label = QLabel("1. Selecione o arquivo 'Motorola - Planilha de Controle.xlsm'")
        self.select_button = QPushButton("Selecionar Arquivo")
        self.file_path_label = QLabel("Nenhum arquivo selecionado.")
        
        self.run_rateio_button = QPushButton("2. Gerar Rateio")
        self.run_update_button = QPushButton("3. Executar Atualização Massiva")
        
        self.log_box = QTextEdit()
        
        self.run_rateio_button.setEnabled(False)
        self.run_update_button.setEnabled(False)
        self.log_box.setReadOnly(True)

        # Layout
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.label)
        main_layout.addWidget(self.select_button)
        main_layout.addWidget(self.file_path_label)
        
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.run_rateio_button)
        button_layout.addWidget(self.run_update_button)
        
        main_layout.addLayout(button_layout)
        main_layout.addWidget(QLabel("Log de Processamento:"))
        main_layout.addWidget(self.log_box)

        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        # Conexões
        self.select_button.clicked.connect(self.select_file)
        self.run_rateio_button.clicked.connect(self.start_rateio_processing)
        self.run_update_button.clicked.connect(self.start_update_processing)

    def select_file(self):
        file_path_tuple = QFileDialog.getOpenFileName(self, "Selecionar Planilha", "", "Arquivos Excel (*.xlsm *.xlsx)")
        file_path = file_path_tuple[0]
        if file_path:
            self.file_path = file_path
            self.file_path_label.setText(f"Arquivo: {self.file_path.split('/')[-1]}")
            self.run_rateio_button.setEnabled(True)
            self.run_update_button.setEnabled(True)
            self.log_box.setText("Arquivo selecionado. Escolha uma ação.")

    def start_rateio_processing(self):
        self.start_processing(run_full_process)

    def start_update_processing(self):
        self.start_processing(run_update_process)

    def start_processing(self, target_function):
        from ui.worker import Worker
        if not self.file_path:
            QMessageBox.warning(self, "Aviso", "Por favor, selecione um arquivo primeiro.")
            return
            
        self._toggle_buttons(False)
        self.log_box.clear()
        
        self.worker = Worker(file_path=self.file_path, target_function=target_function)
        self.worker.progress.connect(self.update_log)
        self.worker.finished.connect(self.on_processing_finished)
        self.worker.start()

    def update_log(self, message):
        self.log_box.append(message)
        
    def _toggle_buttons(self, enabled):
        """Função auxiliar para habilitar/desabilitar botões."""
        self.select_button.setEnabled(enabled)
        self.run_rateio_button.setEnabled(enabled)
        self.run_update_button.setEnabled(enabled)

    def on_processing_finished(self, success):
        self._toggle_buttons(True)
        if success:
            QMessageBox.information(self, "Processo Concluído", "A operação foi finalizada com sucesso! Verifique o log e os arquivos gerados.")
        else:
            QMessageBox.critical(self, "Erro", "Ocorreu um erro durante o processamento. Verifique o log para mais detalhes.")