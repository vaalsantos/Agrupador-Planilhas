import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QVBoxLayout, QWidget, QLabel, QSpacerItem, QSizePolicy
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QPainter, QFont
import pandas as pd

class SpreadsheetMergerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Agrupador")
        self.setFixedSize(900, 700)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Título no topo
        self.title_label = QLabel("Agrupador de Planilha", self)
        self.title_label.setFont(QFont("Open Sans"
                                       "", 22))  # Ajusta para um tamanho que se encaixe bem no topo
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setStyleSheet("color: #242C4C;")
        layout.addWidget(self.title_label)

        layout.addItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        button_style = """
            QPushButton {
                background-color: #242C4C;
                color: #67b3e4;
                font-size: 12px;  # Diminui a fonte para se adequar ao tamanho do botão
                padding: 5px;
                border-radius: 10px;
            }
            QPushButton:disabled {
                background-color: #AAAAAA;
                color: #FFFFFF;
            }
        """

        self.import_button = QPushButton("Importar Planilhas")
        self.import_button.setFont(QFont("Open Sans", 12))
        self.import_button.setStyleSheet(button_style)
        self.import_button.clicked.connect(self.import_spreadsheet)
        layout.addWidget(self.import_button, alignment=Qt.AlignCenter)

        self.save_button = QPushButton("Salvar Planilha Unificada")
        self.save_button.setFont(QFont("Open Sans", 12))
        self.save_button.setStyleSheet(button_style)
        self.save_button.clicked.connect(self.save_merged_spreadsheet)
        self.save_button.setEnabled(False)
        layout.addWidget(self.save_button, alignment=Qt.AlignCenter)

        layout.addItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        self.credit_label = QLabel("Valdemi Silva, Kuehne + Nagel, VCA")
        self.credit_label.setFont(QFont("Open Sans", 8))
        self.credit_label.setAlignment(Qt.AlignCenter)
        self.credit_label.setStyleSheet("color: #242C4C; background-color: transparent;")
        layout.addWidget(self.credit_label)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.drawPixmap(self.rect(), QPixmap('C:/Users/vasan/desktop/quebraplanilha/background_image.png'))

    def import_spreadsheet(self):
        options = QFileDialog.Options()
        files, _ = QFileDialog.getOpenFileNames(self, "Importar Planilhas", "",
                                                "Planilhas (*.xlsx *.xls *.csv);;Todos os Arquivos (*)",
                                                options=options)
        if files:
            dfs = []
            for file in files:
                try:
                    if file.endswith('.csv'):
                        df = pd.read_csv(file)
                    else:
                        df = pd.read_excel(file)
                    dfs.append(df)
                except Exception as e:
                    print(f"Erro ao importar planilha: {e}")
                    return
            if dfs:
                self.merged_df = pd.concat(dfs, ignore_index=True)
                self.save_button.setEnabled(True)
                self.save_button.setStyleSheet("background-color: #242C4C; color: #67b3e4;")
                print("Planilhas importadas.")

    def save_merged_spreadsheet(self):
        if self.merged_df is not None:
            options = QFileDialog.Options()
            fileName, _ = QFileDialog.getSaveFileName(self, "Salvar Planilha Unificada", "", "Planilhas Excel (*.xlsx);;Todos os Arquivos (*)", options=options)
            if fileName:
                try:
                    self.merged_df.to_excel(fileName, index=False)
                    print("Planilha salva com sucesso.")
                except Exception as e:
                    print(f"Erro ao salvar planilha: {e}")
        else:
            print("Não há dados para salvar.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = SpreadsheetMergerApp()
    window.show()
    sys.exit(app.exec_())
