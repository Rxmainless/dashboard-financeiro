import sys
import pandas as pd
import matplotlib.pyplot as plt
import mplcursors
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QPushButton, QWidget,
    QComboBox, QLabel, QScrollArea, QFileDialog, QTabWidget, QGridLayout, QSizePolicy
)
from PyQt5.QtCore import Qt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Dashboard Interativo")
        self.setGeometry(100, 100, 1000, 800)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setAlignment(Qt.AlignCenter)

        self.tab_widget = QTabWidget()
        self.main_layout.addWidget(self.tab_widget)

        self.dashboard_tab = QWidget()
        self.reports_tab = QWidget()
        self.settings_tab = QWidget()

        self.tab_widget.addTab(self.dashboard_tab, "Dashboard")
        self.tab_widget.addTab(self.reports_tab, "Relatórios")
        self.tab_widget.addTab(self.settings_tab, "Configurações")

        self.setup_dashboard_tab()
        self.setup_reports_tab()
        self.setup_settings_tab()

    def setup_dashboard_tab(self):
        self.dashboard_layout = QVBoxLayout(self.dashboard_tab)
        self.dashboard_layout.setAlignment(Qt.AlignCenter)

        # Layout de filtros
        self.filter_layout = QGridLayout()
        self.filter_layout.setAlignment(Qt.AlignCenter)
        self.status_filter = QComboBox()
        self.status_filter.addItems(["Todos", "Ativo", "Inativo"])
        self.product_filter = QComboBox()
        self.product_filter.addItems(["Todos", "Produto A", "Produto B", "Produto C"])
        self.filter_button = QPushButton("Filtrar")
        self.filter_button.clicked.connect(self.load_data)

        self.filter_layout.addWidget(QLabel("Status:"), 0, 0)
        self.filter_layout.addWidget(self.status_filter, 0, 1)
        self.filter_layout.addWidget(QLabel("Produto:"), 1, 0)
        self.filter_layout.addWidget(self.product_filter, 1, 1)
        self.filter_layout.addWidget(self.filter_button, 2, 0, 1, 2)
        self.dashboard_layout.addLayout(self.filter_layout)

        # Botões de carregamento e exportação
        self.button_load = QPushButton("Carregar Planilha")
        self.button_load.clicked.connect(self.load_file)
        self.dashboard_layout.addWidget(self.button_load, alignment=Qt.AlignCenter)

        self.button_export = QPushButton("Exportar Dados")
        self.button_export.clicked.connect(self.export_data)
        self.dashboard_layout.addWidget(self.button_export, alignment=Qt.AlignCenter)

        # Área de scroll para gráficos
        self.scroll_area = QScrollArea()
        self.scroll_area_widget = QWidget()
        self.scroll_area.setWidget(self.scroll_area_widget)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_layout = QVBoxLayout(self.scroll_area_widget)
        self.dashboard_layout.addWidget(self.scroll_area)

        # Configuração do gráfico
        self.figure, self.ax = plt.subplots()
        self.canvas = FigureCanvas(self.figure)
        self.canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.scroll_layout.addWidget(self.canvas)

        self.data = pd.DataFrame()  # Inicialmente vazio

    def setup_reports_tab(self):
        self.reports_layout = QVBoxLayout(self.reports_tab)
        self.reports_layout.setAlignment(Qt.AlignCenter)
        self.reports_layout.addWidget(QLabel("Relatórios"))

        self.report_area = QScrollArea()
        self.report_area_widget = QWidget()
        self.report_area.setWidget(self.report_area_widget)
        self.report_area.setWidgetResizable(True)
        self.report_layout = QVBoxLayout(self.report_area_widget)
        self.reports_layout.addWidget(self.report_area)

    def setup_settings_tab(self):
        self.settings_layout = QVBoxLayout(self.settings_tab)
        self.settings_layout.setAlignment(Qt.AlignCenter)
        self.settings_layout.addWidget(QLabel("Configurações"))

        self.save_settings_button = QPushButton("Salvar Configurações de Filtragem")
        self.save_settings_button.clicked.connect(self.save_settings)
        self.load_settings_button = QPushButton("Carregar Configurações de Filtragem")
        self.load_settings_button.clicked.connect(self.load_settings)

        self.settings_layout.addWidget(self.save_settings_button, alignment=Qt.AlignCenter)
        self.settings_layout.addWidget(self.load_settings_button, alignment=Qt.AlignCenter)

    def load_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Carregar Planilha", "", "CSV Files (*.csv);;Excel Files (*.xlsx; *.xls)", options=options)
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    self.data = pd.read_csv(file_path)
                else:
                    self.data = pd.read_excel(file_path)
                self.load_data()
            except Exception as e:
                print(f"Erro ao carregar arquivo: {e}")

    def load_data(self):
        if self.data.empty:
            return

        status = self.status_filter.currentText()
        product = self.product_filter.currentText()

        df = self.data.copy()

        if status != "Todos":
            df = df[df["Status"] == status]
        if product != "Todos":
            df = df[df["Produto"] == product]

        self.ax.clear()
        bars = df.plot(kind='bar', x='Vigência', y=['Cancelamentos Titulares', 'Cancelamentos Dependentes'], ax=self.ax)
        self.ax.set_title('Cancelamentos por Vigência')
        self.ax.set_xlabel('Vigência')
        self.ax.set_ylabel('Cancelamentos')
        self.canvas.draw()

        # Adiciona interatividade com mplcursors
        cursor = mplcursors.cursor(bars, hover=True)
        cursor.connect("add", lambda sel: sel.annotation.set_text(f'{sel.target[1]:.0f}'))

    def export_data(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "Salvar Dados", "", "CSV Files (*.csv);;Excel Files (*.xlsx);;PNG Files (*.png);;JPEG Files (*.jpg);;PDF Files (*.pdf)", options=options)
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    self.data.to_csv(file_path, index=False)
                elif file_path.endswith('.xlsx') or file_path.endswith('.xls'):
                    self.data.to_excel(file_path, index=False)
                elif file_path.endswith('.png') or file_path.endswith('.jpg') or file_path.endswith('.pdf'):
                    self.figure.savefig(file_path)
            except Exception as e:
                print(f"Erro ao exportar arquivo: {e}")

    def save_settings(self):
        settings = {
            "status": self.status_filter.currentText(),
            "product": self.product_filter.currentText()
        }
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "Salvar Configurações", "", "JSON Files (*.json)", options=options)
        if file_path:
            try:
                with open(file_path, 'w') as f:
                    pd.Series(settings).to_json(f)
            except Exception as e:
                print(f"Erro ao salvar configurações: {e}")

    def load_settings(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Carregar Configurações", "", "JSON Files (*.json)", options=options)
        if file_path:
            try:
                settings = pd.read_json(file_path, typ='series')
                self.status_filter.setCurrentText(settings['status'])
                self.product_filter.setCurrentText(settings['product'])
                self.load_data()
            except Exception as e:
                print(f"Erro ao carregar configurações: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
