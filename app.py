import logging
import os
import pandas as pd
import sys
from src.python.script import main as script_main
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QLabel, QFileDialog, QComboBox, QLineEdit, QMessageBox, QLabel, QDialog, QFrame, QInputDialog
from PyQt5.QtGui import QMovie, QIcon, QPalette, QColor, QFont
from PyQt5.QtCore import QThread, pyqtSignal

script_thread = None

class ScriptThread(QThread):
    finished = pyqtSignal(int)

    def __init__(self, script_func, input_path, output_path, dir):
        super().__init__()
        self.script_func = script_func
        self.input_path = input_path
        self.output_path = output_path
        self.dir = dir

    def run(self):
        try:
            return_code = self.script_func(self, self.input_path, self.output_path, self.dir)
            self.finished.emit(return_code)
        except Exception as e:
            logging.error(f"Error while executing script: {str(e)}")
            self.finished.emit(1)
            
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.abspath(__file__)
        dir = os.path.dirname(base_path)
        
        self.setWindowTitle('Script Runner')
        self.setFixedWidth(700)
        self.setWindowIcon(QIcon(f'{dir}\\src\\assets\\logo.png'))
        self.layout = QVBoxLayout()
        
        # styling
        app.setStyle('fusion')
        palette = app.palette()
        palette.setColor(QPalette.Window, QColor("#09396C"))
        palette.setColor(QPalette.Button, QColor("#879EC3"))
        palette.setColor(QPalette.WindowText, QColor("#ffffff"))
        app.setFont(QFont("slab serif", 10, QFont.Bold))
        app.setPalette(palette)

        # information text at top of app
        with open(f'{dir}\\src\\assets\\info.html', 'r') as file:
            info_html_content = file.read()
        info_html = QLabel()
        info_html.setText(info_html_content)
        info_html.setWordWrap(True)
        self.layout.addWidget(info_html)

        self.layout.addSpacing(10)
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        self.layout.addWidget(line)
        self.layout.addSpacing(10)
        
        # i/o config
        self.show_input = True
        self.input_is_file = True
        self.show_output = True
        self.output_is_file = False

        button_font = QFont("slab serif", 9)
        if self.show_input:
            if self.input_is_file:
                # Input for Outstanding Charges
                input_label = QLabel('Input File')
                self.input_field = QLineEdit()
                input_button = QPushButton('Browse')
                input_button.setFont(button_font)
                input_button.clicked.connect(lambda: self.input_field.setText(QFileDialog.getOpenFileName()[0]))
                self.layout.addWidget(input_label)
                self.layout.addWidget(self.input_field)
                self.layout.addWidget(input_button)

            else:
                input_label = QLabel('Input Directory')
                input_field = QLineEdit()
                input_button = QPushButton('Browse')
                input_button.setFont(button_font)
                input_button.clicked.connect(lambda: input_field.setText(QFileDialog.getExistingDirectory()))
                self.layout.addWidget(input_label)
                self.layout.addWidget(input_field)
                self.layout.addWidget(input_button)

        if self.show_output:
            if self.output_is_file:
                output_label = QLabel('Output File')
                output_field = QLineEdit()
                output_button = QPushButton('Browse')
                output_button.setFont(button_font)
                output_button.clicked.connect(lambda: output_field.setText(QFileDialog.getOpenFileName()[0]))
                self.layout.addWidget(output_label)
                self.layout.addWidget(output_field)
                self.layout.addWidget(output_button)
            else:
                output_label = QLabel('Log File Output Directory')
                self.output_field = QLineEdit()
                output_button = QPushButton('Browse')
                output_button.setFont(button_font)
                output_button.clicked.connect(lambda: self.output_field.setText(QFileDialog.getExistingDirectory()))
                self.layout.addWidget(output_label)
                self.layout.addWidget(self.output_field)
                self.layout.addWidget(output_button)

        self.layout.addSpacing(10)
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        self.layout.addWidget(line)
        self.layout.addSpacing(10)

        self.run_button = QPushButton('Run')
        self.run_button.setFont(button_font)
        self.run_button.setStyleSheet("background-color: #21314d; color: #ffffff")

        exit_button = QPushButton('Exit')
        exit_button.setFont(button_font)
        exit_button.setStyleSheet("background-color: #cc4628; color: #ffffff")

        self.layout.addWidget(self.run_button)
        self.layout.addWidget(exit_button)

        self.loading_gif = QMovie(f'{dir}\\src\\assets\\pacman_loading.gif')
        self.loading_dialog = QDialog(self)
        self.loading_dialog.setWindowTitle("Script Running...")
        loading_label = QLabel(self.loading_dialog)
        loading_label.setMovie(self.loading_gif)
        self.loading_dialog.setLayout(QVBoxLayout())
        self.loading_dialog.layout().addWidget(loading_label)

        central_widget = QWidget()
        central_widget.setLayout(self.layout)
        self.setCentralWidget(central_widget)
        
        self.run_button.clicked.connect(self.run_script)
        exit_button.clicked.connect(lambda: self.close())
        
        self.script_thread = None

    def run_script(self):
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.abspath(__file__)
        dir = os.path.dirname(base_path)

        input_path, output_path = None, None
        if self.show_input:
            input_path = self.input_field.text()
        if self.show_output:
            output_path = self.output_field.text()
            
        if (self.show_input and not input_path) or (self.show_output and not output_path):
            QMessageBox.warning(self, "Warning", "Please provide all required fields.")
            return
        
        # Configure logging
        now = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
        logging.basicConfig(
            filename=os.path.join(output_path, f'{now}.log'),  # Log file name
            level=logging.DEBUG,    # Log level
            format='%(asctime)s - %(levelname)s - %(message)s',  # Log format
            datefmt='%Y-%m-%d %H:%M:%S'  # Date format
        )

        try:
            self.run_button.setEnabled(False)
            self.loading_dialog.show()
            self.loading_gif.start()
            logging.info(f'Running Script: {str(script_main)}')
            self.script_thread = ScriptThread(script_main, input_path, output_path, dir)
            self.script_thread.finished.connect(self.script_finished)
            self.script_thread.start()
        except Exception as e:
            logging.error(f'Error while attempting to run script: {str(e)}')
            
    def script_finished(self, return_code):
        self.loading_gif.stop()
        self.loading_dialog.close()
        self.run_button.setEnabled(True)

        if return_code == 0:
            QMessageBox.information(self, "Information", "Script ran successfully!")
        else:
            QMessageBox.critical(self, "Error", "Script finished running with errors")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec()