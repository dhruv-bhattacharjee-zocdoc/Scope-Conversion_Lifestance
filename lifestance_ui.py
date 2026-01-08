import sys
import subprocess
import webbrowser
import json
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QFileDialog, QTextEdit, QVBoxLayout,
    QWidget, QLabel, QHBoxLayout, QGroupBox, QGridLayout
)
from PyQt5.QtGui import QFont, QPalette, QColor
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal
import os

class ProcessWorker(QThread):
    log_line = pyqtSignal(str)
    finished = pyqtSignal()
    def __init__(self, exe, args):
        super().__init__()
        self.exe = exe
        self.args = args
        self._process = None
    def run(self):
        import subprocess
        self._process = subprocess.Popen(
            [self.exe] + self.args,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8"
        )
        for line in self._process.stdout:
            self.log_line.emit(line.rstrip())
        self._process.wait()
        self.finished.emit()
    def stop(self):
        if self._process is not None:
            self._process.terminate()
            self._process = None

class LifestanceTranspositionTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Lifestance Transposition Tool")
        self.setGeometry(100, 100, 800, 460)
        self.setStyleSheet("background-color: #faf7fc;")
        font = QFont("Segoe UI", 11)
        self.setFont(font)
        self.process_thread = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_timer)
        self.seconds = 0

        # Title Label
        title = QLabel("Lifestance Transposition Tool")
        title.setFont(QFont("Segoe UI", 16, QFont.Bold))
        title.setStyleSheet("margin: 10px 0 0 8px; color: #333333;")

        # Actions group
        actions_group = QGroupBox()
        actions_group.setTitle("")  # No external title
        actions_group.setStyleSheet('''
            QGroupBox {
                background: #fff;
                border: 1px solid #e0e0e0;
                border-radius: 13px;
                padding: 18px 12px 12px 12px;
                margin-top: 8px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            }
        ''')
        actions_layout = QVBoxLayout()
        actions_title = QLabel("Actions")
        actions_title.setStyleSheet("font-size: 13px; color: #7680A5; font-weight: regular; margin-bottom: 5px;")
        actions_layout.addWidget(actions_title, 0, Qt.AlignLeft)
        buttons_layout = QHBoxLayout()
        self.select_button = QPushButton("\U0001F4C1  Select Input File")
        self.select_button.setMinimumWidth(180)
        self.select_button.setStyleSheet('''
            QPushButton {
                background-color: #111827;
                color: #fff;
                font-weight: bold;
                border-radius: 8px;
                padding: 10px 28px;
                font-size: 15px;
                border: none;
            }
            QPushButton:hover {
                background-color: #23263b;
            }
        ''')
        self.select_button.clicked.connect(self.select_file_and_run)
        self.stop_button = QPushButton("Stop")
        self.stop_button.setEnabled(False)
        self.stop_button.setStyleSheet('''
            QPushButton {
                background: #e63946;
                color: #fff;
                font-weight: bold;
                border-radius: 8px;
                padding: 10px 28px;
                font-size: 15px;
                border: none;
                margin-left: 10px;
            }
            QPushButton:hover { background: #b2182b; }
            QPushButton:disabled { background: #f3cccc; color: #fff; }
        ''')
        self.stop_button.clicked.connect(self.stop_process)
        self.gsheet_btn = QPushButton("\u2B73  Open Gsheets")
        self.gsheet_btn.setStyleSheet('''
            QPushButton {
                background-color: #f4f4f4;
                color: #2d3a53;
                font-weight: 400;
                border-radius: 8px;
                padding: 10px 24px;
                font-size: 15px;
                border: none;
                margin-left: 10px;
                box-shadow: 0 1px 4px rgba(22,29,37,0.08);
            }
            QPushButton:hover {
                background-color: #e6e6e6;
            }
            QPushButton:disabled { color: #bbb; }
        ''')
        self.gsheet_btn.setEnabled(False)
        self.gsheet_btn.clicked.connect(self.open_gsheets_and_reveal_output)
        buttons_layout.addWidget(self.select_button)
        buttons_layout.addWidget(self.stop_button)
        buttons_layout.addWidget(self.gsheet_btn)
        actions_layout.addLayout(buttons_layout)
        actions_group.setLayout(actions_layout)

        # Timer group
        timer_group = QGroupBox()
        timer_group.setTitle("")
        timer_group.setStyleSheet('''
            QGroupBox {
                background: #fff;
                border: 1px solid #e0e0e0;
                border-radius: 13px;
                padding: 20px 12px 14px 12px;
                margin-top: 8px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            }
        ''')
        timer_layout = QVBoxLayout()
        timer_title = QLabel("Timer")
        timer_title.setStyleSheet("font-size: 13px; color: #7680A5; font-weight: regular; margin-bottom: 7px;")
        timer_layout.addWidget(timer_title, 0, Qt.AlignLeft)
        self.timer_label = QLabel("00:00")
        self.timer_label.setFont(QFont("Segoe UI", 18, QFont.Medium))
        self.timer_label.setStyleSheet("color: #23263b; margin-bottom: 2px;")
        timer_layout.addWidget(self.timer_label, 0, Qt.AlignHCenter)
        timer_group.setLayout(timer_layout)
        timer_group.setMaximumWidth(210)

        # --- Add user info panel ---
        # Read credentials.json
        self.user_display = QLabel()
        self.role_display = QLabel()
        self.user_display.setAlignment(Qt.AlignRight | Qt.AlignTop)
        self.role_display.setAlignment(Qt.AlignRight | Qt.AlignTop)
        credentials_path = os.path.join(os.path.dirname(__file__), 'credentials.json')
        user_name_camel = "Unknown"
        user_role = "Unknown"
        try:
            with open(credentials_path, 'r') as f:
                creds = json.load(f)
                user_email = creds.get('user', '')
                user_role = creds.get('role', '')
                if user_email:
                    base = user_email.split('@')[0]
                    parts = base.replace('.', ' ').replace('_', ' ').split()
                    user_name_camel = ' '.join(w.capitalize() for w in parts)
        except Exception:
            pass
        self.user_display.setText(f"User: {user_name_camel}")
        self.role_display.setText(f"Role: {user_role}")
        self.user_display.setStyleSheet("font-weight: bold; color: #3A425D; font-size: 11px; margin-bottom:0px; margin-top:0px;")
        self.role_display.setStyleSheet("color: #78809d; font-size: 10px; margin-top:0px; margin-bottom:0px;")
        self.cred_hint = QLabel("To change the User and Role, change in credentials.json")
        self.cred_hint.setAlignment(Qt.AlignRight)
        self.cred_hint.setStyleSheet("font-size: 9px; color: #a0a4b8; font-style: italic; margin:0px;")

        # Input/Output fields
        inout_group = QGridLayout()
        input_box = QGroupBox()
        input_box.setTitle("")
        input_box.setStyleSheet('''
            QGroupBox {
                background: #fff;
                border: 1px solid #e0e0e0;
                border-radius: 13px;
                padding: 14px 12px 10px 12px;
                margin-top: 8px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            }
        ''')
        input_vbox = QVBoxLayout()
        input_title = QLabel("Input")
        input_title.setStyleSheet("font-size: 13px; color: #7680A5; font-weight: regular; margin-bottom: 5px;")
        input_vbox.addWidget(input_title, 0, Qt.AlignLeft)
        self.input_label = QLabel("No folder selected")
        self.input_label.setStyleSheet("color: #555; background: #fff; padding: 8px; border-radius: 5px;")
        input_vbox.addWidget(self.input_label)
        input_box.setLayout(input_vbox)

        io_arrow = QLabel("\u2194")  # ↔
        io_arrow.setFont(QFont("Segoe UI", 26, QFont.Bold))
        io_arrow.setAlignment(Qt.AlignCenter)
        io_arrow.setStyleSheet("margin: 16px 6px;")

        output_box = QGroupBox()
        output_box.setTitle("")
        output_box.setStyleSheet('''
            QGroupBox {
                background: #fff;
                border: 1px solid #e0e0e0;
                border-radius: 13px;
                padding: 14px 12px 10px 12px;
                margin-top: 8px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            }
        ''')
        output_vbox = QVBoxLayout()
        output_title = QLabel("Output")
        output_title.setStyleSheet("font-size: 13px; color: #7680A5; font-weight: regular; margin-bottom: 5px;")
        output_vbox.addWidget(output_title, 0, Qt.AlignLeft)
        self.output_label = QLabel("Awaiting input")
        self.output_label.setStyleSheet("color: #555; background: #fff; padding: 8px; border-radius: 5px;")
        output_vbox.addWidget(self.output_label)
        output_box.setLayout(output_vbox)

        inout_group.addWidget(input_box, 0, 0)
        inout_group.addWidget(io_arrow, 0, 1)
        inout_group.addWidget(output_box, 0, 2)

        # Console Log area
        console_group = QGroupBox()
        console_group.setTitle("")
        console_group.setStyleSheet('''
            QGroupBox {
                background: #fff;
                border: 1px solid #e0e0e0;
                border-radius: 13px;
                padding: 14px 12px 5px 12px;
                margin-top: 8px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            }
        ''')
        console_layout = QVBoxLayout()
        console_title = QLabel("Console Log")
        console_title.setStyleSheet("font-size: 13px; color: #7680A5; font-weight: regular; margin-bottom: 5px;")
        console_layout.addWidget(console_title, 0, Qt.AlignLeft)
        self.console = QTextEdit("Ready. Select an input folder to begin…")
        self.console.setReadOnly(True)
        self.console.setStyleSheet("background: #202840; color: #cde5fd; font-family: 'Consolas'; padding: 8px; border-radius: 6px;")
        console_layout.addWidget(self.console)
        console_group.setLayout(console_layout)

        # --- Title and user info on same row ---
        title_user_row = QHBoxLayout()
        title_label = QLabel("Lifestance Transposition Tool")
        title_label.setFont(QFont("Segoe UI", 16, QFont.Bold))
        title_label.setStyleSheet("margin: 10px 0 0 8px; color: #333333;")
        title_user_row.addWidget(title_label, alignment=Qt.AlignVCenter | Qt.AlignLeft)

        userinfo_layout = QVBoxLayout()
        userinfo_layout.addWidget(self.user_display, alignment=Qt.AlignRight)
        userinfo_layout.addWidget(self.role_display, alignment=Qt.AlignRight)
        userinfo_layout.addWidget(self.cred_hint, alignment=Qt.AlignRight)
        userinfo_layout.addStretch()
        title_user_row.addLayout(userinfo_layout)
        title_user_row.setStretch(0, 5)
        title_user_row.setStretch(1, 2)

        main_layout = QVBoxLayout()
        main_layout.addLayout(title_user_row)
        # ... rest of the layouts (top_row, timer, inout_group, etc) follow as before ...
        top_row = QHBoxLayout()
        top_row.addWidget(actions_group)
        top_row.addStretch()
        top_right = QVBoxLayout()
        top_right.addWidget(timer_group, alignment=Qt.AlignRight)
        top_right.addSpacing(6)
        top_row.addLayout(top_right)
        main_layout.addLayout(top_row)
        main_layout.addLayout(inout_group)
        main_layout.addWidget(console_group)

        wrapper = QWidget()
        wrapper.setLayout(main_layout)
        self.setCentralWidget(wrapper)

    def update_timer(self):
        self.seconds += 1
        mins, secs = divmod(self.seconds, 60)
        self.timer_label.setText(f"{mins:02d}:{secs:02d}")

    def select_file_and_run(self):
        import os
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Input File", "", "Excel Files (*.xlsx *.xls)", options=options
        )
        if file_path:
            self.input_label.setText(os.path.basename(file_path))
            self.console.append(f"Selected file: {file_path}\nRunning _main_1.py ...")
            self.seconds = 0
            self.timer.start(1000)  # Start timer (ticks every 1 sec)
            self.output_label.setText("Running ...")
            self.select_button.setEnabled(False)
            self.stop_button.setEnabled(True)
            self.process_thread = ProcessWorker(sys.executable, ["_main_1.py", file_path])
            self.process_thread.log_line.connect(self.append_log)
            self.process_thread.finished.connect(self.process_finished)
            self.process_thread.start()

    def stop_process(self):
        if self.process_thread is not None:
            self.process_thread.stop()
            self.stop_button.setEnabled(False)
            self.select_button.setEnabled(True)
            self.console.append("\nProcess interrupted by user.")
            self.output_label.setText("Stopped")

    def append_log(self, line):
        self.console.append(line)

    def process_finished(self):
        self.timer.stop()
        self.output_label.setText("Completed")
        self.stop_button.setEnabled(False)
        self.select_button.setEnabled(True)
        self.gsheet_btn.setEnabled(True)

    def open_gsheets_and_reveal_output(self):
        try:
            from Report import main as report_main
            report_main()
        except Exception as e:
            print(f"Failed to run report: {e}")
        import webbrowser
        import os
        import sys
        import subprocess
        webbrowser.open_new_tab('https://sheets.new')
        merged_output = os.path.join(os.path.dirname(__file__), 'Excel Files', 'Mergedoutput.xlsx')
        folder = os.path.dirname(merged_output)
        if os.path.exists(folder):
            if sys.platform.startswith('win'):
                os.startfile(os.path.normpath(folder))
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', folder])
            else:
                subprocess.Popen(['xdg-open', folder])

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = LifestanceTranspositionTool()
    window.show()
    sys.exit(app.exec_())
