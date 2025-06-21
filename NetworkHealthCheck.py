'''
--------------------------------------------------------------------------------
| No |  Date      | Version |      remarks
--------------------------------------------------------------------------------
| 1  | 21 June 25 | V1.00   |  Initial version
--------------------------------------------------------------------------------
'''
import sys
import os
import csv
import json
import threading
import subprocess
import platform
import webbrowser
from datetime import datetime
from collections import deque, defaultdict

import pandas as pd
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QHBoxLayout,
                             QLineEdit, QFileDialog, QSpinBox, QGridLayout)
from PyQt5.QtCore import QTimer, Qt, QMetaObject, Q_ARG, pyqtSlot
from PyQt5.QtGui import QCursor
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.dates as mdates


class PingMonitor(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Network Health Check Tool - v1.00")
        self.resize(1200, 800)

        self.max_sites = 6
        self.ip_inputs = []
        self.rtt_data = defaultdict(lambda: deque(maxlen=288))
        self.success_data = defaultdict(list)
        self.output_folder = os.getcwd()
        self.ping_interval = 10
        self.timeout = 3000
        self.retry_count = 3
        self.is_running = False

        self.graphs = {}
        self.stat_labels = {}

        self.setup_ui()
        self.timer = QTimer()
        self.timer.timeout.connect(self.run_all_pings)

    def setup_ui(self):
        layout = QVBoxLayout(self)
        config_layout = QGridLayout()

        for i in range(self.max_sites):
            ip_input = QLineEdit(self)
            ip_input.setPlaceholderText(f"Enter IP {i + 1}")
            config_layout.addWidget(ip_input, i, 0)
            self.ip_inputs.append(ip_input)

        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(10, 3600)
        self.interval_spin.setValue(self.ping_interval)
        self.interval_spin.setSuffix(" sec")
        config_layout.addWidget(QLabel("Ping Interval:"), 0, 1)
        config_layout.addWidget(self.interval_spin, 0, 2)

        self.timeout_spin = QSpinBox()
        self.timeout_spin.setRange(100, 5000)
        self.timeout_spin.setValue(self.timeout)
        self.timeout_spin.setSuffix(" ms")
        config_layout.addWidget(QLabel("Ping Timeout:"), 1, 1)
        config_layout.addWidget(self.timeout_spin, 1, 2)

        self.retry_spin = QSpinBox()
        self.retry_spin.setRange(1, 10)
        self.retry_spin.setValue(self.retry_count)
        config_layout.addWidget(QLabel("Retry Times:"), 2, 1)
        config_layout.addWidget(self.retry_spin, 2, 2)

        folder_btn = QPushButton("Select Output Folder")
        folder_btn.clicked.connect(self.select_folder)
        config_layout.addWidget(folder_btn, 3, 1, 1, 2)

        self.folder_label = QLabel(f"<a href='#'>Output Folder: {self.output_folder}</a>")
        self.folder_label.setOpenExternalLinks(False)
        self.folder_label.setTextInteractionFlags(Qt.TextBrowserInteraction)
        self.folder_label.setCursor(QCursor(Qt.PointingHandCursor))
        self.folder_label.linkActivated.connect(self.open_output_folder)
        config_layout.addWidget(self.folder_label, self.max_sites, 0, 1, 3)

        layout.addLayout(config_layout)

        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("Start Monitoring")
        self.start_btn.clicked.connect(self.start_monitoring)
        btn_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("Stop")
        self.stop_btn.clicked.connect(self.stop_monitoring)
        self.stop_btn.setEnabled(False)
        btn_layout.addWidget(self.stop_btn)

        self.export_btn = QPushButton("Export Summary")
        self.export_btn.clicked.connect(self.export_summary)
        btn_layout.addWidget(self.export_btn)

        self.reset_btn = QPushButton("Reset All Charts")
        self.reset_btn.clicked.connect(self.reset_all_data)
        btn_layout.addWidget(self.reset_btn)

        self.save_btn = QPushButton("Save Config")
        self.save_btn.clicked.connect(self.save_config)
        btn_layout.addWidget(self.save_btn)

        self.load_btn = QPushButton("Load Config")
        self.load_btn.clicked.connect(self.load_config)
        btn_layout.addWidget(self.load_btn)

        layout.addLayout(btn_layout)

        self.canvas_layout = QGridLayout()
        layout.addLayout(self.canvas_layout)

    def open_output_folder(self):
        webbrowser.open(f'file://{os.path.abspath(self.output_folder)}')

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.output_folder = folder
            self.folder_label.setText(f"<a href='#'>Output Folder: {self.output_folder}</a>")

    def save_config(self):
        config = [ip.text().strip() for ip in self.ip_inputs if ip.text().strip()]
        with open(os.path.join(self.output_folder, 'config.json'), 'w') as f:
            json.dump(config, f)
        print("Config saved.")

    def load_config(self):
        try:
            with open(os.path.join(self.output_folder, 'config.json'), 'r') as f:
                config = json.load(f)
            for i, ip in enumerate(config):
                if i < self.max_sites:
                    self.ip_inputs[i].setText(ip)
            print("Config loaded.")
        except Exception as e:
            print(f"Failed to load config: {e}")

    def start_monitoring(self):
        self.ping_interval = self.interval_spin.value()
        self.timeout = self.timeout_spin.value()
        self.retry_count = self.retry_spin.value()
        self.is_running = True
        self.timer.start(self.ping_interval * 1000)
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.run_all_pings()
        for ip_input in self.ip_inputs:
            ip = ip_input.text().strip()
            if ip:
                QMetaObject.invokeMethod(self, "update_graph_slot", Qt.QueuedConnection, Q_ARG(str, ip))

    def stop_monitoring(self):
        self.is_running = False
        self.timer.stop()
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

    def run_all_pings(self):
        if not self.is_running:
            return
        for ip_input in self.ip_inputs:
            ip = ip_input.text().strip()
            if ip:
                threading.Thread(target=self.ping_and_log, args=(ip,), daemon=True).start()

    def ping_and_log(self, ip):
        timestamp = datetime.now()
        date_str = timestamp.strftime("%Y%m%d")
        filename = os.path.join(self.output_folder, f"{ip.replace('.', '_')}_{date_str}.csv")
        success = 0
        rtt = None

        is_windows = platform.system().lower() == 'windows'
        timeout_arg = ['-w', str(self.timeout)] if is_windows else ['-W', str(int(self.timeout / 1000))]

        for attempt in range(self.retry_count):
            try:
                command = ['ping', '-n' if is_windows else '-c', '1'] + timeout_arg + [ip]
                result = subprocess.run(command,
                                        stdout=subprocess.PIPE,
                                        stderr=subprocess.PIPE,
                                        text=True,
                                        timeout=self.timeout / 1000 + 1)
                output = result.stdout
                if "time=" in output.lower() or "time<" in output.lower():
                    success = 1
                    rtt_str = output.lower().split("time=")[1].split()[0].replace("ms", "").replace("<", "").strip()
                    rtt = float(rtt_str)
                    break
            except Exception as e:
                continue

        self.rtt_data[ip].append((timestamp, rtt))
        self.success_data[ip].append(success)

        with open(filename, "a", newline='') as f:
            writer = csv.writer(f)
            if f.tell() == 0:
                writer.writerow(["Timestamp", "RTT (ms)", "Success"])
            writer.writerow([timestamp.strftime("%Y-%m-%d %H:%M:%S"), rtt if rtt is not None else "", success])

        QMetaObject.invokeMethod(self, "update_graph_slot", Qt.QueuedConnection, Q_ARG(str, ip))

    @pyqtSlot(str)
    def update_graph_slot(self, ip):
        if ip not in self.graphs:
            fig = Figure(figsize=(6, 2))
            canvas = FigureCanvas(fig)
            ax = fig.add_subplot(111)
            row = len(self.graphs) // 2
            col = len(self.graphs) % 2
            self.canvas_layout.addWidget(canvas, row, col)
            stat_label = QLabel("")
            self.canvas_layout.addWidget(stat_label, row + 1, col)
            self.graphs[ip] = (fig, canvas, ax)
            self.stat_labels[ip] = stat_label

        fig, canvas, ax = self.graphs[ip]
        ax.clear()
        filtered_data = [(t, r) for t, r in self.rtt_data[ip] if r is not None]
        if filtered_data:
            timestamps, rtts = zip(*filtered_data)
        else:
            timestamps, rtts = [], []

        if timestamps:
            ax.plot(timestamps, rtts, marker='o', label="RTT (ms)")
            ax.set_title(f"{ip} | Success Rate: {sum(self.success_data[ip]) / len(self.success_data[ip]) * 100:.1f}%")
            ax.set_xlabel("Time")
            ax.set_ylabel("RTT (ms)")
            ax.legend()
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
            ax.set_xlim(left=max(timestamps[-1] - pd.Timedelta(minutes=30), timestamps[0]), right=timestamps[-1])
            fig.autofmt_xdate()

            min_rtt = min(rtts)
            max_rtt = max(rtts)
            avg_rtt = sum(rtts) / len(rtts)
            self.stat_labels[ip].setText(f"Min: {min_rtt:.1f} ms | Max: {max_rtt:.1f} ms | Avg: {avg_rtt:.1f} ms")
        else:
            self.stat_labels[ip].setText("No data")

        canvas.draw()

    def reset_all_data(self):
        self.rtt_data.clear()
        self.success_data.clear()
        for ip in self.graphs:
            self.stat_labels[ip].setText("")
            fig, canvas, ax = self.graphs[ip]
            ax.clear()
            canvas.draw()

    def export_summary(self):
        summary_rows = []
        for ip in self.rtt_data:
            filtered_rtts = [rtt for _, rtt in self.rtt_data[ip] if rtt is not None]
            total_pings = len(self.success_data[ip])
            success_count = sum(self.success_data[ip])
            if filtered_rtts:
                min_rtt = min(filtered_rtts)
                max_rtt = max(filtered_rtts)
                avg_rtt = sum(filtered_rtts) / len(filtered_rtts)
            else:
                min_rtt = max_rtt = avg_rtt = None
            success_rate = (success_count / total_pings * 100) if total_pings > 0 else 0
            summary_rows.append([ip, min_rtt, max_rtt, avg_rtt, success_rate, total_pings])

        df = pd.DataFrame(summary_rows,
                          columns=["IP", "Min RTT (ms)", "Max RTT (ms)", "Avg RTT (ms)", "Success Rate (%)",
                                   "Total Pings"])
        filename = os.path.join(self.output_folder, f"Summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        try:
            df.to_excel(filename, index=False)
            print(f"Exported summary to {filename}")
        except Exception as e:
            print(f"Export failed: {e}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = PingMonitor()
    win.show()
    sys.exit(app.exec_())
