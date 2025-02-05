from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QComboBox, QTableWidget, QTableWidgetItem, QTextEdit, QCheckBox, QHBoxLayout, QMessageBox, QLineEdit, QInputDialog, QTabWidget
)
from PyQt5.QtCore import QTimer, QTime, QDate
import sys
import pandas as pd
import os
from datetime import datetime

class AttendanceApp(QWidget):
    def __init__(self):
        super().__init__()
        self.manager_password = "admin123"
        self.leave_types =["عارضة", "طارئة", "سنوية", "منحة","مستشفى","ميدانية"]
        self.datasheet = ["الرقم العسكرى","الرتبة", "الموظف",  "الرتبة", "رصيد عارضة", "رصيد سنوية", "رصيد عارضة طارئة", "آخر تهيئة"]
        self.out_types = ["مؤتمر", "إلتزام", "مأمورية", "مركز"]
        self.filename = "attendance.xlsx"
        self.load_data()
        self.initUI()
    
    def load_data(self):
        if os.path.exists(self.filename):
            self.data = pd.read_excel(self.filename, sheet_name="الضباط", engine='openpyxl')
        else:
            self.data = pd.DataFrame(columns= self.datasheet)
            default_employees = [
                {"الرقم العسكرى": 1, "الرتبة": "مدير", "الموظف": "أحمد علي", "القسم": "الموارد البشرية", "رصيد عارضة": 7, "رصيد سنوية": 15, "رصيد عارضة طارئة": 2, "آخر تهيئة": datetime.now().strftime('%Y-%m-%d')}
            ]
            self.data = pd.DataFrame(default_employees)
            self.save_data()
    
    def save_data(self):
        self.data.to_excel(self.filename, sheet_name="الضباط", index=False, engine='openpyxl')
    
    def update_time(self):
        current_time = QTime.currentTime()
        self.time_label.setText(f"الوقت: {current_time.toString('hh:mm:ss')}")
        self.date_label.setText(f"التاريخ: {QDate.currentDate().toString('dd/MM/yyyy')}")
    
    def initUI(self):
        self.setWindowTitle("برنامج تسجيل حضور الضباط")
        self.setGeometry(100, 100, 800, 600)
        layout = QVBoxLayout()
        # عرض التاريخ والوقت
        self.date_label = QLabel("التاريخ: ")
        self.time_label = QLabel("الوقت: ")
        layout.addWidget(self.date_label)
        layout.addWidget(self.time_label)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
        # اختيار القسم
        self.department_label = QLabel("اختر القسم:")
        self.department_combo = QComboBox()
        self.department_combo.addItems(self.data['القسم'].unique().tolist())  # إضافة الأقسام من بيانات الموظفين
        self.department_combo.currentTextChanged.connect(self.check_manager_access)
        layout.addWidget(self.department_label)
        layout.addWidget(self.department_combo)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = AttendanceApp()
    ex.show()
    sys.exit(app.exec_())
