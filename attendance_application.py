from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QComboBox, QTableWidget, 
    QTableWidgetItem, QTextEdit, QCheckBox, QHBoxLayout, QMessageBox, QInputDialog
)
from PyQt5.QtCore import QTimer, QTime, QDate
import sys
import pandas as pd
import os
import openpyxl
from datetime import datetime
import msoffcrypto
from io import BytesIO

class AttendanceApp(QWidget):
    def __init__(self):
        super().__init__()
        self.filename = "attendance.xlsx"
        self.leave_types = ["عارضة", "طارئة", "سنوية", "منحة", "مستشفى", "ميدانية"]
        self.out_types = ["مؤتمر", "إلتزام", "مأمورية", "مركز"]
        self.load_data()
        self.initUI()
        self.password="admin123"
    
    def load_data(self):
        if os.path.exists(self.filename):
            self.data = pd.read_excel(self.filename, sheet_name="الضباط", engine='openpyxl')
        else:
            self.data = pd.DataFrame(columns=["الرقم العسكرى", "الرتبة", "الضابط", "القسم", "رصيد عارضة", "رصيد سنوية"])
            self.data.to_excel(self.filename, sheet_name="الضباط", index=False, engine='openpyxl')
    
    def initUI(self):
        self.setWindowTitle("برنامج تسجيل حضور الضباط")
        self.setGeometry(100, 100, 900, 600)
        layout = QVBoxLayout()
        
        self.date_label = QLabel(f"التاريخ: {QDate.currentDate().toString('dd/MM/yyyy')}")
        self.time_label = QLabel(f"الوقت: {QTime.currentTime().toString('hh:mm:ss')}")
        layout.addWidget(self.date_label)
        layout.addWidget(self.time_label)
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
        
        self.department_label = QLabel("اختر القسم:")
        self.department_combo = QComboBox()
        
        self.department_combo.addItems(self.data['القسم'].dropna().unique().tolist())
        self.department_combo.currentTextChanged.connect(self.update_employee_table)
        layout.addWidget(self.department_label)
        layout.addWidget(self.department_combo)
        
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels(["الرقم العسكرى", "الرتبة", "الضابط", "حاضر", "غائب", "الإجازات", "الخوارج", "ملاحظات"])
        layout.addWidget(self.table)
        
        self.save_button = QPushButton("حفظ التمام")
        self.save_button.clicked.connect(self.save_attendance)
        layout.addWidget(self.save_button)

        # زر إعادة تهيئة الإجازات
        self.reset_leave_button = QPushButton("إعادة تهيئة الإجازات")
        self.reset_leave_button.clicked.connect(self.ask_for_password)
        layout.addWidget(self.reset_leave_button)
        
        self.setLayout(layout)
    
    def update_time(self):
        current_time = QTime.currentTime()
        self.time_label.setText(f"الوقت: {current_time.toString('hh:mm:ss')}")
        self.date_label.setText(f"التاريخ: {QDate.currentDate().toString('dd/MM/yyyy')}")
    
    def update_employee_table(self):
        selected_department = self.department_combo.currentText()
        employees = self.data[self.data['القسم'] == selected_department]
        self.table.setRowCount(len(employees))
        
        # Resize columns to fit contents before populating with data
        self.table.horizontalHeader().setStretchLastSection(True)
        
        for row, (_, emp) in enumerate(employees.iterrows()):
            self.table.setItem(row, 0, QTableWidgetItem(str(emp['الرقم العسكرى'])))
            self.table.setItem(row, 1, QTableWidgetItem(emp['الرتبة']))
            self.table.setItem(row, 2, QTableWidgetItem(emp['الضابط']))
            
            present_btn = QPushButton("✔")
            absent_btn = QPushButton("❌")
            present_btn.clicked.connect(lambda _, r=row: self.mark_attendance(r, "حاضر"))
            absent_btn.clicked.connect(lambda _, r=row: self.mark_attendance(r, "غائب"))
            
            self.table.setCellWidget(row, 3, present_btn)
            self.table.setCellWidget(row, 4, absent_btn)
            
            leave_layout = QHBoxLayout()
            leave_widget = QWidget()
            leave_checkboxes = [QCheckBox(lt) for lt in self.leave_types]
            for cb in leave_checkboxes:
                leave_layout.addWidget(cb)
            leave_widget.setLayout(leave_layout)
            
            out_layout = QHBoxLayout()
            out_widget = QWidget()
            out_checkboxes = [QCheckBox(ot) for ot in self.out_types]
            for cb in out_checkboxes:
                out_layout.addWidget(cb)
            out_widget.setLayout(out_layout)
            
            self.table.setCellWidget(row, 5, leave_widget)
            self.table.setCellWidget(row, 6, out_widget)
            self.table.setCellWidget(row, 7, QTextEdit())
            
            self.toggle_leave_options(row, False)
            self.table.setRowHeight(row, 50)

        # Resize columns that do not have custom widgets
        self.table.setColumnWidth(2,250)
        self.table.setColumnWidth(5,550)
        self.table.setColumnWidth(6,400)
        
        
        # Adjust widths manually for columns with custom widgets (buttons, checkboxes, text editor)
        #self.resize_column_widths()

    def resize_column_widths(self):
        # Calculate maximum size of custom widgets and adjust column widths accordingly
        max_leave_width = max(self.get_widget_width(self.table.cellWidget(row, 5)) for row in range(self.table.rowCount()))
        max_out_width = max(self.get_widget_width(self.table.cellWidget(row, 6)) for row in range(self.table.rowCount()))
        # Set column widths based on widget content
        self.table.setColumnWidth(5, max_leave_width + 80)  # Column 5 (Leave checkboxes)
        self.table.setColumnWidth(6, max_out_width + 80)  # Column 6 (External commitments checkboxes)

    def get_widget_width(self, widget):
        # Calculate the maximum width of a widget (e.g., button or checkbox)
        if isinstance(widget, QPushButton):
            return widget.sizeHint().width()
        elif isinstance(widget, QCheckBox):
            return widget.sizeHint().width()
        elif isinstance(widget, QTextEdit):
            return widget.sizeHint().width()
        else:
            return 0

    def mark_attendance(self, row, status):
        self.toggle_leave_options(row, status == "غائب")
        self.table.item(row, 2).setData(100, status)
        if status == "حاضر":
            # إزالة أي تحديد سابق للإجازات في حالة الحضور
            for cb in self.table.cellWidget(row, 5).findChildren(QCheckBox):
                cb.setChecked(False)
    
    def toggle_leave_options(self, row, show):
        self.table.cellWidget(row, 5).setVisible(show)
        self.table.cellWidget(row, 6).setVisible(show)

    def save_attendance(self):
        selected_department = self.department_combo.currentText()
        current_date = QDate.currentDate().toString('yyyy-MM-dd')
        attendance_record = []
        current_time = QTime.currentTime()
        ent_time  = QTime(12,0,0)
        if current_time > ent_time:
            QMessageBox.warning(self, "تحذير", "تم تجاوز الوقت المسموح بتسجيل الحضور!")
            return

        for row in range(self.table.rowCount()):
            emp_name = self.table.item(row, 2).text()
            status = self.table.item(row, 2).data(100) if self.table.item(row, 2).data(100) == "حاضر" else "غائب"
            leave_types = [cb.text() for cb in self.table.cellWidget(row, 5).findChildren(QCheckBox) if cb.isChecked()]
            out_types = [cb.text() for cb in self.table.cellWidget(row, 6).findChildren(QCheckBox) if cb.isChecked()]
            notes = self.table.cellWidget(row, 7).toPlainText()
            
            # Load data from the Excel file to update the balance
            officers_data = pd.read_excel(self.filename, sheet_name="الضباط", engine='openpyxl')

            # Find the row for the employee
            emp_row = officers_data[officers_data["الضابط"] == emp_name].index
            if not emp_row.empty:
                emp_index = emp_row[0]
                
                # Calculate leave deductions
                casual_balance = officers_data.at[emp_index, "رصيد عارضة"]
                annual_balance = officers_data.at[emp_index, "رصيد سنوية"]
                
                # Prevent saving if casual balance is zero but leave is selected
                for leave in leave_types:
                    if leave == "عارضة" and casual_balance <= 0:
                        QMessageBox.warning(self, "رصيد غير كافٍ", f"لا يمكن تسجيل إجازة عارضة لـ {emp_name} لأن رصيده صفر!")
                        return  # Stop saving
                    elif leave == "طارئة" and casual_balance < 2:
                        QMessageBox.warning(self, "رصيد غير كافٍ", f"لا يمكن تسجيل إجازة طارئة لـ {emp_name} لأن رصيد العارضة أقل من 2!")
                        return  # Stop saving
                    elif leave == "سنوية" and annual_balance <= 0:
                        QMessageBox.warning(self, "رصيد غير كافٍ", f"لا يمكن تسجيل إجازة سنوية لـ {emp_name} لأن رصيده صفر!")
                        return  # Stop saving
                for leave in leave_types:
                    if leave == "عارضة" and casual_balance > 0:
                        casual_balance -= 1
                    elif leave == "طارئة" and casual_balance >= 2:
                        casual_balance -= 2  # Emergency leave counts as 2 casual leaves
                    elif leave == "سنوية" and annual_balance > 0:
                        annual_balance -= 1

                
                # Update the balance in the file
                officers_data.at[emp_index, "رصيد عارضة"] = max(0, casual_balance)
                officers_data.at[emp_index, "رصيد سنوية"] = max(0, annual_balance)

                # Save updates to the 'officers' sheet
                if os.path.exists(self.filename):
                    wb = openpyxl.load_workbook(self.filename)
                    if "الضباط" in wb.sheetnames:
                        # Do not delete the sheet, just update it
                        with pd.ExcelWriter(self.filename, engine='openpyxl', mode='w') as writer:
                            officers_data.to_excel(writer, sheet_name="الضباط", index=False)
                
                attendance_record.append([emp_name, status, ", ".join(leave_types), ", ".join(out_types), notes])

        # Create the attendance DataFrame
        df = pd.DataFrame(attendance_record, columns=["الضابط", "الحالة", "نوع الإجازة", "الخوارج", "الملاحظات"])
        sheet_name = f"{selected_department}-{current_date}"
        
        # Save attendance in the main file
        with pd.ExcelWriter(self.filename, engine='openpyxl', mode='a', if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        fname= sheet_name + ".xlsx"
        # Save the same attendance sheet in a new backup file
        with pd.ExcelWriter(fname, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        QMessageBox.information(self, "تم الحفظ", "تم حفظ بيانات الحضور بنجاح!")

    def reset_leave_balances(self):
        """إعادة تهيئة رصيد الإجازات لكل الموظفين بعد 6 شهور."""
        current_date = datetime.now()
        employees = self.data['الضابط'].unique()
        # Reset the leave balances for each employee
        for emp in employees:
            #self.leave_balances[emp] = {'رصيد عارضة': 7, 'رصيد سنوية': 15}  # Reset balances
            
            # Update the DataFrame with the new leave balances
            self.data.loc[self.data['الضابط'] == emp, 'رصيد عارضة'] = 7
            self.data.loc[self.data['الضابط'] == emp, 'رصيد سنوية'] = 15
            self.data.loc[self.data['الضابط'] == emp, 'آخر تهيئة'] = current_date.strftime('%Y-%m-%d')
        
        # Save the updated data back to the Excel file
        if os.path.exists(self.filename):
            with pd.ExcelWriter(self.filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                self.data.to_excel(writer, sheet_name="الضباط", index=False)

        QMessageBox.information(self, "تم التهيئة", "تمت إعادة تهيئة رصيد الإجازات بنجاح!")
    
    def ask_for_password(self):
        password, ok = QInputDialog.getText(self, "إدخال كلمة المرور", "الرجاء إدخال كلمة المرور:")
        if ok and password == self.password:
            self.reset_leave_balances()
        else:
            QMessageBox.warning(self, "خطأ", "كلمة المرور غير صحيحة!")
    
if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = AttendanceApp()
    ex.show()
    sys.exit(app.exec_())
