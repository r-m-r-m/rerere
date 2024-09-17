import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, 
                             QComboBox, QMessageBox, QLineEdit, QHBoxLayout, QDialog, 
                             QDialogButtonBox, QTableView, QAbstractItemView)
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QFont, QStandardItemModel, QStandardItem, QIntValidator

class AttendanceApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("誤配管理システム")

        layout = QVBoxLayout()
        self.font = QFont("Arial", 12)

        # 日付選択セクション
        date_layout = QHBoxLayout()
        
        self.year_combobox = QComboBox()
        self.month_combobox = QComboBox()
        self.day_combobox = QComboBox()

        current_date = QDate.currentDate()
        current_year = current_date.year()
        current_month = current_date.month()
        current_day = current_date.day()

        self.year_combobox.setFont(self.font)
        self.year_combobox.addItems([str(year) for year in range(current_year - 5, 2025)])
        self.year_combobox.setCurrentText(str(current_year))

        self.month_combobox.setFont(self.font)
        self.month_combobox.addItems([str(month).zfill(2) for month in range(1, 13)])
        self.month_combobox.setCurrentText(str(current_month).zfill(2))

        self.day_combobox.setFont(self.font)
        self.day_combobox.addItems([str(day).zfill(2) for day in range(1, 32)])
        self.day_combobox.setCurrentText(str(current_day).zfill(2))

        date_layout.addWidget(QLabel("年:", font=self.font))
        date_layout.addWidget(self.year_combobox)
        date_layout.addWidget(QLabel("月:", font=self.font))
        date_layout.addWidget(self.month_combobox)
        date_layout.addWidget(QLabel("日:", font=self.font))
        date_layout.addWidget(self.day_combobox)

        layout.addLayout(date_layout)

        # 出勤人数入力セクション
        self.attendance_label = QLabel("出勤人数:", font=self.font)
        layout.addWidget(self.attendance_label)

        self.attendance_input = QLineEdit()
        self.attendance_input.setFont(self.font)
        self.attendance_input.setPlaceholderText("出勤人数を入力")
        self.attendance_input.setValidator(QIntValidator(1, 999))
        layout.addWidget(self.attendance_input)

        # 次へボタン
        self.next_button = QPushButton("次へ")
        self.next_button.setFont(self.font)
        self.next_button.clicked.connect(self.check_file_existence)
        layout.addWidget(self.next_button)

        self.setLayout(layout)
        self.adjustSize()  # 解像度自動調整

    def check_file_existence(self):
        # 入力バリデーション
        if not self.attendance_input.text().isdigit():
            QMessageBox.critical(self, "入力エラー", "出勤人数は正の整数で入力してください。")
            return

        year = self.year_combobox.currentText()
        month = self.month_combobox.currentText()
        day = self.day_combobox.currentText()
        date = f"{year}_{month}_{day}"

        file_name = f"誤配管理_{date}.xlsx"

        # アプリケーションのディレクトリからの相対パスを使用
        base_directory = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(base_directory, year, month, file_name)

        if os.path.exists(file_path):
            self.show_modify_append_dialog(file_path)
        else:
            self.show_attendance_input()

    def show_modify_append_dialog(self, file_path):
        dialog = QDialog(self)
        dialog.setWindowTitle("ファイルが存在します")
        dialog.setWindowModality(Qt.ApplicationModal)

        dialog_layout = QVBoxLayout()
        dialog_label = QLabel("同じ日付のファイルが存在します。どうしますか？", dialog)
        dialog_label.setFont(self.font)
        dialog_layout.addWidget(dialog_label)

        # 既存データの表示
        self.view_existing_data(dialog_layout, file_path)

        button_box = QDialogButtonBox(QDialogButtonBox.NoButton)
        modify_button = QPushButton("修正")
        append_button = QPushButton("追記")
        cancel_button = QPushButton("キャンセル")
        button_box.addButton(modify_button, QDialogButtonBox.ActionRole)
        button_box.addButton(append_button, QDialogButtonBox.ActionRole)
        button_box.addButton(cancel_button, QDialogButtonBox.RejectRole)

        modify_button.clicked.connect(lambda: self.modify_data(file_path, dialog))
        append_button.clicked.connect(lambda: self.append_data(file_path, dialog))
        cancel_button.clicked.connect(dialog.reject)

        dialog_layout.addWidget(button_box)
        dialog.setLayout(dialog_layout)

        dialog.exec_()

    def view_existing_data(self, layout, file_path):
        df = pd.read_excel(file_path)

        table_view = QTableView()
        model = QStandardItemModel()

        for column in df.columns:
            model.setHorizontalHeaderItem(df.columns.get_loc(column), QStandardItem(column))

        for row in df.itertuples(index=False):
            items = [QStandardItem(str(item)) for item in row]
            model.appendRow(items)

        table_view.setModel(model)
        table_view.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table_view.resizeColumnsToContents()

        layout.addWidget(table_view)

    def modify_data(self, file_path, dialog):
        dialog.accept()
        self.show_input_window(file_path, "修正")

    def append_data(self, file_path, dialog):
        dialog.accept()
        self.show_input_window(file_path, "追記")

    def show_attendance_input(self, existing_file_path=None):
        self.attendance_label.hide()
        self.attendance_input.hide()

        self.input_window = InputWindow(self.year_combobox.currentText(),
                                        self.month_combobox.currentText(),
                                        self.day_combobox.currentText(),
                                        existing_file_path,
                                        parent=self)
        self.input_window.show()

    def show_input_window(self, file_path, mode):
        self.input_window = InputWindow(self.year_combobox.currentText(),
                                        self.month_combobox.currentText(),
                                        self.day_combobox.currentText(),
                                        file_path,
                                        mode=mode,
                                        parent=self)
        self.input_window.show()

class InputWindow(QWidget):
    def __init__(self, year, month, day, existing_file_path=None, mode=None, parent=None):
        super().__init__()
        self.year = year
        self.month = month
        self.day = day
        self.existing_file_path = existing_file_path
        self.mode = mode
        self.attendance = 0
        self.parent = parent  # 親ウィンドウを保持
        self.font = QFont("Arial", 12)
        self.employee_list = self.load_employee_list()
        self.employee_inputs = []
        self.init_ui()

    def load_employee_list(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(script_dir, "社員名.txt")
        if os.path.exists(file_path):
            with open(file_path, "r", encoding="utf-8") as file:
                return [line.strip() for line in file.readlines()]
        else:
            QMessageBox.critical(self, "エラー", "社員名.txtが見つかりません。")
            sys.exit()

    def init_ui(self):
        self.setWindowTitle("データ入力")
        layout = QVBoxLayout()

        if self.mode == "修正":
            self.load_existing_data(layout)
        else:
            self.create_attendance_input(layout)

        self.setLayout(layout)
        self.adjustSize()

    def create_attendance_input(self, layout):
        if self.mode == "追記":
            self.attendance_label = QLabel("追加する人数の入力:", font=self.font)
        else:
            self.attendance_label = QLabel("出勤人数:", font=self.font)
        
        layout.addWidget(self.attendance_label)

        self.attendance_input = QLineEdit()
        self.attendance_input.setFont(self.font)
        layout.addWidget(self.attendance_input)

        self.next_button = QPushButton("次へ")
        self.next_button.setFont(self.font)
        self.next_button.clicked.connect(self.load_employee_input)
        layout.addWidget(self.next_button)

    def load_existing_data(self, layout):
        if self.existing_file_path:
            df = pd.read_excel(self.existing_file_path)
            self.attendance = len(df)

            for i, row in enumerate(df.itertuples(index=False)):
                self.create_employee_input(layout, i, row)

        # 修正モード時に完了ボタンを追加
        if self.mode == "修正":
            self.complete_button = QPushButton("完了")
            self.complete_button.setFont(self.font)
            self.complete_button.clicked.connect(self.show_confirmation_dialog)
            layout.addWidget(self.complete_button)

        # 終了ボタンを追加
        self.quit_button = QPushButton("終了")
        self.quit_button.setFont(self.font)
        self.quit_button.clicked.connect(self.close)
        layout.addWidget(self.quit_button)

    def load_employee_input(self):
        if not self.attendance_input.text().isdigit() or int(self.attendance_input.text()) <= 0:
            QMessageBox.critical(self, "入力エラー", "追加する人数は正の整数で入力してください。")
            return
        
        self.attendance = int(self.attendance_input.text())
        self.init_employee_input()

    def init_employee_input(self):
        layout = self.layout()

        for i in range(self.attendance):
            self.create_employee_input(layout, i)

        button_layout = QHBoxLayout()
        button_layout.addStretch(1)

        self.save_button = QPushButton("保存")
        self.save_button.setFont(self.font)
        self.save_button.clicked.connect(self.show_confirmation_dialog)
        button_layout.addWidget(self.save_button)

        self.quit_button = QPushButton("終了")
        self.quit_button.setFont(self.font)
        self.quit_button.clicked.connect(self.close)
        button_layout.addWidget(self.quit_button)

        button_layout.addStretch(1)

        layout.addLayout(button_layout)
        self.adjustSize()

    def create_employee_input(self, layout, i, row=None):
        emp_layout = QHBoxLayout()

        emp_label = QLabel(f"社員 {i+1}:")
        emp_label.setFont(self.font)
        emp_layout.addWidget(emp_label)

        employee_combobox = QComboBox()
        employee_combobox.setFont(self.font)
        employee_combobox.addItems(self.employee_list)
        if row:
            employee_combobox.setCurrentText(row.社員)
        emp_layout.addWidget(employee_combobox)

        morning_input = self.create_input_field("午前の持ち出し個数", self.font, emp_layout)
        if row:
            morning_input.setText(str(row.午前の持ち出し個数))

        afternoon_input = self.create_input_field("午後の持ち出し個数", self.font, emp_layout)
        if row:
            afternoon_input.setText(str(row.午後の持ち出し個数))

        error_input = self.create_input_field("誤配数", self.font, emp_layout)
        if row:
            error_input.setText(str(row.誤配数))

        delete_button = QPushButton("削除")
        delete_button.setFont(self.font)
        delete_button.clicked.connect(lambda: self.delete_employee_input(emp_layout, i))
        emp_layout.addWidget(delete_button)

        self.employee_inputs.append((employee_combobox, morning_input, afternoon_input, error_input, emp_layout))
        layout.addLayout(emp_layout)

    def delete_employee_input(self, layout, index):
        # リストから削除
        if 0 <= index < len(self.employee_inputs):
            for widget in self.employee_inputs[index][:-1]:  # レイアウト以外のウィジェットを削除
                widget.deleteLater()
            self.employee_inputs[index][-1].deleteLater()  # レイアウトを削除
            self.employee_inputs.pop(index)

        # ウィンドウサイズを内容に基づいて自動調整
        self.adjustSize()

    def create_input_field(self, label_text, font, layout):
        label = QLabel(label_text)
        label.setFont(font)
        layout.addWidget(label)
        line_edit = QLineEdit()
        line_edit.setFont(font)
        layout.addWidget(line_edit)
        return line_edit

    def save_data(self):
        data = []
        for i in range(self.attendance):
            if i < len(self.employee_inputs):
                employee_name = self.employee_inputs[i][0].currentText()
                morning = int(self.employee_inputs[i][1].text() or '0')
                afternoon = int(self.employee_inputs[i][2].text() or '0')
                error = int(self.employee_inputs[i][3].text() or '0')

                total_delivery = morning + afternoon
                error_rate = (error / total_delivery * 100) if total_delivery > 0 else 0

                data.append({
                    "社員": employee_name,
                    "午前の持ち出し個数": morning,
                    "午後の持ち出し個数": afternoon,
                    "持ち出し総数": total_delivery,
                    "誤配数": error,
                    "誤配率 (%)": f"{error_rate:.2f}%"
                })

        df = pd.DataFrame(data)
        file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), self.year, self.month, f"誤配管理_{self.year}_{self.month}_{self.day}.xlsx")

        if not os.path.exists(os.path.dirname(file_path)):
            os.makedirs(os.path.dirname(file_path))

        if os.path.exists(file_path):
            existing_df = pd.read_excel(file_path)
            if self.mode == "修正":
                # 修正モードの場合、データを上書き
                df = pd.concat([existing_df.iloc[:0], df], ignore_index=True)
            else:
                df = pd.concat([existing_df, df], ignore_index=True)

        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
            worksheet = writer.sheets['Sheet1']
            for column in df:
                column_width = 20
                col_idx = df.columns.get_loc(column)
                worksheet.set_column(col_idx, col_idx, column_width)
        
        QMessageBox.information(self, "保存完了", f"データが {file_path} に保存されました。")
        QApplication.quit()  # 保存後にアプリケーションを終了

    def show_confirmation_dialog(self):
        confirmation = QMessageBox.question(
            self, "確認", "入力内容を保存しますか？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if confirmation == QMessageBox.Yes:
            self.save_data()

if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = AttendanceApp()
    window.show()

    sys.exit(app.exec_())
