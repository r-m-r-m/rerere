from PyQt5 import QtWidgets, QtGui
import pandas as pd
import os
from datetime import datetime

class SagawaManagementSystem(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("佐川急便 履行率管理システム")
        
        # メインウィジェットとレイアウト
        self.central_widget = QtWidgets.QWidget(self)
        self.setCentralWidget(self.central_widget)
        self.main_layout = QtWidgets.QVBoxLayout(self.central_widget)

        # 日付選択レイアウト
        self.date_layout = QtWidgets.QHBoxLayout()
        self.year_label = QtWidgets.QLabel("年:")
        self.year_input = QtWidgets.QComboBox()
        self.year_input.addItems([str(year) for year in range(2020, 2031)])

        self.month_label = QtWidgets.QLabel("月:")
        self.month_input = QtWidgets.QComboBox()
        self.month_input.addItems([f"{month:02d}" for month in range(1, 13)])

        self.day_label = QtWidgets.QLabel("日:")
        self.day_input = QtWidgets.QComboBox()
        self.day_input.addItems([f"{day:02d}" for day in range(1, 32)])

        # 現在の日付をデフォルト設定
        self.set_current_date()

        # 日付選択をレイアウトに追加
        self.date_layout.addWidget(self.year_label)
        self.date_layout.addWidget(self.year_input)
        self.date_layout.addWidget(self.month_label)
        self.date_layout.addWidget(self.month_input)
        self.date_layout.addWidget(self.day_label)
        self.date_layout.addWidget(self.day_input)

        self.main_layout.addLayout(self.date_layout)

        # 読みやすさのためのフォント設定
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        self.setFont(font)

        # 保存・追記オプションのレイアウト
        self.action_layout = QtWidgets.QHBoxLayout()
        self.modify_button = QtWidgets.QPushButton("修正")
        self.modify_button.clicked.connect(self.modify_data)
        self.modify_button.setVisible(False)  # 初期状態では非表示

        self.append_button = QtWidgets.QPushButton("追記")
        self.append_button.clicked.connect(self.append_data)
        self.append_button.setVisible(False)  # 初期状態では非表示

        self.action_layout.addWidget(self.modify_button)
        self.action_layout.addWidget(self.append_button)
        self.main_layout.addLayout(self.action_layout)

        # 出勤人数入力
        self.employee_count_label = QtWidgets.QLabel("出勤人数:")
        self.employee_count_input = QtWidgets.QLineEdit()
        self.employee_count_input.setPlaceholderText("出勤人数を入力")
        self.employee_count_input.setValidator(QtGui.QIntValidator(1, 999))  # 入力を有効な整数に制限
        self.employee_count_input.textChanged.connect(self.update_employee_fields)
        self.employee_count_layout = QtWidgets.QHBoxLayout()
        self.employee_count_layout.addWidget(self.employee_count_label)
        self.employee_count_layout.addWidget(self.employee_count_input)
        self.main_layout.addLayout(self.employee_count_layout)

        # データ入力エリア
        self.data_area = QtWidgets.QWidget(self)
        self.data_layout = QtWidgets.QGridLayout(self.data_area)
        self.main_layout.addWidget(self.data_area)

        # 保存ボタン
        self.submit_button = QtWidgets.QPushButton("データ保存")
        self.submit_button.clicked.connect(self.save_data)
        self.main_layout.addWidget(self.submit_button)

        # 終了ボタン
        self.exit_button = QtWidgets.QPushButton("終了")
        self.exit_button.clicked.connect(self.close)
        self.main_layout.addWidget(self.exit_button)

        # コンテンツに基づいてウィンドウサイズを調整
        self.adjustSize()
        self.show()

    def set_current_date(self):
        # 今日の日付を取得
        today = datetime.today()
        current_year = today.year
        current_month = today.month
        current_day = today.day

        # コンボボックスに現在の日付を設定
        self.year_input.setCurrentText(str(current_year))
        self.month_input.setCurrentText(f"{current_month:02d}")
        self.day_input.setCurrentText(f"{current_day:02d}")

    def check_existing_date(self):
        # 選択された日付を取得
        year = self.year_input.currentText()
        month = self.month_input.currentText()
        day = self.day_input.currentText()
        date = f"{year}-{month}-{day}"

        # ファイルの存在を確認
        file_path = f"C:/Users/Owner/OneDrive/デスクトップ/佐川急便管理/履行率管理_{month}_{day}.xlsx"
        if os.path.exists(file_path):
            # ファイルが存在する場合は修正と追記ボタンを表示
            self.modify_button.setVisible(True)
            self.append_button.setVisible(True)
            QtWidgets.QMessageBox.information(self, "日付確認", f"{date}のデータが既に存在します。修正または追記を選択してください。")
        else:
            # ファイルが存在しない場合は新規データ入力を続行
            self.modify_button.setVisible(False)
            self.append_button.setVisible(False)

    def modify_data(self):
        # 既存データの修正ロジック
        QtWidgets.QMessageBox.information(self, "修正", "既存データの修正を行います。")
        # データ修正機能の実装はここに追加

    def append_data(self):
        # 既存データへの追記ロジック
        QtWidgets.QMessageBox.information(self, "追記", "既存データに追記を行います。")
        # データ追記機能の実装はここに追加

    def load_employee_names(self):
        file_path = r"C:\Users\Owner\OneDrive\デスクトップ\佐川急便管理\社員名.txt"
        if os.path.exists(file_path):
            with open(file_path, "r", encoding="utf-8") as file:
                names = [line.strip() for line in file.readlines()]
            return names
        else:
            QtWidgets.QMessageBox.warning(self, "ファイルエラー", "社員名ファイルが見つかりません。")
            return []

    def update_employee_fields(self):
        if not self.employee_count_input.text().isdigit():
            return

        count = int(self.employee_count_input.text())
        employee_names = self.load_employee_names()
        # レイアウト内の既存ウィジェットをクリア
        for i in reversed(range(self.data_layout.count())): 
            self.data_layout.itemAt(i).widget().setParent(None)
        
        # 各社員の入力フィールドを追加
        self.employee_fields = []
        for i in range(count):
            row = []
            labels = ["社員名", "持ち出し総数", "不履行数", "クレーム", "誤配", "遅刻", "事故"]
            for j, label in enumerate(labels):
                lbl = QtWidgets.QLabel(label)
                if j == 0:
                    input_field = QtWidgets.QComboBox()  # 社員名にはコンボボックスを使用
                    input_field.addItems(employee_names)
                else:
                    input_field = QtWidgets.QLineEdit()
                    input_field.setText("0")  # デフォルト値は0
                self.data_layout.addWidget(lbl, i, j * 2)
                self.data_layout.addWidget(input_field, i, j * 2 + 1)
                row.append(input_field)
            self.employee_fields.append(row)
        
        # 新しいコンテンツに基づいてウィンドウサイズを調整
        self.adjustSize()

    def save_data(self):
        year = self.year_input.currentText()
        month = self.month_input.currentText()
        day = self.day_input.currentText()
        date = f"{year}-{month}-{day}"
        count = int(self.employee_count_input.text())
        data = []

        for fields in self.employee_fields:
            employee_data = [field.currentText() if isinstance(field, QtWidgets.QComboBox) else field.text() for field in fields]
            data.append(employee_data)

        columns = ["社員名", "持ち出し総数", "不履行数", "クレーム", "誤配", "遅刻", "事故"]
        df = pd.DataFrame(data, columns=columns)

        # 履行率を計算（小数点第2位まで）
        df['履行率'] = df.apply(lambda row: round((int(row['持ち出し総数']) - int(row['不履行数'])) / int(row['持ち出し総数']) * 100, 2) 
                                if int(row['持ち出し総数']) > 0 else 0, axis=1)

        # 全体平均履行率を計算
        average_fulfillment_rate = df['履行率'].mean()

        file_path = f"C:/Users/Owner/OneDrive/デスクトップ/佐川急便管理/履行率管理_{month}_{day}.xlsx"
        if os.path.exists(file_path):
            existing_df = pd.read_excel(file_path)
            existing_employee_names = existing_df['社員名'].tolist()

            for _, row in df.iterrows():
                employee_name = row['社員名']
                if employee_name in existing_employee_names:
                    existing_df.loc[existing_df['社員名'] == employee_name, df.columns] = row
                else:
                    existing_df = pd.concat([existing_df, row.to_frame().T], ignore_index=True)

            # 上書き保存
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                existing_df.to_excel(writer, index=False)
                # 列幅調整
                for column in existing_df:
                    column_width = max(existing_df[column].astype(str).map(len).max(), len(column))
                    col_idx = existing_df.columns.get_loc(column)
                    writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)
        else:
            # 新規保存
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
                # 列幅調整
                for column in df:
                    column_width = max(df[column].astype(str).map(len).max(), len(column))
                    col_idx = df.columns.get_loc(column)
                    writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)

        QtWidgets.QMessageBox.information(self, "保存完了", f"{date}のデータが正常に保存されました。\n全体平均履行率: {average_fulfillment_rate:.2f}%")
        self.close()

    def closeEvent(self, event):
        reply = QtWidgets.QMessageBox.question(self, '終了', '本当に終了しますか?', QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No)
        if reply == QtWidgets.QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    mainWin = SagawaManagementSystem()
    mainWin.raise_()
    sys.exit(app.exec_())
