import sys
import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLabel, QComboBox, QMessageBox

# 誤配率の計算関数
def calculate_misdelivery_rate(total_deliveries, total_misdeliveries):
    if total_deliveries == 0:
        return 0
    return (total_misdeliveries / total_deliveries) * 100

# 月次集計機能（社員ごとの総持ち出し総数と総誤配数、全体の総誤配率の計算）
def monthly_aggregation(base_directory, year, month):
    try:
        month_folder = os.path.join(base_directory, year, month)
        monthly_summary = pd.DataFrame()

        # 月フォルダ内の全Excelファイルを結合、既存の月次集計ファイルは除外
        for filename in os.listdir(month_folder):
            if filename.endswith(".xlsx") and not filename.startswith(f"{year}_{month}_月次集計"):
                file_path = os.path.join(month_folder, filename)
                df = pd.read_excel(file_path)
                monthly_summary = pd.concat([monthly_summary, df])

        # 社員ごとの総持ち出し総数と総誤配数の集計
        monthly_total = monthly_summary.groupby("社員").agg({
            '持ち出し総数': 'sum',
            '誤配数': 'sum'
        }).reset_index()

        # 全体の持ち出し総数と誤配数の計算
        total_deliveries = monthly_total['持ち出し総数'].sum()
        total_misdeliveries = monthly_total['誤配数'].sum()
        overall_misdelivery_rate = calculate_misdelivery_rate(total_deliveries, total_misdeliveries)

        # 社員ごとの誤配率の計算
        monthly_total['誤配率'] = monthly_total.apply(
            lambda x: calculate_misdelivery_rate(x['持ち出し総数'], x['誤配数']), axis=1
        )
        monthly_total['誤配率'] = monthly_total['誤配率'].apply(lambda x: f"{x:.2f}%")

        # 全体の誤配率の行を追加
        overall_row = pd.DataFrame({
            '社員': ['全体'],
            '持ち出し総数': [total_deliveries],
            '誤配数': [total_misdeliveries],
            '誤配率': [f"{overall_misdelivery_rate:.2f}%"]
        })
        monthly_total = pd.concat([monthly_total, overall_row], ignore_index=True)

        # 保存
        output_path = os.path.join(month_folder, f"{year}_{month}_月次集計.xlsx")
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            monthly_total.to_excel(writer, index=False)
            worksheet = writer.sheets['Sheet1']
            for idx, col in enumerate(monthly_total.columns):
                worksheet.set_column(idx, idx, 20)  # 列幅を20に固定

        return output_path

    except PermissionError:
        QMessageBox.critical(None, "アクセスエラー", f"ファイルにアクセスできません。Excelが開いていないことを確認してください。")
    except Exception as e:
        QMessageBox.critical(None, "エラー", f"月次集計中にエラーが発生しました: {str(e)}")

# 年次集計機能（社員ごとの総持ち出し総数と総誤配数、全体の総誤配率の計算）
def yearly_aggregation(base_directory, year):
    try:
        year_folder = os.path.join(base_directory, year)
        yearly_summary = pd.DataFrame()

        # 年フォルダ内の月フォルダから全Excelファイルを読み込み、既存の集計ファイルは除外
        for month in os.listdir(year_folder):
            month_folder = os.path.join(year_folder, month)
            if os.path.isdir(month_folder):
                for filename in os.listdir(month_folder):
                    if filename.endswith(".xlsx") and not filename.startswith(f"{year}_") and not filename.endswith("月次集計.xlsx"):
                        file_path = os.path.join(month_folder, filename)
                        df = pd.read_excel(file_path)
                        yearly_summary = pd.concat([yearly_summary, df])

        # 社員ごとの総持ち出し総数と総誤配数の集計
        yearly_total = yearly_summary.groupby("社員").agg({
            '持ち出し総数': 'sum',
            '誤配数': 'sum'
        }).reset_index()

        # 全体の持ち出し総数と誤配数の計算
        total_deliveries = yearly_total['持ち出し総数'].sum()
        total_misdeliveries = yearly_total['誤配数'].sum()
        overall_misdelivery_rate = calculate_misdelivery_rate(total_deliveries, total_misdeliveries)

        # 社員ごとの誤配率の計算
        yearly_total['誤配率'] = yearly_total.apply(
            lambda x: calculate_misdelivery_rate(x['持ち出し総数'], x['誤配数']), axis=1
        )
        yearly_total['誤配率'] = yearly_total['誤配率'].apply(lambda x: f"{x:.2f}%")

        # 全体の誤配率の行を追加
        overall_row = pd.DataFrame({
            '社員': ['全体'],
            '持ち出し総数': [total_deliveries],
            '誤配数': [total_misdeliveries],
            '誤配率': [f"{overall_misdelivery_rate:.2f}%"]
        })
        yearly_total = pd.concat([yearly_total, overall_row], ignore_index=True)

        # 保存
        output_path = os.path.join(year_folder, f"{year}_年次集計.xlsx")
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            yearly_total.to_excel(writer, index=False)
            worksheet = writer.sheets['Sheet1']
            for idx, col in enumerate(yearly_total.columns):
                worksheet.set_column(idx, idx, 20)  # 列幅を20に固定

        return output_path

    except PermissionError:
        QMessageBox.critical(None, "アクセスエラー", f"ファイルにアクセスできません。Excelが開いていないことを確認してください。")
    except Exception as e:
        QMessageBox.critical(None, "エラー", f"年次集計中にエラーが発生しました: {str(e)}")

# メインウィンドウの実装
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        self.label = QLabel("集計処理を実行してください。")
        layout.addWidget(self.label)

        # 年選択用コンボボックス
        self.year_combobox = QComboBox()
        self.year_combobox.addItems([str(y) for y in range(2023, 2031)])  # 適宜、年を追加
        layout.addWidget(self.year_combobox)

        # 月選択用コンボボックス
        self.month_combobox = QComboBox()
        self.month_combobox.addItems([f"{m:02}" for m in range(1, 13)])
        layout.addWidget(self.month_combobox)

        self.monthly_button = QPushButton('月次集計')
        self.monthly_button.clicked.connect(self.monthly_aggregation)
        layout.addWidget(self.monthly_button)

        self.yearly_button = QPushButton('年次集計')
        self.yearly_button.clicked.connect(self.yearly_aggregation)
        layout.addWidget(self.yearly_button)

        self.exit_button = QPushButton('終了')
        self.exit_button.clicked.connect(self.close_application)
        layout.addWidget(self.exit_button)

        self.setLayout(layout)
        self.setWindowTitle('誤配管理集計')
        self.adjustSize()  # ウィンドウサイズを自動調整

    def monthly_aggregation(self):
        try:
            base_directory = 'C:\\Users\\Owner\\OneDrive\\デスクトップ\\誤配管理'
            year = self.year_combobox.currentText()
            month = self.month_combobox.currentText()
            output_path = monthly_aggregation(base_directory, year, month)
            self.label.setText(f"月次集計完了: {output_path}")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"月次集計中にエラーが発生しました: {str(e)}")

    def yearly_aggregation(self):
        try:
            base_directory = 'C:\\Users\\Owner\\OneDrive\\デスクトップ\\誤配管理'
            year = self.year_combobox.currentText()
            output_path = yearly_aggregation(base_directory, year)
            self.label.setText(f"年次集計完了: {output_path}")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"年次集計中にエラーが発生しました: {str(e)}")

    def close_application(self):
        self.close()

# アプリケーションのエントリーポイント
if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
