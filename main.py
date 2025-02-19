"""
コールセンター業務効率化ツール

このスクリプトは、コールセンター業務の効率化を目的としたGUIアプリケーションです。
PySide6を使用してUIを構築し、Google Spreadsheetsとの連携機能を提供します。

主な機能：
- 顧客情報の入力
- CTIフォーマットの生成
- スプレッドシートへのデータ転記
"""

import sys
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                              QHBoxLayout, QLabel, QLineEdit, QComboBox,
                              QPushButton, QTextEdit, QGroupBox, QMessageBox,
                              QScrollArea)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

class MainWindow(QMainWindow):
    """
    メインウィンドウクラス
    
    アプリケーションのメインウィンドウとUIコンポーネントを管理します。
    """
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("コールセンター業務効率化ツール")
        self.setMinimumSize(1000, 800)
        
        # メインウィジェットの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # メインレイアウトの設定
        main_layout = QVBoxLayout(main_widget)
        
        # トップバーの作成
        self.create_top_bar(main_layout)
        
        # メイン部分のレイアウト
        content_layout = QHBoxLayout()
        
        # 入力フォームエリア（左側70%）をスクロール可能に
        form_widget = QWidget()
        form_layout = QVBoxLayout(form_widget)
        self.create_input_form(form_layout)
        
        # スクロールエリアの作成
        scroll_area = QScrollArea()
        scroll_area.setWidget(form_widget)
        scroll_area.setWidgetResizable(True)  # ウィジェットのリサイズを許可
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # スクロールエリアのスタイルシートを更新
        scroll_area.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QScrollBar:vertical {
                border: none;
                background: #E0E0E0;
                width: 12px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #808080;
                min-height: 20px;
                border-radius: 6px;
                margin: 2px;
            }
            QScrollBar::handle:vertical:hover {
                background: #606060;
            }
            QScrollBar::handle:vertical:pressed {
                background: #404040;
            }
            QScrollBar::add-line:vertical {
                height: 0px;
            }
            QScrollBar::sub-line:vertical {
                height: 0px;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: #E0E0E0;
                border-radius: 6px;
            }
        """)
        
        content_layout.addWidget(scroll_area, 70)
        
        # プレビューエリア（右側30%）
        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)
        self.create_preview_area(preview_layout)
        content_layout.addWidget(preview_widget, 30)
        
        main_layout.addLayout(content_layout)

        # 全てのUIコンポーネントを作成した後にシグナルを接続
        self.setup_signals()

        # Google Sheets APIの設定
        self.CREDENTIALS_PATH = "D:\\Cursor\\key\\telephonetool-42f9d38d14ee.json"
        self.SPREADSHEET_ID = "1Y3M8YZ0ywLdMxCVY6EB8COcPkBaOS6e27jzEJWcU8Tw"
        
        try:
            self.setup_google_sheets()
        except Exception as e:
            print(f"Google Sheets設定エラー: {str(e)}")

    def create_top_bar(self, parent_layout):
        """トップバーを作成"""
        top_bar = QWidget()
        top_bar.setStyleSheet("background-color: #2C3E50; color: white;")
        top_bar_layout = QHBoxLayout(top_bar)
        
        # ボタンの作成
        self.clear_btn = QPushButton("入力クリア")
        self.cti_copy_btn = QPushButton("CTIコピー")
        self.spreadsheet_btn = QPushButton("スプレッドシート転記")
        
        # ボタンのスタイル設定
        button_style = """
            QPushButton {
                color: white;
                border: 1px solid white;
                padding: 5px;
                border-radius: 3px;
                background-color: #2C3E50;
            }
            QPushButton:hover {
                background-color: #34495E;
                border: 1px solid #3498DB;
            }
            QPushButton:pressed {
                background-color: #2980B9;
            }
        """
        self.clear_btn.setStyleSheet(button_style)
        self.cti_copy_btn.setStyleSheet(button_style)
        self.spreadsheet_btn.setStyleSheet(button_style)
        
        top_bar_layout.addWidget(self.clear_btn)
        top_bar_layout.addWidget(self.cti_copy_btn)
        top_bar_layout.addWidget(self.spreadsheet_btn)
        
        parent_layout.addWidget(top_bar)

    def create_input_form(self, parent_layout):
        """入力フォームを作成"""
        # 基本情報セクション
        basic_info_group = QGroupBox("基本情報")
        basic_layout = QVBoxLayout()
        
        # 対応者名
        basic_layout.addWidget(QLabel("対応者名"))
        self.operator_input = QLineEdit()
        basic_layout.addWidget(self.operator_input)
        
        # 携帯電話番号
        basic_layout.addWidget(QLabel("携帯電話番号"))
        self.mobile_input = QLineEdit()
        basic_layout.addWidget(self.mobile_input)
        
        # 契約者名
        basic_layout.addWidget(QLabel("契約者名"))
        self.contractor_input = QLineEdit()
        basic_layout.addWidget(self.contractor_input)
        
        # フリガナ
        basic_layout.addWidget(QLabel("フリガナ"))
        self.furigana_input = QLineEdit()
        basic_layout.addWidget(self.furigana_input)
        
        # 生年月日
        birth_layout = QHBoxLayout()
        birth_layout.addWidget(QLabel("生年月日"))
        
        self.era_combo = QComboBox()
        self.era_combo.addItems(["昭和", "平成", "西暦"])
        birth_layout.addWidget(self.era_combo)
        
        self.year_combo = QComboBox()
        # 初期値として昭和の年を設定
        self.year_combo.addItems([str(i) for i in range(1, 65)])
        birth_layout.addWidget(self.year_combo)
        birth_layout.addWidget(QLabel("年"))
        
        self.month_combo = QComboBox()
        self.month_combo.addItems([str(i) for i in range(1, 13)])
        birth_layout.addWidget(self.month_combo)
        birth_layout.addWidget(QLabel("月"))
        
        self.day_combo = QComboBox()
        self.day_combo.addItems([str(i) for i in range(1, 32)])
        birth_layout.addWidget(self.day_combo)
        birth_layout.addWidget(QLabel("日"))
        
        basic_layout.addLayout(birth_layout)
        basic_info_group.setLayout(basic_layout)
        parent_layout.addWidget(basic_info_group)
        
        # 住所情報セクション
        address_group = QGroupBox("住所情報")
        address_layout = QVBoxLayout()
        
        # 郵便番号
        address_layout.addWidget(QLabel("郵便番号"))
        self.postal_code_input = QLineEdit()
        address_layout.addWidget(self.postal_code_input)
        
        # 住所
        address_layout.addWidget(QLabel("住所"))
        self.address_input = QLineEdit()
        address_layout.addWidget(self.address_input)
        
        address_group.setLayout(address_layout)
        parent_layout.addWidget(address_group)
        
        # リスト情報セクション
        list_group = QGroupBox("リスト情報")
        list_layout = QVBoxLayout()
        
        # リスト名
        list_layout.addWidget(QLabel("リスト名"))
        self.list_name_input = QLineEdit()
        list_layout.addWidget(self.list_name_input)
        
        # リストフリガナ
        list_layout.addWidget(QLabel("リストフリガナ"))
        self.list_furigana_input = QLineEdit()
        list_layout.addWidget(self.list_furigana_input)
        
        # 電話番号
        list_layout.addWidget(QLabel("電話番号"))
        self.list_phone_input = QLineEdit()
        list_layout.addWidget(self.list_phone_input)
        
        # リスト郵便番号
        list_layout.addWidget(QLabel("リスト郵便番号"))
        self.list_postal_code_input = QLineEdit()
        list_layout.addWidget(self.list_postal_code_input)
        
        # リスト住所
        list_layout.addWidget(QLabel("リスト住所"))
        self.list_address_input = QLineEdit()
        list_layout.addWidget(self.list_address_input)
        
        list_group.setLayout(list_layout)
        parent_layout.addWidget(list_group)
        
        # 受注情報セクション
        order_group = QGroupBox("受注情報")
        order_layout = QVBoxLayout()
        
        # 現状回線
        order_layout.addWidget(QLabel("現状回線"))
        self.current_line_combo = QComboBox()
        self.current_line_combo.addItems(["アナログ"])
        order_layout.addWidget(self.current_line_combo)
        
        # 受注日（本日自動入力）
        order_layout.addWidget(QLabel("受注日"))
        self.order_date_input = QLineEdit()
        self.order_date_input.setText(datetime.datetime.now().strftime("%Y/%m/%d"))
        self.order_date_input.setReadOnly(True)
        order_layout.addWidget(self.order_date_input)
        
        # 受注者名
        order_layout.addWidget(QLabel("受注者名"))
        self.order_person_input = QLineEdit()
        order_layout.addWidget(self.order_person_input)
        
        # 提供判定
        order_layout.addWidget(QLabel("提供判定"))
        self.judgment_combo = QComboBox()
        self.judgment_combo.addItems(["OK", "NG"])
        order_layout.addWidget(self.judgment_combo)
        
        order_group.setLayout(order_layout)
        parent_layout.addWidget(order_group)
        
        # その他情報セクション
        other_group = QGroupBox("その他情報")
        other_layout = QVBoxLayout()
        
        # 料金認識
        other_layout.addWidget(QLabel("料金認識"))
        self.fee_input = QLineEdit()
        self.fee_input.setText("3000円～3500円")
        other_layout.addWidget(self.fee_input)
        
        # ネット利用
        other_layout.addWidget(QLabel("ネット利用"))
        self.net_usage_combo = QComboBox()
        self.net_usage_combo.addItems(["あり", "なし"])
        other_layout.addWidget(self.net_usage_combo)
        
        # 家族了承
        other_layout.addWidget(QLabel("家族了承"))
        self.family_approval_combo = QComboBox()
        self.family_approval_combo.addItems(["あり", "なし"])
        other_layout.addWidget(self.family_approval_combo)
        
        # 備考欄を追加
        other_layout.addWidget(QLabel("備考"))
        self.remarks_input = QTextEdit()  # QTextEditを使用して複数行の入力を可能に
        self.remarks_input.setMaximumHeight(100)  # 高さを制限
        other_layout.addWidget(self.remarks_input)
        
        other_group.setLayout(other_layout)
        parent_layout.addWidget(other_group)

        # 各QGroupBoxのマージンとスペーシングを調整
        margin = 5
        spacing = 5
        
        for group in [basic_info_group, address_group, list_group, order_group, other_group]:
            group.setContentsMargins(margin, margin, margin, margin)
            if hasattr(group.layout(), 'setSpacing'):
                group.layout().setSpacing(spacing)
        
        # 親レイアウトのマージンとスペーシングも調整
        parent_layout.setContentsMargins(margin, margin, margin, margin)
        parent_layout.setSpacing(spacing)
        
        # 各QLineEditの最小高さを設定
        for widget in parent_layout.parentWidget().findChildren(QLineEdit):
            widget.setMinimumHeight(25)
        
        # 各QComboBoxの最小高さを設定
        for widget in parent_layout.parentWidget().findChildren(QComboBox):
            widget.setMinimumHeight(25)

    def create_preview_area(self, parent_layout):
        """プレビューエリアを作成"""
        preview_label = QLabel("CTIフォーマットプレビュー")
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        
        parent_layout.addWidget(preview_label)
        parent_layout.addWidget(self.preview_text)

    def format_phone_number(self):
        """電話番号の自動フォーマット処理"""
        sender = self.sender()
        text = sender.text().replace('-', '')  # ハイフンを除去
        
        # 数字以外を除去
        text = ''.join(filter(str.isdigit, text))
        
        # 11桁以内に制限
        text = text[:11]
        
        # ハイフン挿入
        if len(text) > 7:
            text = f'{text[:3]}-{text[3:7]}-{text[7:]}'
        elif len(text) > 3:
            text = f'{text[:3]}-{text[3:]}'
        
        sender.setText(text)

    def format_postal_code(self):
        """郵便番号の自動フォーマット処理"""
        sender = self.sender()
        text = sender.text().replace('-', '')  # ハイフンを除去
        
        # 数字以外を除去
        text = ''.join(filter(str.isdigit, text))
        
        # 7桁以内に制限
        text = text[:7]
        
        # ハイフン挿入
        if len(text) > 3:
            text = f'{text[:3]}-{text[3:]}'
        
        sender.setText(text)

    def convert_to_half_width(self):
        """全角文字を半角に変換"""
        sender = self.sender()
        text = sender.text()
        
        # 全角数字を半角に変換
        for i in range(10):
            text = text.replace(chr(0xFF10 + i), str(i))
        
        # 全角ハイフン（マイナス）を半角に変換
        text = text.replace('−', '-')  # 全角マイナス
        text = text.replace('ー', '-')  # 長音符
        text = text.replace('―', '-')  # ダッシュ
        text = text.replace('‐', '-')  # ハイフン
        text = text.replace('－', '-')  # 全角ハイフン
        
        sender.setText(text)

    def generate_cti_format(self):
        """CTIフォーマットの生成とプレビュー表示"""
        # 生年月日の生成（和暦を西暦に変換）
        era = self.era_combo.currentText()
        year = int(self.year_combo.currentText())
        month = self.month_combo.currentText()
        day = self.day_combo.currentText()
        
        # 和暦を西暦に変換
        if era == "昭和":
            year = year + 1925  # 昭和元年は1926年
        elif era == "平成":
            year = year + 1988  # 平成元年は1989年
        # 西暦の場合はそのまま
        
        birth_date = f"{year}/{month}/{day}"
        
        # 現在の日付を取得（先頭の0を除去）
        now = datetime.datetime.now()
        month_str = str(now.month)
        day_str = str(now.day)
        
        # CTIフォーマットの生成
        cti_text = f"""工事希望日
携帯：{self.mobile_input.text()}
アナログ→光電話
契約者(書類名義)：{self.contractor_input.text()}
フリガナ：{self.furigana_input.text()}
生年月日：{birth_date}
郵便番号：{self.postal_code_input.text()}
住所：{self.address_input.text()}
リスト名：{self.list_name_input.text()}
リスト名フリガナ：{self.list_furigana_input.text()}
電話番号：{self.list_phone_input.text()}
リスト郵便番号：{self.list_postal_code_input.text()}
リスト住所：{self.list_address_input.text()}
現状回線：{self.current_line_combo.currentText()}
受注日：{month_str}/{day_str}
受注者：{self.order_person_input.text()}
提供判定：{self.judgment_combo.currentText()}

料金認識：{self.fee_input.text()}
ネット利用：{self.net_usage_combo.currentText()}
家族了承：{self.family_approval_combo.currentText()}

他番号：なし
電話機：プッシュホン
禁止回線：なし
ND：

備考：{self.remarks_input.toPlainText()}
お客様が今使っている回線：アナログ
案内料金：2500円
※リスト名との関係性："""
        
        # プレビューに表示
        self.preview_text.setText(cti_text)
        
        # クリップボードにコピー
        clipboard = QApplication.clipboard()
        clipboard.setText(cti_text)

    def clear_all_inputs(self):
        """全ての入力フィールドをクリア"""
        # テキスト入力フィールドのクリア
        self.operator_input.clear()
        self.mobile_input.clear()
        self.contractor_input.clear()
        self.furigana_input.clear()
        self.postal_code_input.clear()
        self.address_input.clear()
        self.list_name_input.clear()
        self.list_furigana_input.clear()
        self.list_phone_input.clear()
        self.list_postal_code_input.clear()
        self.list_address_input.clear()
        self.order_person_input.clear()
        
        # コンボボックスを初期値に戻す
        self.era_combo.setCurrentIndex(0)
        self.year_combo.setCurrentIndex(0)
        self.month_combo.setCurrentIndex(0)
        self.day_combo.setCurrentIndex(0)
        self.current_line_combo.setCurrentIndex(0)
        self.judgment_combo.setCurrentIndex(0)
        self.net_usage_combo.setCurrentIndex(0)
        self.family_approval_combo.setCurrentIndex(0)
        
        # 料金認識を初期値に戻す
        self.fee_input.setText("3000円～3500円")
        
        # プレビューエリアをクリア
        self.preview_text.clear()
        
        # 備考欄をクリア
        self.remarks_input.clear()

    def setup_google_sheets(self):
        """Google Sheets APIの設定"""
        scope = ['https://spreadsheets.google.com/feeds',
                'https://www.googleapis.com/auth/drive']
        
        credentials = ServiceAccountCredentials.from_json_keyfile_name(
            self.CREDENTIALS_PATH, scope)
        self.gc = gspread.authorize(credentials)
        self.sheet = self.gc.open_by_key(self.SPREADSHEET_ID).sheet1

    def write_to_spreadsheet(self):
        """スプレッドシートにデータを書き込む"""
        try:
            # 現在時刻と日付を取得
            current_time = datetime.datetime.now().strftime("%H:%M")
            current_date = datetime.datetime.now().strftime("%Y/%m/%d")
            
            # 電話番号からハイフンを除去
            tel1 = self.list_phone_input.text().replace('-', '')
            tel2 = self.mobile_input.text().replace('-', '')
            
            # 空のリストを作成（スプレッドシートの列数分）
            row_data = [''] * 16  # 画像から見える列数に合わせて調整
            
            # 必要な列にデータを設定（0から始まるインデックスなので、A列が0、B列が1）
            row_data[3] = tel1                          # TEL1（E列）
            row_data[4] = tel2                          # TEL2（F列）
            row_data[7] = current_time                  # 架電時間（I列）
            row_data[10] = current_date                 # トス日（L列）
            
            # スプレッドシートに追記
            self.sheet.append_row(row_data, value_input_option='RAW')
            
            # 成功メッセージを表示
            QMessageBox.information(self, "成功", "スプレッドシートへの転記が完了しました。")
            
        except Exception as e:
            # エラーメッセージを表示
            QMessageBox.critical(self, "エラー", f"スプレッドシートへの転記に失敗しました。\n{str(e)}")

    def update_year_combo(self):
        """元号に応じて年の選択肢を更新"""
        current_era = self.era_combo.currentText()
        current_year = datetime.datetime.now().year
        
        self.year_combo.clear()
        
        if current_era == "昭和":
            # 昭和1年から64年まで
            self.year_combo.addItems([str(i) for i in range(1, 65)])
        elif current_era == "平成":
            # 平成1年から31年まで
            self.year_combo.addItems([str(i) for i in range(1, 32)])
        else:  # 西暦
            # 1900年から現在の年まで
            self.year_combo.addItems([str(i) for i in range(1900, current_year + 1)])
        
        # 最初の項目を選択
        self.year_combo.setCurrentIndex(0)

    def setup_signals(self):
        """シグナルの接続"""
        # 自動フォーマット用のシグナル
        self.mobile_input.textChanged.connect(self.format_phone_number)
        self.list_phone_input.textChanged.connect(self.format_phone_number_without_hyphen)
        self.postal_code_input.textChanged.connect(self.format_postal_code)
        self.postal_code_input.textChanged.connect(self.convert_to_half_width)
        self.list_postal_code_input.textChanged.connect(self.format_postal_code)
        self.list_postal_code_input.textChanged.connect(self.convert_to_half_width)
        self.address_input.textChanged.connect(self.convert_to_half_width)
        self.list_address_input.textChanged.connect(self.convert_to_half_width)
        self.era_combo.currentTextChanged.connect(self.update_year_combo)
        
        # ボタンのシグナル接続
        self.clear_btn.clicked.connect(self.clear_all_inputs)
        self.cti_copy_btn.clicked.connect(self.generate_cti_format)
        self.spreadsheet_btn.clicked.connect(self.write_to_spreadsheet)

    def format_phone_number_without_hyphen(self):
        """電話番号の自動フォーマット処理（ハイフンなし）"""
        sender = self.sender()
        text = sender.text().replace('-', '')  # ハイフンを除去
        
        # 数字以外を除去
        text = ''.join(filter(str.isdigit, text))
        
        # 11桁以内に制限
        text = text[:11]
        
        sender.setText(text)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec()) 