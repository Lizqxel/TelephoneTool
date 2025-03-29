"""
使いやすいモードのダイアログ

このモジュールは、使いやすいモードの各ステップのダイアログを
提供します。
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
                              QLabel, QLineEdit, QPushButton,
                              QGroupBox, QMessageBox, QWidget, QComboBox, QScrollArea)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QIntValidator
import datetime
import logging

def convert_to_half_width(text):
    """
    全角文字を半角文字に変換する
    
    Args:
        text (str): 変換する文字列
        
    Returns:
        str: 変換後の文字列
    """
    if not text:
        return text
        
    # まず、全角ハイフンを半角に変換
    text = text.replace('－', '-')  # 全角ハイフン
    text = text.replace('ー', '-')  # 長音記号
    text = text.replace('−', '-')   # 別種の全角ハイフン
    text = text.replace('―', '-')   # ダッシュ
    text = text.replace('‐', '-')   # 別種のハイフン
    
    # 全角文字と半角文字の対応表
    trans_table = str.maketrans({
        '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
        '５': '5', '６': '6', '７': '7', '８': '8', '９': '9',
        'Ａ': 'A', 'Ｂ': 'B', 'Ｃ': 'C', 'Ｄ': 'D', 'Ｅ': 'E',
        'Ｆ': 'F', 'Ｇ': 'G', 'Ｈ': 'H', 'Ｉ': 'I', 'Ｊ': 'J',
        'Ｋ': 'K', 'Ｌ': 'L', 'Ｍ': 'M', 'Ｎ': 'N', 'Ｏ': 'O',
        'Ｐ': 'P', 'Ｑ': 'Q', 'Ｒ': 'R', 'Ｓ': 'S', 'Ｔ': 'T',
        'Ｕ': 'U', 'Ｖ': 'V', 'Ｗ': 'W', 'Ｘ': 'X', 'Ｙ': 'Y',
        'Ｚ': 'Z', 'ａ': 'a', 'ｂ': 'b', 'ｃ': 'c', 'ｄ': 'd',
        'ｅ': 'e', 'ｆ': 'f', 'ｇ': 'g', 'ｈ': 'h', 'ｉ': 'i',
        'ｊ': 'j', 'ｋ': 'k', 'ｌ': 'l', 'ｍ': 'm', 'ｎ': 'n',
        'ｏ': 'o', 'ｐ': 'p', 'ｑ': 'q', 'ｒ': 'r', 'ｓ': 's',
        'ｔ': 't', 'ｕ': 'u', 'ｖ': 'v', 'ｗ': 'w', 'ｘ': 'x',
        'ｙ': 'y', 'ｚ': 'z', '　': ' ', '！': '!', '＂': '"',
        '＃': '#', '＄': '$', '％': '%', '＆': '&', '＇': "'",
        '（': '(', '）': ')', '＊': '*', '＋': '+', '，': ',',
        '．': '.', '／': '/', '：': ':', '；': ';',
        '＜': '<', '＝': '=', '＞': '>', '？': '?', '＠': '@',
        '［': '[', '＼': '\\', '］': ']', '＾': '^', '＿': '_',
        '｀': '`', '｛': '{', '｜': '|', '｝': '}', '～': '~'
    })
    
    return text.translate(trans_table)

class AddressInfoDialog(QDialog):
    """住所情報入力ダイアログ"""
    
    def __init__(self, parent=None, address_data=None):
        """
        住所情報入力ダイアログの初期化
        
        Args:
            parent: 親ウィジェット
            address_data: 住所情報データ
        """
        super().__init__(parent)
        self.setWindowTitle("住所情報入力")
        self.setModal(True)
        self.setMinimumWidth(500)
        
        # メインレイアウト
        layout = QVBoxLayout()
        
        # タイトル
        title_label = QLabel("住所情報の確認")
        title_label.setFont(QFont("MS Gothic", 12, QFont.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)
        
        # 住所情報入力グループ
        address_group = QGroupBox("住所情報")
        address_layout = QVBoxLayout()
        
        # 郵便番号
        address_layout.addWidget(QLabel("郵便番号"))
        self.postal_code_input = QLineEdit()
        self.postal_code_input.setPlaceholderText("例：123-4567")
        if address_data and 'postal_code' in address_data:
            self.postal_code_input.setText(address_data['postal_code'])
        address_layout.addWidget(self.postal_code_input)
        
        # 住所
        address_layout.addWidget(QLabel("住所"))
        self.address_input = QLineEdit()
        self.address_input.setPlaceholderText("例：東京都渋谷区...")
        if address_data and 'address' in address_data:
            self.address_input.setText(address_data['address'])
        address_layout.addWidget(self.address_input)
        
        address_group.setLayout(address_layout)
        layout.addWidget(address_group)
        
        # 提供判定ボタン
        self.judgment_btn = QPushButton("提供判定実行")
        self.judgment_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3e8e41;
            }
        """)
        layout.addWidget(self.judgment_btn)
        
        # 提供判定結果表示ラベル
        self.judgment_result = QLabel()
        self.judgment_result.setStyleSheet("""
            QLabel {
                font-size: 14px;
                padding: 10px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: #f8f9fa;
            }
        """)
        self.judgment_result.hide()
        layout.addWidget(self.judgment_result)
        
        # ボタンエリア
        button_layout = QHBoxLayout()
        
        # 次へボタン
        self.next_btn = QPushButton("次へ")
        self.next_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #1565C0;
            }
        """)
        button_layout.addWidget(self.next_btn)
        
        # 作成中止ボタン
        self.cancel_btn = QPushButton("作成中止")
        self.cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
            QPushButton:pressed {
                background-color: #c62828;
            }
        """)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
        # シグナルの接続
        self.next_btn.clicked.connect(self.accept)
        self.cancel_btn.clicked.connect(self.reject)
    
    def get_address_data(self):
        """
        住所情報を取得する
        
        Returns:
            dict: 住所情報
        """
        try:
            # 住所の全角ハイフンを半角に変換
            address = self.address_input.text()
            address = address.replace('－', '-')  # 全角ハイフンを半角に
            address = address.replace('ー', '-')  # 長音記号を半角ハイフンに
            address = address.replace('−', '-')  # 別種の全角ハイフンを半角に
            address = address.replace('―', '-')  # ダッシュを半角ハイフンに
            address = address.replace('‐', '-')  # 別種のハイフンを半角ハイフンに
            
            # データを辞書形式で返す
            data = {
                'postal_code': self.postal_code_input.text(),
                'address': address
            }
            
            logging.info(f"住所データを取得: {data}")
            return data
            
        except Exception as e:
            logging.error(f"住所データの取得中にエラー: {e}")
            return {}
    
    def show_judgment_result(self, result):
        """
        提供判定結果を表示
        
        Args:
            result (str): 判定結果
        """
        self.judgment_result.setText(f"提供判定結果: {result}")
        self.judgment_result.show()

    def set_address_data(self, data):
        """
        住所情報を設定する
        
        Args:
            data: 住所情報データ
        """
        try:
            # 郵便番号
            if data.get('postal_code'):
                converted_postal_code = convert_to_half_width(data['postal_code'])
                self.postal_code_input.setText(converted_postal_code)
            
            # 住所
            if data.get('address'):
                converted_address = convert_to_half_width(data['address'])
                self.address_input.setText(converted_address)
            
            logging.info("住所データを正常に設定しました")
            
        except Exception as e:
            logging.error(f"住所情報の設定中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"住所情報の設定中にエラーが発生しました: {e}")

class ListInfoDialog(QDialog):
    """リスト情報入力ダイアログ"""
    
    def __init__(self, parent=None, list_data=None):
        """
        リスト情報入力ダイアログの初期化
        
        Args:
            parent: 親ウィジェット
            list_data: リスト情報データ
        """
        super().__init__(parent)
        self.setWindowTitle("リスト情報入力")
        self.setModal(True)
        self.setMinimumWidth(500)
        
        # メインレイアウト
        layout = QVBoxLayout()
        
        # タイトル
        title_label = QLabel("リスト情報の確認")
        title_label.setFont(QFont("MS Gothic", 12, QFont.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)
        
        # リスト情報入力グループ
        list_group = QGroupBox("リスト情報")
        list_layout = QVBoxLayout()
        
        # リスト名
        list_layout.addWidget(QLabel("リスト名"))
        self.list_name_input = QLineEdit()
        if list_data and 'list_name' in list_data:
            self.list_name_input.setText(list_data['list_name'])
        list_layout.addWidget(self.list_name_input)
        
        # リストフリガナ
        list_layout.addWidget(QLabel("リストフリガナ"))
        self.list_furigana_input = QLineEdit()
        self.list_furigana_input.setPlaceholderText("例：タナカタロウ")
        if list_data and 'list_furigana' in list_data:
            self.list_furigana_input.setText(list_data['list_furigana'])
        list_layout.addWidget(self.list_furigana_input)
        
        # 電話番号
        list_layout.addWidget(QLabel("電話番号"))
        self.list_phone_input = QLineEdit()
        self.list_phone_input.setPlaceholderText("例：090-1234-5678")
        if list_data and 'list_phone' in list_data:
            self.list_phone_input.setText(list_data['list_phone'])
        list_layout.addWidget(self.list_phone_input)
        
        # リスト郵便番号
        list_layout.addWidget(QLabel("リスト郵便番号"))
        self.list_postal_code_input = QLineEdit()
        self.list_postal_code_input.setPlaceholderText("例：123-4567")
        if list_data and 'list_postal_code' in list_data:
            self.list_postal_code_input.setText(list_data['list_postal_code'])
        list_layout.addWidget(self.list_postal_code_input)
        
        # リスト住所
        list_layout.addWidget(QLabel("リスト住所"))
        self.list_address_input = QLineEdit()
        self.list_address_input.setPlaceholderText("例：東京都渋谷区...")
        if list_data and 'list_address' in list_data:
            self.list_address_input.setText(list_data['list_address'])
        list_layout.addWidget(self.list_address_input)
        
        list_group.setLayout(list_layout)
        layout.addWidget(list_group)
        
        # ボタンエリア
        button_layout = QHBoxLayout()
        
        # 戻るボタン
        self.back_btn = QPushButton("戻る")
        self.back_btn.setStyleSheet("""
            QPushButton {
                background-color: #757575;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #616161;
            }
            QPushButton:pressed {
                background-color: #424242;
            }
        """)
        button_layout.addWidget(self.back_btn)
        
        # 次へボタン
        self.next_btn = QPushButton("次へ")
        self.next_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #1565C0;
            }
        """)
        button_layout.addWidget(self.next_btn)
        
        # 作成中止ボタン
        self.cancel_btn = QPushButton("作成中止")
        self.cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
            QPushButton:pressed {
                background-color: #c62828;
            }
        """)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
        # シグナルの接続
        self.back_btn.clicked.connect(self.reject)
        self.next_btn.clicked.connect(self.accept)
        self.cancel_btn.clicked.connect(self.reject)
    
    def get_list_data(self):
        """
        リスト情報を取得する
        
        Returns:
            dict: リスト情報
        """
        try:
            # リスト住所の全角ハイフンを半角に変換
            list_address = self.list_address_input.text()
            list_address = list_address.replace('－', '-')  # 全角ハイフンを半角に
            list_address = list_address.replace('ー', '-')  # 長音記号を半角ハイフンに
            list_address = list_address.replace('−', '-')  # 別種の全角ハイフンを半角に
            list_address = list_address.replace('―', '-')  # ダッシュを半角ハイフンに
            list_address = list_address.replace('‐', '-')  # 別種のハイフンを半角ハイフンに
            
            # データを辞書形式で返す
            data = {
                'list_name': self.list_name_input.text(),
                'list_furigana': self.list_furigana_input.text(),
                'list_phone': self.list_phone_input.text(),
                'list_postal_code': self.list_postal_code_input.text(),
                'list_address': list_address
            }
            
            logging.info(f"リストデータを取得: {data}")
            return data
            
        except Exception as e:
            logging.error(f"リストデータの取得中にエラー: {e}")
            return {}

    def set_list_data(self, data):
        """
        リスト情報を設定する
        
        Args:
            data: リスト情報データ
        """
        try:
            # リスト名
            if data.get('list_name'):
                converted_name = convert_to_half_width(data['list_name'])
                self.list_name_input.setText(converted_name)
            
            # リストフリガナ
            if data.get('list_furigana'):
                converted_furigana = convert_to_half_width(data['list_furigana'])
                self.list_furigana_input.setText(converted_furigana)
            
            # 電話番号
            if data.get('list_phone'):
                converted_phone = convert_to_half_width(data['list_phone'])
                self.list_phone_input.setText(converted_phone)
            
            # リスト郵便番号
            if data.get('list_postal_code'):
                converted_postal_code = convert_to_half_width(data['list_postal_code'])
                self.list_postal_code_input.setText(converted_postal_code)
            
            # リスト住所
            if data.get('list_address'):
                converted_address = convert_to_half_width(data['list_address'])
                self.list_address_input.setText(converted_address)
            
            logging.info("リストデータを正常に設定しました")
            
        except Exception as e:
            logging.error(f"リスト情報の設定中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"リスト情報の設定中にエラーが発生しました: {e}")

class OrdererInputDialog(QDialog):
    """受注者入力項目ダイアログ"""
    
    def __init__(self, parent=None, orderer_data=None):
        """
        受注者入力項目ダイアログの初期化
        
        Args:
            parent: 親ウィジェット
            orderer_data: 受注者情報データ
        """
        super().__init__(parent)
        self.setWindowTitle("受注者入力項目")
        self.setModal(True)
        self.setMinimumWidth(500)
        
        # メインレイアウト
        layout = QVBoxLayout()
        
        # タイトル
        title_label = QLabel("受注者入力項目")
        title_label.setFont(QFont("MS Gothic", 12, QFont.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)
        
        # スクロールエリアの作成
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # スクロールエリアのスタイルシートを設定
        scroll_area.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QScrollBar:vertical {
                border: none;
                background: #F0F0F0;
                width: 10px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #CCCCCC;
                min-height: 20px;
                border-radius: 5px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: none;
            }
        """)
        
        # スクロールエリア内のウィジェット
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        
        # 受注者情報入力グループ
        orderer_group = QGroupBox("受注者情報")
        orderer_layout = QVBoxLayout()
        
        # 対応者名
        orderer_layout.addWidget(QLabel("対応者名"))
        self.operator_input = QLineEdit()
        if orderer_data and 'operator' in orderer_data:
            self.operator_input.setText(orderer_data['operator'])
        orderer_layout.addWidget(self.operator_input)
        
        # 出やすい時間帯
        orderer_layout.addWidget(QLabel("出やすい時間帯"))
        self.available_time_input = QLineEdit()
        self.available_time_input.setPlaceholderText("AMPM希望　固定or携帯　000-0000-0000")
        if orderer_data and 'available_time' in orderer_data:
            self.available_time_input.setText(orderer_data['available_time'])
        orderer_layout.addWidget(self.available_time_input)
        
        # 契約者名
        orderer_layout.addWidget(QLabel("契約者名"))
        self.contractor_input = QLineEdit()
        if orderer_data and 'contractor' in orderer_data:
            self.contractor_input.setText(orderer_data['contractor'])
        orderer_layout.addWidget(self.contractor_input)
        
        # フリガナ
        furigana_layout = QHBoxLayout()
        furigana_layout.addWidget(QLabel("フリガナ"))
        self.furigana_mode_combo = QComboBox()
        self.furigana_mode_combo.addItems(["自動", "手動"])
        furigana_layout.addWidget(self.furigana_mode_combo)
        orderer_layout.addLayout(furigana_layout)
        self.furigana_input = QLineEdit()
        if orderer_data and 'furigana' in orderer_data:
            self.furigana_input.setText(orderer_data['furigana'])
        orderer_layout.addWidget(self.furigana_input)
        
        # 生年月日
        birth_layout = QVBoxLayout()
        birth_layout.addWidget(QLabel("生年月日"))
        
        # 生年月日の入力部分を横並びにする
        birth_input_layout = QHBoxLayout()
        
        self.era_combo = QComboBox()
        self.era_combo.addItems(["昭和", "平成", "西暦"])
        self.era_combo.setFixedWidth(60)
        birth_input_layout.addWidget(self.era_combo)
        
        self.year_combo = QComboBox()
        self.year_combo.addItems([str(i) for i in range(1, 65)])
        self.year_combo.setEditable(True)
        self.year_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.year_combo.lineEdit().setMaxLength(4)
        self.year_combo.lineEdit().setValidator(QIntValidator(1, 9999))
        self.year_combo.setFixedWidth(60)
        birth_input_layout.addWidget(self.year_combo)
        birth_input_layout.addWidget(QLabel("年"))
        
        self.month_combo = QComboBox()
        self.month_combo.addItems([str(i) for i in range(1, 13)])
        self.month_combo.setEditable(True)
        self.month_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.month_combo.lineEdit().setMaxLength(2)
        self.month_combo.lineEdit().setValidator(QIntValidator(1, 12))
        self.month_combo.setFixedWidth(40)
        birth_input_layout.addWidget(self.month_combo)
        birth_input_layout.addWidget(QLabel("月"))
        
        self.day_combo = QComboBox()
        self.day_combo.addItems([str(i) for i in range(1, 32)])
        self.day_combo.setEditable(True)
        self.day_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.day_combo.lineEdit().setMaxLength(2)
        self.day_combo.lineEdit().setValidator(QIntValidator(1, 31))
        self.day_combo.setFixedWidth(40)
        birth_input_layout.addWidget(self.day_combo)
        birth_input_layout.addWidget(QLabel("日"))
        
        birth_layout.addLayout(birth_input_layout)
        orderer_layout.addLayout(birth_layout)
        
        # 受注者名
        orderer_layout.addWidget(QLabel("受注者名"))
        self.order_person_input = QLineEdit()
        if orderer_data and 'order_person' in orderer_data:
            self.order_person_input.setText(orderer_data['order_person'])
        orderer_layout.addWidget(self.order_person_input)
        
        # 社番
        orderer_layout.addWidget(QLabel("社番"))
        self.employee_number_input = QLineEdit()
        if orderer_data and 'employee_number' in orderer_data:
            self.employee_number_input.setText(orderer_data['employee_number'])
        orderer_layout.addWidget(self.employee_number_input)
        
        # 料金認識
        orderer_layout.addWidget(QLabel("料金認識"))
        self.fee_input = QLineEdit()
        self.fee_input.setText("2500円～3000円")
        if orderer_data and 'fee' in orderer_data:
            self.fee_input.setText(orderer_data['fee'])
        orderer_layout.addWidget(self.fee_input)
        
        # ネット利用
        orderer_layout.addWidget(QLabel("ネット利用"))
        self.net_usage_combo = QComboBox()
        self.net_usage_combo.addItems(["なし", "あり"])
        if orderer_data and 'net_usage' in orderer_data:
            self.net_usage_combo.setCurrentText(orderer_data['net_usage'])
        orderer_layout.addWidget(self.net_usage_combo)
        
        # 家族了承
        orderer_layout.addWidget(QLabel("家族了承"))
        self.family_approval_combo = QComboBox()
        self.family_approval_combo.addItems(["ok", "なし"])
        if orderer_data and 'family_approval' in orderer_data:
            self.family_approval_combo.setCurrentText(orderer_data['family_approval'])
        orderer_layout.addWidget(self.family_approval_combo)
        
        # 他番号
        orderer_layout.addWidget(QLabel("他番号"))
        self.other_number_input = QLineEdit()
        self.other_number_input.setText("なし")
        if orderer_data and 'other_number' in orderer_data:
            self.other_number_input.setText(orderer_data['other_number'])
        orderer_layout.addWidget(self.other_number_input)
        
        # 電話機
        orderer_layout.addWidget(QLabel("電話機"))
        self.phone_device_input = QLineEdit()
        self.phone_device_input.setText("プッシュホン")
        if orderer_data and 'phone_device' in orderer_data:
            self.phone_device_input.setText(orderer_data['phone_device'])
        orderer_layout.addWidget(self.phone_device_input)
        
        # 禁止回線
        orderer_layout.addWidget(QLabel("禁止回線"))
        self.forbidden_line_input = QLineEdit()
        self.forbidden_line_input.setText("なし")
        if orderer_data and 'forbidden_line' in orderer_data:
            self.forbidden_line_input.setText(orderer_data['forbidden_line'])
        orderer_layout.addWidget(self.forbidden_line_input)
        
        # ND
        orderer_layout.addWidget(QLabel("ND"))
        self.nd_input = QLineEdit()
        if orderer_data and 'nd' in orderer_data:
            self.nd_input.setText(orderer_data['nd'])
        orderer_layout.addWidget(self.nd_input)
        
        # リストとの関係性
        relationship_layout = QHBoxLayout()
        relationship_layout.addWidget(QLabel("備考："))
        self.relationship_input = QLineEdit()
        self.relationship_input.setPlaceholderText("名義人の...")
        if orderer_data and 'relationship' in orderer_data:
            self.relationship_input.setText(orderer_data['relationship'])
        relationship_layout.addWidget(self.relationship_input)
        orderer_layout.addLayout(relationship_layout)
        
        orderer_group.setLayout(orderer_layout)
        scroll_layout.addWidget(orderer_group)
        
        # スクロールエリアにウィジェットを設定
        scroll_area.setWidget(scroll_widget)
        layout.addWidget(scroll_area)
        
        # ボタンエリア
        button_layout = QHBoxLayout()
        
        # 戻るボタン
        self.back_btn = QPushButton("戻る")
        self.back_btn.setStyleSheet("""
            QPushButton {
                background-color: #757575;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #616161;
            }
            QPushButton:pressed {
                background-color: #424242;
            }
        """)
        button_layout.addWidget(self.back_btn)
        
        # 次へボタン
        self.next_btn = QPushButton("次へ")
        self.next_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #1565C0;
            }
        """)
        button_layout.addWidget(self.next_btn)
        
        # 作成中止ボタン
        self.cancel_btn = QPushButton("作成中止")
        self.cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
            QPushButton:pressed {
                background-color: #c62828;
            }
        """)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
        # シグナルの接続
        self.back_btn.clicked.connect(self.reject)
        self.next_btn.clicked.connect(self.accept)
        self.cancel_btn.clicked.connect(self.reject)
    
    def get_orderer_data(self):
        """
        受注者情報を取得する
        
        Returns:
            dict: 受注者情報
        """
        try:
            # 生年月日を西暦のYYYY/MM/DD形式で生成
            era = self.era_combo.currentText()
            year = int(self.year_combo.currentText())
            month = int(self.month_combo.currentText())
            day = int(self.day_combo.currentText())
            
            # 和暦を西暦に変換
            if era == "昭和":
                year = year + 1925
            elif era == "平成":
                year = year + 1988
            # 西暦の場合はそのまま
            
            # 月と日を2桁の文字列に変換
            month_str = str(month).zfill(2)
            day_str = str(day).zfill(2)
            
            # YYYY/MM/DD形式で生成
            birth_date = f"{year}/{month_str}/{day_str}"
            
            # データを辞書形式で返す
            data = {
                'operator': self.operator_input.text(),
                'available_time': self.available_time_input.text(),
                'contractor': self.contractor_input.text(),
                'furigana': self.furigana_input.text(),
                'birth_date': birth_date,  # 生年月日を設定
                'order_person': self.order_person_input.text(),
                'employee_number': self.employee_number_input.text(),
                'fee': self.fee_input.text(),
                'net_usage': self.net_usage_combo.currentText(),
                'family_approval': self.family_approval_combo.currentText(),
                'other_number': self.other_number_input.text(),
                'phone_device': self.phone_device_input.text(),
                'forbidden_line': self.forbidden_line_input.text(),
                'nd': self.nd_input.text(),
                'relationship': self.relationship_input.text()
            }
            
            logging.info(f"受注者データを取得: {data}")
            return data
            
        except Exception as e:
            logging.error(f"受注者データの取得中にエラー: {e}")
            return {}

    def set_orderer_data(self, data):
        """
        受注者情報を設定する
        
        Args:
            data: 受注者情報データ
        """
        try:
            # 対応者名
            if data.get('operator'):
                converted_operator = convert_to_half_width(data['operator'])
                self.operator_input.setText(converted_operator)
            
            # 出やすい時間帯
            if data.get('available_time'):
                converted_time = convert_to_half_width(data['available_time'])
                self.available_time_input.setText(converted_time)
            
            # 契約者名
            if data.get('contractor'):
                converted_contractor = convert_to_half_width(data['contractor'])
                self.contractor_input.setText(converted_contractor)
            
            # フリガナ
            if data.get('furigana'):
                converted_furigana = convert_to_half_width(data['furigana'])
                self.furigana_input.setText(converted_furigana)
            
            # 生年月日
            if data.get('era'):
                self.era_combo.setCurrentText(data['era'])
            if data.get('year'):
                converted_year = convert_to_half_width(data['year'])
                self.year_combo.setCurrentText(converted_year)
            if data.get('month'):
                converted_month = convert_to_half_width(data['month'])
                self.month_combo.setCurrentText(converted_month)
            if data.get('day'):
                converted_day = convert_to_half_width(data['day'])
                self.day_combo.setCurrentText(converted_day)
            
            # 受注者名
            if data.get('order_person'):
                converted_person = convert_to_half_width(data['order_person'])
                self.order_person_input.setText(converted_person)
            
            # 社番
            if data.get('employee_number'):
                converted_number = convert_to_half_width(data['employee_number'])
                self.employee_number_input.setText(converted_number)
            
            # 料金認識
            if data.get('fee'):
                converted_fee = convert_to_half_width(data['fee'])
                self.fee_input.setText(converted_fee)
            
            # ネット利用
            if data.get('net_usage'):
                self.net_usage_combo.setCurrentText(data['net_usage'])
            
            # 家族了承
            if data.get('family_approval'):
                self.family_approval_combo.setCurrentText(data['family_approval'])
            
            # 他番号
            if data.get('other_number'):
                converted_other = convert_to_half_width(data['other_number'])
                self.other_number_input.setText(converted_other)
            
            # 電話機
            if data.get('phone_device'):
                converted_device = convert_to_half_width(data['phone_device'])
                self.phone_device_input.setText(converted_device)
            
            # 禁止回線
            if data.get('forbidden_line'):
                converted_forbidden = convert_to_half_width(data['forbidden_line'])
                self.forbidden_line_input.setText(converted_forbidden)
            
            # ND
            if data.get('nd'):
                converted_nd = convert_to_half_width(data['nd'])
                self.nd_input.setText(converted_nd)
            
            # リストとの関係性
            if data.get('relationship'):
                converted_relationship = convert_to_half_width(data['relationship'])
                self.relationship_input.setText(converted_relationship)
            
            logging.info("受注者データを正常に設定しました")
            
        except Exception as e:
            logging.error(f"受注者情報の設定中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"受注者情報の設定中にエラーが発生しました: {e}")

class OrderInfoDialog(QDialog):
    """受注情報入力ダイアログ"""
    
    def __init__(self, parent=None, order_data=None):
        """
        受注情報入力ダイアログの初期化
        
        Args:
            parent: 親ウィジェット
            order_data: 受注情報データ
        """
        super().__init__(parent)
        self.setWindowTitle("受注情報入力")
        self.setModal(True)
        self.setMinimumWidth(500)
        
        # メインレイアウト
        layout = QVBoxLayout()
        
        # タイトル
        title_label = QLabel("受注情報の確認")
        title_label.setFont(QFont("MS Gothic", 12, QFont.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)
        
        # 受注情報入力グループ
        order_group = QGroupBox("受注情報")
        order_layout = QVBoxLayout()
        
        # 現状回線
        order_layout.addWidget(QLabel("現状回線"))
        self.current_line_combo = QComboBox()
        self.current_line_combo.addItems(["アナログ"])
        if order_data and 'current_line' in order_data:
            self.current_line_combo.setCurrentText(order_data['current_line'])
        order_layout.addWidget(self.current_line_combo)
        
        # 受注日（本日自動入力）
        order_layout.addWidget(QLabel("受注日"))
        self.order_date_input = QLineEdit()
        # 0埋めなしの月/日フォーマットを生成
        now = datetime.datetime.now()
        month = str(now.month)  # 0埋めなしの月
        day = str(now.day)      # 0埋めなしの日
        self.order_date_input.setText(f"{month}/{day}")
        self.order_date_input.setReadOnly(True)
        order_layout.addWidget(self.order_date_input)
        
        # 提供判定
        order_layout.addWidget(QLabel("提供判定"))
        self.judgment_combo = QComboBox()
        self.judgment_combo.addItems(["OK", "NG"])
        if order_data and 'judgment' in order_data:
            self.judgment_combo.setCurrentText(order_data['judgment'])
        order_layout.addWidget(self.judgment_combo)
        
        order_group.setLayout(order_layout)
        layout.addWidget(order_group)
        
        # ボタンエリア
        button_layout = QHBoxLayout()
        
        # 戻るボタン
        self.back_btn = QPushButton("戻る")
        self.back_btn.setStyleSheet("""
            QPushButton {
                background-color: #757575;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #616161;
            }
            QPushButton:pressed {
                background-color: #424242;
            }
        """)
        button_layout.addWidget(self.back_btn)
        
        # 営コメ作成ボタン
        self.create_comment_btn = QPushButton("営コメ作成")
        self.create_comment_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3e8e41;
            }
        """)
        button_layout.addWidget(self.create_comment_btn)
        
        # 作成中止ボタン
        self.cancel_btn = QPushButton("作成中止")
        self.cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
            QPushButton:pressed {
                background-color: #c62828;
            }
        """)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
        # シグナルの接続
        self.back_btn.clicked.connect(self.reject)
        self.create_comment_btn.clicked.connect(self.create_comment)
        self.cancel_btn.clicked.connect(self.reject)
    
    def get_order_data(self):
        """
        入力された受注情報を取得
        
        Returns:
            dict: 受注情報
        """
        return {
            'current_line': self.current_line_combo.currentText(),
            'order_date': self.order_date_input.text(),
            'judgment': self.judgment_combo.currentText()
        }
    
    def create_comment(self):
        """
        営コメを作成し、プレビューに表示する
        """
        try:
            logging.info("営コメ作成を開始")
            
            # 親ウィンドウのメソッドを呼び出して営コメを作成
            if hasattr(self.parent(), 'generate_preview_text'):
                logging.info("親ウィンドウのgenerate_preview_textメソッドを呼び出し")
                
                # 現在のダイアログのデータを保存
                self.parent().current_dialog = self
                logging.info("現在のダイアログデータを親ウィンドウに保存")
                
                # プレビューテキストを生成
                preview_text = self.parent().generate_preview_text()
                if preview_text:
                    logging.info("プレビューテキストの生成に成功")
                    # プレビューに表示
                    self.parent().preview_text.setText(preview_text)
                    logging.info("プレビューにテキストを表示")
                    self.accept()
                else:
                    logging.error("プレビューテキストの生成に失敗")
                    QMessageBox.warning(self, "警告", "営コメの作成に失敗しました。")
            else:
                logging.error("親ウィンドウにgenerate_preview_textメソッドが存在しません")
                QMessageBox.warning(self, "警告", "プレビュー機能が利用できません。")
        except Exception as e:
            logging.error(f"営コメ作成中にエラー: {e}", exc_info=True)
            QMessageBox.critical(self, "エラー", f"営コメの作成中にエラーが発生しました: {e}") 