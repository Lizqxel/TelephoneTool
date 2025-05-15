"""
住所情報入力ダイアログと関連コンポーネント

このモジュールは、簡易モードの住所情報入力ダイアログを提供します。
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
                              QLabel, QLineEdit, QPushButton,
                              QGroupBox, QMessageBox, QWidget, QComboBox, QScrollArea)
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtCore import Qt, QThread, Signal, QEvent, QMetaObject, Q_ARG, QTimer, QPoint, QUrl, QObject
from PySide6.QtGui import QFont, QIntValidator, QPalette, QColor
import datetime
import logging
from services.area_search import search_service_area, normalize_address
import threading
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Slot

# 高齢者向けのスタイル設定
LARGE_FONT = QFont("MS Gothic", 16)
MEDIUM_FONT = QFont("MS Gothic", 14)
SMALL_FONT = QFont("MS Gothic", 12)

BUTTON_STYLE = """
    QPushButton {
        background-color: #4CAF50;
        color: white;
        border: none;
        padding: 15px 30px;
        font-size: 16px;
        border-radius: 8px;
        min-width: 200px;
    }
    QPushButton:hover {
        background-color: #45a049;
    }
    QPushButton:pressed {
        background-color: #3e8e41;
    }
    QPushButton:disabled {
        background-color: #cccccc;
    }
"""

GROUP_BOX_STYLE = """
    QGroupBox {
        font-size: 16px;
        font-weight: bold;
        border: 2px solid #4CAF50;
        border-radius: 8px;
        margin-top: 1em;
        padding-top: 1em;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 3px 0 3px;
    }
"""

LABEL_STYLE = """
    QLabel {
        font-size: 16px;
        color: #333333;
        padding: 5px;
    }
"""

INPUT_STYLE = """
    QLineEdit, QComboBox {
        font-size: 16px;
        padding: 10px;
        border: 2px solid #cccccc;
        border-radius: 4px;
        min-height: 40px;
    }
    QLineEdit:focus, QComboBox:focus {
        border: 2px solid #4CAF50;
    }
"""

# ダイアログの戻り値定数
DIALOG_BACK = -1
DIALOG_NEXT = 1
DIALOG_CANCEL = 0

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

class NoWheelComboBox(QComboBox):
    """スクロールイベントを無視するQComboBox"""
    def wheelEvent(self, event):
        event.ignore()

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
        self.setMinimumWidth(600)
        
        # 検索スレッドの初期化
        self.search_thread = None
        self.parent_window = parent
        
        # 保存データの初期化
        self.saved_data = address_data or {}
        
        # メインレイアウト
        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # タイトル
        title_label = QLabel("住所情報の確認")
        title_label.setFont(LARGE_FONT)
        title_label.setStyleSheet("color: #2E7D32; font-weight: bold;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)
        
        # 住所情報入力グループ
        address_group = QGroupBox("住所情報")
        address_group.setStyleSheet(GROUP_BOX_STYLE)
        address_layout = QVBoxLayout()
        address_layout.setSpacing(15)
        
        # 郵便番号
        postal_label = QLabel("郵便番号")
        postal_label.setFont(MEDIUM_FONT)
        postal_label.setStyleSheet(LABEL_STYLE)
        address_layout.addWidget(postal_label)
        
        self.postal_code_input = QLineEdit()
        self.postal_code_input.setFont(MEDIUM_FONT)
        self.postal_code_input.setStyleSheet(INPUT_STYLE)
        self.postal_code_input.setPlaceholderText("例：123-4567")
        if self.saved_data and 'postal_code' in self.saved_data:
            self.postal_code_input.setText(self.saved_data['postal_code'])
        address_layout.addWidget(self.postal_code_input)
        
        # 住所
        address_label = QLabel("住所")
        address_label.setFont(MEDIUM_FONT)
        address_label.setStyleSheet(LABEL_STYLE)
        address_layout.addWidget(address_label)
        
        self.address_input = QLineEdit()
        self.address_input.setFont(MEDIUM_FONT)
        self.address_input.setStyleSheet(INPUT_STYLE)
        self.address_input.setPlaceholderText("例：東京都渋谷区...")
        if self.saved_data and 'address' in self.saved_data:
            self.address_input.setText(self.saved_data['address'])
        address_layout.addWidget(self.address_input)
        
        address_group.setLayout(address_layout)
        layout.addWidget(address_group)
        
        # 提供判定ボタン
        self.judgment_btn = QPushButton("提供判定実行")
        self.judgment_btn.setFont(LARGE_FONT)
        self.judgment_btn.setStyleSheet(BUTTON_STYLE)
        self.judgment_btn.clicked.connect(self.search_area)
        layout.addWidget(self.judgment_btn, alignment=Qt.AlignmentFlag.AlignCenter)
        
        # 提供判定結果表示ラベル
        self.judgment_result = QLabel("提供エリア: 未検索")
        self.judgment_result.setFont(MEDIUM_FONT)
        self.judgment_result.setStyleSheet("""
            QLabel {
                font-size: 16px;
                padding: 15px;
                border: 2px solid #dddddd;
                border-radius: 8px;
                background-color: #f8f9fa;
            }
        """)
        layout.addWidget(self.judgment_result)
        
        # ボタンエリア
        button_layout = QHBoxLayout()
        button_layout.setSpacing(20)
        
        # 次へボタン
        self.next_btn = QPushButton("次へ")
        self.next_btn.setFont(LARGE_FONT)
        self.next_btn.setStyleSheet(BUTTON_STYLE)
        self.next_btn.clicked.connect(self.on_next_clicked)
        button_layout.addWidget(self.next_btn)
        
        # 作成中止ボタン
        self.cancel_btn = QPushButton("作成中止")
        self.cancel_btn.setFont(LARGE_FONT)
        self.cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 15px 30px;
                font-size: 16px;
                border-radius: 8px;
                min-width: 200px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
            QPushButton:pressed {
                background-color: #c62828;
            }
        """)
        self.cancel_btn.clicked.connect(self.on_cancel_clicked)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def on_next_clicked(self):
        """次へボタンがクリックされた時の処理"""
        # 現在の入力内容を保存
        self.saved_data = self.get_address_data()
        self.accept()

    def get_saved_data(self):
        """保存されたデータを取得"""
        return self.saved_data

    def get_address_data(self):
        """
        住所情報を取得する
        
        Returns:
            dict: 住所情報
        """
        try:
            # 入力値を取得
            input_postal_code = self.postal_code_input.text()
            input_address = self.address_input.text()
            
            # 全角文字を半角に変換
            normalized_postal_code = normalize_address(input_postal_code)
            normalized_address = normalize_address(input_address)
            
            # データを辞書形式で返す
            data = {
                'postal_code': normalized_postal_code,
                'address': normalized_address,
                'original_postal_code': input_postal_code,  # 元の入力値も保持
                'original_address': input_address
            }
            
            logging.info(f"住所データを取得: {data}")
            return data
            
        except Exception as e:
            logging.error(f"住所データの取得中にエラー: {e}")
            return {}
    
    def search_area(self):
        """
        提供エリアを検索する
        """
        try:
            # 住所データを取得
            data = self.get_address_data()
            if not data:
                return
            
            # 元の入力値を使用して検索
            postal_code = data['original_postal_code']
            address = data['original_address']
            
            # 検索を実行
            result = search_service_area(postal_code, address)
            logging.info(f"★★★ 検索結果: {result} ★★★")
            
            # 結果を表示
            self.show_result(result)
            
        except Exception as e:
            logging.error(f"エリア検索中にエラー: {e}")
            self.show_error("エリア検索中にエラーが発生しました。")

    def on_search_finished(self, result):
        """検索完了時の処理"""
        
        # 追加ログ：resultの内容を詳細に出力
        logging.info(f"[UI] on_search_completed 受信 result: {result}")
        logging.info(f"[UI] on_search_completed 受信 status: {result.get('status', 'キーなし')}, message: {result.get('message', 'キーなし')}")
        try:
            logging.info(f"★★★ 検索完了: {result} ★★★")
            
            # 検索結果のステータスを取得
            status = result.get("status", "failure")
            message = result.get("message", "判定失敗")
            logging.info(f"★★★ 検索結果のステータス: {status}, メッセージ: {message} ★★★")
            
            # 判定結果テキストと表示スタイルを設定
            if status == "available":
                result_text = "提供エリア: 提供可能"
                style = """
                    QLabel {
                        font-size: 14px;
                        padding: 10px;
                        border: 1px solid #27AE60;
                        border-radius: 4px;
                        background-color: #E8F5E9;
                        color: #27AE60;
                    }
                """
            elif status == "unavailable":
                result_text = "提供エリア: 提供エリア外"
                style = """
                    QLabel {
                        font-size: 14px;
                        padding: 10px;
                        border: 1px solid #E74C3C;
                        border-radius: 4px;
                        background-color: #FFEBEE;
                        color: #E74C3C;
                    }
                """
            elif status == "apartment":
                result_text = "提供エリア: 集合住宅（アパート・マンション等）"
                style = """
                    QLabel {
                        font-size: 14px;
                        padding: 10px;
                        border: 1px solid #388E3C;
                        border-radius: 4px;
                        background-color: #C8E6C9;
                        color: #388E3C;
                    }
                """
            else:
                result_text = "提供エリア: 判定失敗"
                style = """
                    QLabel {
                        font-size: 14px;
                        padding: 10px;
                        border: 1px solid #95A5A6;
                        border-radius: 4px;
                        background-color: #ECEFF1;
                        color: #95A5A6;
                    }
                """

            # 判定結果ラベルを更新
            self.judgment_result.setText(result_text)
            self.judgment_result.setStyleSheet(style)
            logging.info(f"★★★ 判定結果ラベルを更新: {result_text} ★★★")

            # 親ウィンドウの判定結果も更新
            if self.parent_window:
                logging.info(f"★★★ 親ウィンドウの判定結果を更新: {result_text} ★★★")
                self.parent_window.update_judgment_result(result_text)

            # 検索ボタンを有効化
            self.judgment_btn.setEnabled(True)
            self.judgment_btn.setText("検索")

            # 詳細情報がある場合は表示
            if "details" in result:
                details = result["details"]
                logging.info(f"★★★ 詳細情報: {details} ★★★")

        except Exception as e:
            logging.error(f"★★★ 検索結果の処理でエラー: {e} ★★★", exc_info=True)
            self.on_search_error(str(e))

    def on_search_error(self, error_message):
        """検索エラー時の処理"""
        self.judgment_result.setText("提供エリア: 検索エラー")
        self.judgment_result.setStyleSheet("""
            QLabel {
                font-size: 14px;
                padding: 5px;
                border: 1px solid #f44336;
                border-radius: 4px;
                background-color: #FFEBEE;
                color: #B71C1C;
            }
        """)
        QMessageBox.critical(self, "エラー", f"提供エリアの検索中にエラーが発生しました: {error_message}")

    def on_cancel_clicked(self):
        """作成中止ボタンがクリックされた時の処理"""
        # 確認ダイアログを表示
        reply = QMessageBox.question(
            self,
            "確認",
            "作成を中止しますか？\n入力したデータは保存されません。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            # メイン画面に戻る
            self.done(QDialog.DialogCode.Rejected)
        else:
            # キャンセルをキャンセル
            pass

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
        
        # 親ウィンドウへの参照を保持
        self.parent_window = parent
        
        # 保存データの初期化
        self.saved_data = list_data or {}
        
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
        list_layout.addWidget(self.list_name_input)
        
        # リストフリガナ（自動・手動切り替え機能付き）
        list_furigana_layout = QHBoxLayout()
        list_furigana_layout.addWidget(QLabel("リストフリガナ"))
        self.list_furigana_mode_combo = NoWheelComboBox()
        self.list_furigana_mode_combo.addItems(["自動", "手動"])
        list_furigana_layout.addWidget(self.list_furigana_mode_combo)
        list_layout.addLayout(list_furigana_layout)
        
        self.list_furigana_input = QLineEdit()
        self.list_furigana_input.setPlaceholderText("例：タナカタロウ")
        list_layout.addWidget(self.list_furigana_input)
        
        # 電話番号
        list_layout.addWidget(QLabel("電話番号"))
        self.list_phone_input = QLineEdit()
        self.list_phone_input.setPlaceholderText("例：090-1234-5678")
        if self.saved_data and 'list_phone' in self.saved_data:
            self.list_phone_input.setText(self.saved_data['list_phone'])
        list_layout.addWidget(self.list_phone_input)
        
        # リスト郵便番号
        list_layout.addWidget(QLabel("リスト郵便番号"))
        self.list_postal_code_input = QLineEdit()
        self.list_postal_code_input.setPlaceholderText("例：123-4567")
        if self.saved_data and 'list_postal_code' in self.saved_data:
            self.list_postal_code_input.setText(self.saved_data['list_postal_code'])
        list_layout.addWidget(self.list_postal_code_input)
        
        # リスト住所
        list_layout.addWidget(QLabel("リスト住所"))
        self.list_address_input = QLineEdit()
        self.list_address_input.setPlaceholderText("例：東京都渋谷区...")
        if self.saved_data and 'list_address' in self.saved_data:
            self.list_address_input.setText(self.saved_data['list_address'])
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
        self.cancel_btn.clicked.connect(self.on_cancel_clicked)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
        # シグナルの接続
        self.back_btn.clicked.connect(self.on_back_clicked)
        self.next_btn.clicked.connect(self.on_next_clicked)
        
        # ここでデータをセットしてから、シグナルを接続する（先にデータセット）
        if self.saved_data and 'list_name' in self.saved_data:
            self.list_name_input.setText(self.saved_data['list_name'])
        
        if self.saved_data and 'list_furigana' in self.saved_data:
            self.list_furigana_input.setText(self.saved_data['list_furigana'])
        
        # リスト名入力時のフリガナ自動生成のシグナルを接続
        self.list_name_input.textChanged.connect(self.auto_generate_list_furigana)
        # コンボボックス変更時にも自動生成を試行
        self.list_furigana_mode_combo.currentTextChanged.connect(lambda: self.auto_generate_list_furigana())
        
        # 初期表示時に一度だけ自動生成を実行
        QTimer.singleShot(100, self.auto_generate_list_furigana)
    
    def on_back_clicked(self):
        """戻るボタンがクリックされた時の処理"""
        # 現在の入力内容を保存
        self.saved_data = self.get_list_data()
        # 一つ前のダイアログに戻る
        self.done(DIALOG_BACK)

    def on_next_clicked(self):
        """次へボタンがクリックされた時の処理"""
        # 現在の入力内容を保存
        self.saved_data = self.get_list_data()
        self.accept()

    def get_saved_data(self):
        """保存されたデータを取得"""
        return self.saved_data

    def auto_generate_list_furigana(self):
        """リスト名からフリガナを自動生成する"""
        try:
            # 自動モードの場合のみ処理
            if self.list_furigana_mode_combo.currentText() != "自動":
                logging.info("フリガナ自動生成: 手動モードのため生成をスキップします")
                return
                
            # リスト名が空の場合は何もしない
            name = self.list_name_input.text()
            if not name:
                logging.info("フリガナ自動生成: リスト名が空のため生成をスキップします")
                return
                
            # 既にフリガナが入力されている場合はスキップする特殊なケース
            current_furigana = self.list_furigana_input.text()
            if current_furigana and len(current_furigana) > 1 and name in self.saved_data.get('list_name', ''):
                logging.info(f"フリガナ自動生成: 既にフリガナ({current_furigana})が設定されているためスキップします")
                return
            
            # フリガナ変換APIを使用
            logging.info(f"フリガナ自動生成: 変換を開始します（リスト名: {name}）")
            from utils.furigana_utils import convert_to_furigana
            furigana = convert_to_furigana(name)
            
            if furigana:
                # blockSignalsでシグナルを一時的にブロックして再帰呼び出しを防止
                self.list_furigana_input.blockSignals(True)
                self.list_furigana_input.setText(furigana)
                self.list_furigana_input.blockSignals(False)
                logging.info(f"フリガナ自動生成: 成功 - {name} → {furigana}")
            else:
                logging.warning(f"フリガナ自動生成: 変換APIから結果が返ってきませんでした（リスト名: {name}）")
        
        except Exception as e:
            logging.error(f"リストフリガナ自動生成エラー: {str(e)}", exc_info=True)

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

    def on_cancel_clicked(self):
        """作成中止ボタンがクリックされた時の処理"""
        # 確認ダイアログを表示
        reply = QMessageBox.question(
            self,
            "確認",
            "作成を中止しますか？\n入力したデータは保存されません。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            # メイン画面に戻る
            self.done(QDialog.DialogCode.Rejected)
        else:
            # キャンセルをキャンセル
            pass

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
        self.setWindowTitle("受注者情報入力")
        self.setModal(True)
        
        # 親ウィンドウへの参照を保持
        self.parent_window = parent
        
        # 保存データの初期化
        self.saved_data = orderer_data if orderer_data is not None else {}
        
        # フリガナ自動生成用のタイマー
        self.furigana_timer = QTimer()
        self.furigana_timer.setSingleShot(True)
        self.furigana_timer.timeout.connect(self.auto_generate_furigana)
        
        # 最後に変換した契約者名を保持
        self.last_converted_name = ""
        
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
        orderer_group.setStyleSheet(GROUP_BOX_STYLE)
        orderer_layout = QVBoxLayout()
        orderer_layout.setSpacing(15)
        
        # 対応者名
        orderer_layout.addWidget(QLabel("対応者名"))
        self.operator_input = QLineEdit()
        self.operator_input.textChanged.connect(self.check_input_fields)
        orderer_layout.addWidget(self.operator_input)
        
        # 出やすい時間帯
        orderer_layout.addWidget(QLabel("出やすい時間帯"))
        self.available_time_input = QLineEdit()
        self.available_time_input.setPlaceholderText("AMPM希望　固定or携帯　000-0000-0000")
        self.available_time_input.textChanged.connect(self.check_input_fields)
        orderer_layout.addWidget(self.available_time_input)
        
        # 契約者名
        orderer_layout.addWidget(QLabel("契約者名"))
        self.contractor_input = QLineEdit()
        # リストから取得した顧客名を初期値として設定
        if parent and hasattr(parent, 'list_data') and 'list_name' in parent.list_data:
            self.contractor_input.setText(parent.list_data['list_name'])
        self.contractor_input.textChanged.connect(self.check_input_fields)
        # 契約者名の変更時にフリガナ自動生成を実行（タイマーを使用）
        self.contractor_input.textChanged.connect(self.schedule_furigana_generation)
        orderer_layout.addWidget(self.contractor_input)
        
        # フリガナ
        furigana_layout = QHBoxLayout()
        furigana_layout.addWidget(QLabel("フリガナ"))
        self.furigana_mode_combo = QComboBox()
        self.furigana_mode_combo.addItems(["自動", "手動"])
        furigana_layout.addWidget(self.furigana_mode_combo)
        orderer_layout.addLayout(furigana_layout)
        self.furigana_input = QLineEdit()
        self.furigana_input.textChanged.connect(self.check_input_fields)
        orderer_layout.addWidget(self.furigana_input)
        
        # 生年月日
        birth_layout = QVBoxLayout()
        birth_layout.addWidget(QLabel("生年月日"))
        
        # 生年月日の入力部分を横並びにする
        birth_input_layout = QHBoxLayout()
        
        self.era_combo = NoWheelComboBox()
        self.era_combo.addItems(["昭和", "平成", "西暦"])
        self.era_combo.setFixedWidth(60)
        self.era_combo.currentTextChanged.connect(self.check_birth_date_age)
        birth_input_layout.addWidget(self.era_combo)
        
        self.year_combo = NoWheelComboBox()
        self.year_combo.addItems([str(i) for i in range(1, 65)])
        self.year_combo.setEditable(True)
        self.year_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.year_combo.lineEdit().setMaxLength(4)
        self.year_combo.lineEdit().setValidator(QIntValidator(1, 9999))
        self.year_combo.setFixedWidth(60)
        self.year_combo.currentTextChanged.connect(self.check_birth_date_age)
        birth_input_layout.addWidget(self.year_combo)
        birth_input_layout.addWidget(QLabel("年"))
        
        self.month_combo = NoWheelComboBox()
        self.month_combo.addItems([str(i) for i in range(1, 13)])
        self.month_combo.setEditable(True)
        self.month_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.month_combo.lineEdit().setMaxLength(2)
        self.month_combo.lineEdit().setValidator(QIntValidator(1, 12))
        self.month_combo.setFixedWidth(40)
        self.month_combo.currentTextChanged.connect(self.check_birth_date_age)
        birth_input_layout.addWidget(self.month_combo)
        birth_input_layout.addWidget(QLabel("月"))
        
        self.day_combo = NoWheelComboBox()
        self.day_combo.addItems([str(i) for i in range(1, 32)])
        self.day_combo.setEditable(True)
        self.day_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.day_combo.lineEdit().setMaxLength(2)
        self.day_combo.lineEdit().setValidator(QIntValidator(1, 31))
        self.day_combo.setFixedWidth(40)
        self.day_combo.currentTextChanged.connect(self.check_birth_date_age)
        birth_input_layout.addWidget(self.day_combo)
        birth_input_layout.addWidget(QLabel("日"))
        
        birth_layout.addLayout(birth_input_layout)
        orderer_layout.addLayout(birth_layout)
        
        # 受注者名
        orderer_layout.addWidget(QLabel("受注者名"))
        self.order_person_input = QLineEdit()
        self.order_person_input.textChanged.connect(self.check_input_fields)
        orderer_layout.addWidget(self.order_person_input)
        
        # 料金認識
        orderer_layout.addWidget(QLabel("料金認識"))
        fee_layout = QHBoxLayout()
        self.fee_combo = NoWheelComboBox()
        self.fee_combo.addItems(["2500円～3000円", "3500円～4000円"])
        self.fee_combo.currentTextChanged.connect(self.on_fee_combo_changed)
        fee_layout.addWidget(self.fee_combo)
        self.fee_input = QLineEdit()
        self.fee_input.setPlaceholderText("手動入力")
        self.fee_input.textChanged.connect(self.check_input_fields)
        fee_layout.addWidget(self.fee_input)
        orderer_layout.addLayout(fee_layout)
        
        # ネット利用
        orderer_layout.addWidget(QLabel("ネット利用"))
        self.net_usage_input = QLineEdit()
        self.net_usage_input.setText("なし")
        self.net_usage_input.textChanged.connect(self.check_input_fields)
        orderer_layout.addWidget(self.net_usage_input)
        
        # 家族了承
        orderer_layout.addWidget(QLabel("家族了承"))
        self.family_approval_input = QLineEdit()
        self.family_approval_input.setText("ok")
        self.family_approval_input.textChanged.connect(self.check_input_fields)
        orderer_layout.addWidget(self.family_approval_input)
        
        # 他番号
        orderer_layout.addWidget(QLabel("他番号"))
        self.other_number_input = QLineEdit()
        self.other_number_input.setText("なし")
        self.other_number_input.textChanged.connect(self.check_input_fields)
        orderer_layout.addWidget(self.other_number_input)
        
        # 電話機
        orderer_layout.addWidget(QLabel("電話機"))
        self.phone_device_input = QLineEdit()
        self.phone_device_input.setText("プッシュホン")
        self.phone_device_input.textChanged.connect(self.check_input_fields)
        orderer_layout.addWidget(self.phone_device_input)
        
        # 禁止回線
        orderer_layout.addWidget(QLabel("禁止回線"))
        self.forbidden_line_input = QLineEdit()
        self.forbidden_line_input.setText("なし")
        self.forbidden_line_input.textChanged.connect(self.check_input_fields)
        orderer_layout.addWidget(self.forbidden_line_input)
        
        # ND
        orderer_layout.addWidget(QLabel("ND"))
        self.nd_input = QLineEdit()
        self.nd_input.textChanged.connect(self.check_input_fields)
        orderer_layout.addWidget(self.nd_input)
        
        # リストとの関係性
        relationship_layout = QHBoxLayout()
        relationship_layout.addWidget(QLabel("備考："))
        self.relationship_input = QLineEdit()
        self.relationship_input.setPlaceholderText("名義人の...")
        self.relationship_input.textChanged.connect(self.check_input_fields)
        relationship_layout.addWidget(self.relationship_input)
        orderer_layout.addLayout(relationship_layout)
        
        # 提供判定エリアを削除
        
        orderer_group.setLayout(orderer_layout)
        scroll_layout.addWidget(orderer_group)
        
        # スクロールエリアにウィジェットを設定
        scroll_area.setWidget(scroll_widget)
        layout.addWidget(scroll_area)
        
        # ボタンエリア
        button_layout = QHBoxLayout()
        
        # 作成ボタン（旧「次へ」ボタンの代わり）
        self.create_btn = QPushButton("作成")
        self.create_btn.setFont(LARGE_FONT)
        self.create_btn.setStyleSheet(BUTTON_STYLE)
        self.create_btn.clicked.connect(self.on_create_clicked)
        button_layout.addWidget(self.create_btn)
        
        # 作成中止ボタン
        self.cancel_btn = QPushButton("作成中止")
        self.cancel_btn.setFont(LARGE_FONT)
        self.cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 15px 30px;
                font-size: 16px;
                border-radius: 8px;
                min-width: 200px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
            QPushButton:pressed {
                background-color: #c62828;
            }
        """)
        self.cancel_btn.clicked.connect(self.on_cancel_clicked)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
        # 初期表示時に一度だけ自動生成を実行
        QTimer.singleShot(100, self.auto_generate_furigana)
        
        # 対応者名にフォーカスを設定
        self.operator_input.setFocus()
        
        # エンターキーで次の項目に移動するためのイベントフィルターを設定
        self.input_fields = [
            self.operator_input,
            self.available_time_input,
            self.contractor_input,
            self.furigana_input,
            self.order_person_input,
            self.fee_input,
            self.net_usage_input,
            self.family_approval_input,
            self.other_number_input,
            self.phone_device_input,
            self.forbidden_line_input,
            self.nd_input,
            self.relationship_input
        ]
        
        for field in self.input_fields:
            field.installEventFilter(self)
        
        # 初期表示時に未入力チェックを実行
        QTimer.singleShot(100, self.check_input_fields)
    
    def check_input_fields(self):
        """
        入力フィールドの未入力状態をチェックし、背景色を変更する
        """
        # 必須入力フィールドのリスト
        required_fields = [
            self.operator_input,
            self.available_time_input,
            self.contractor_input,
            self.furigana_input,
            self.order_person_input,
            self.fee_input,
            self.net_usage_input,
            self.family_approval_input,
            self.other_number_input,
            self.phone_device_input,
            self.forbidden_line_input,
            self.nd_input,
            self.relationship_input
        ]
        
        # 各フィールドの背景色を設定
        for field in required_fields:
            if not field.text().strip():
                field.setStyleSheet("background-color: #FFEBEE;")  # 赤系の背景色
            else:
                field.setStyleSheet("")  # デフォルトの背景色
    
    def on_create_clicked(self):
        """作成ボタンがクリックされた時の処理"""
        try:
            logging.info("作成ボタンがクリックされました")
            
            # 現在の入力内容を保存
            self.saved_data = self.get_orderer_data()
            logging.info(f"受注者データを取得: {self.saved_data}")
            
            # 各フィールドの文末スペースを削除
            for key in self.saved_data:
                if isinstance(self.saved_data[key], str):
                    self.saved_data[key] = self.saved_data[key].rstrip()
            
            # 生年月日を正しいフォーマットに変換（YYYY/M/D形式）
            if 'birth_date' in self.saved_data:
                birth_parts = self.saved_data['birth_date'].split('/')
                if len(birth_parts) == 3:
                    year = int(birth_parts[0])
                    month = int(birth_parts[1])
                    day = int(birth_parts[2])
                    self.saved_data['birth_date'] = f"{year}/{month}/{day}"
            
            # リスト名フリガナをセット
            if self.parent_window and hasattr(self.parent_window, 'list_data') and 'list_furigana' in self.parent_window.list_data:
                list_furigana = self.parent_window.list_data.get('list_furigana')
                if not list_furigana and 'furigana' in self.saved_data:
                    # リスト名フリガナが空で、フリガナが入力されている場合はそれをセット
                    self.parent_window.list_data['list_furigana'] = self.saved_data['furigana']
            
            # 受注情報を作成（現在の日付と初期設定の判定結果を使用）
            now = datetime.datetime.now()
            month = now.month  # 0埋めなしの月
            day = now.day      # 0埋めなしの日
            
            order_data = {
                'current_line': 'アナログ',  # デフォルト値
                'order_date': f"{month}/{day}",
                'judgment': 'OK',  # 初期設定でOKを設定
                'remarks': self.saved_data.get('remarks', '')  # 備考を追加
            }
            
            # 親ウィンドウのプロパティに保存
            if self.parent_window:
                # 重要：最新の受注者データを親ウィンドウに確実に保存
                self.parent_window.orderer_data = self.saved_data
                self.parent_window.order_data = order_data
                
                # 営コメを作成して表示
                if hasattr(self.parent_window, 'generate_preview_text'):
                    logging.info("親ウィンドウのgenerate_preview_textメソッドを呼び出し")
                    
                    # プレビューテキストを生成
                    preview_text = self.parent_window.generate_preview_text()
                    if preview_text:
                        logging.info("プレビューテキストの生成に成功")
                        # プレビューに表示
                        self.parent_window.preview_text.setText(preview_text)
                        logging.info("プレビューにテキストを表示")
                        self.accept()
                    else:
                        logging.error("プレビューテキストの生成に失敗")
                        QMessageBox.warning(self, "警告", "営コメの作成に失敗しました。")
                else:
                    logging.error("親ウィンドウにgenerate_preview_textメソッドが存在しません")
                    QMessageBox.warning(self, "警告", "プレビュー機能が利用できません。")
                    self.accept()
            else:
                logging.error("親ウィンドウへの参照が存在しません")
                QMessageBox.warning(self, "警告", "親ウィンドウへの参照が失われました。")
                self.accept()

        except Exception as e:
            logging.error(f"作成ボタンのクリック処理中にエラー: {e}", exc_info=True)
            QMessageBox.critical(self, "エラー", f"作成処理中にエラーが発生しました: {e}")

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
                'fee': self.fee_input.text(),
                'net_usage': self.net_usage_input.text(),
                'family_approval': self.family_approval_input.text(),
                'other_number': self.other_number_input.text(),
                'phone_device': self.phone_device_input.text(),
                'forbidden_line': self.forbidden_line_input.text(),
                'nd': self.nd_input.text(),
                'relationship': self.relationship_input.text(),
                'remarks': self.relationship_input.text()  # 備考としてリストとの関係性を設定
            }
            
            logging.info(f"受注者データを取得: {data}")
            return data
            
        except Exception as e:
            logging.error(f"受注者データの取得中にエラー: {e}")
            return {}
    
    def get_saved_data(self):
        """保存されたデータを取得"""
        return self.get_orderer_data()

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
            
            # 料金認識
            if data.get('fee'):
                converted_fee = convert_to_half_width(data['fee'])
                self.fee_input.setText(converted_fee)
            
            # ネット利用
            if data.get('net_usage'):
                self.net_usage_input.setText(data['net_usage'])
            
            # 家族了承
            if data.get('family_approval'):
                self.family_approval_input.setText(data['family_approval'])
            
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

    def on_cancel_clicked(self):
        """作成中止ボタンがクリックされた時の処理"""
        # 確認ダイアログを表示
        reply = QMessageBox.question(
            self,
            "確認",
            "作成を中止しますか？\n入力したデータは保存されません。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            # メイン画面に戻る
            self.done(QDialog.DialogCode.Rejected)
        else:
            # キャンセルをキャンセル
            pass

    def schedule_furigana_generation(self):
        """
        フリガナ自動生成をスケジュールする
        """
        # タイマーをリセットして再起動
        self.furigana_timer.stop()
        self.furigana_timer.start(500)  # 500ミリ秒後に実行

    def auto_generate_furigana(self):
        """
        契約者名からフリガナを自動生成する
        """
        try:
            # 自動モードの場合のみ処理
            if self.furigana_mode_combo.currentText() != "自動":
                logging.info("フリガナ自動生成: 手動モードのため生成をスキップします")
                return
                
            # 契約者名が空の場合は何もしない
            name = self.contractor_input.text()
            if not name:
                logging.info("フリガナ自動生成: 契約者名が空のため生成をスキップします")
                return
                
            # 前回と同じ名前の場合はスキップ
            if name == self.last_converted_name:
                logging.info("フリガナ自動生成: 前回と同じ名前のため生成をスキップします")
                return
                
            # フリガナ変換APIを使用
            logging.info(f"フリガナ自動生成: 変換を開始します（契約者名: {name}）")
            from utils.furigana_utils import convert_to_furigana
            furigana = convert_to_furigana(name)
            
            if furigana:
                # blockSignalsでシグナルを一時的にブロックして再帰呼び出しを防止
                self.furigana_input.blockSignals(True)
                self.furigana_input.setText(furigana)
                self.furigana_input.blockSignals(False)
                # 変換した名前を保存
                self.last_converted_name = name
                logging.info(f"フリガナ自動生成: 成功 - {name} → {furigana}")
            else:
                logging.warning(f"フリガナ自動生成: 変換APIから結果が返ってきませんでした（契約者名: {name}）")
        
        except Exception as e:
            logging.error(f"フリガナ自動生成エラー: {str(e)}", exc_info=True)

    def showEvent(self, event):
        """ダイアログが表示される際のイベント"""
        super().showEvent(event)
        
        # 画面のサイズを取得
        screen = QApplication.primaryScreen().geometry()
        screen_width = screen.width()
        screen_height = screen.height()
        
        # ダイアログのサイズを取得
        dialog_width = self.width()
        dialog_height = self.height()
        
        # 提供判定サイトのサイズを想定（一般的なブラウザウィンドウサイズ）
        browser_width = 800
        browser_height = 600
        
        # 提供判定サイトの位置を計算（画面右下）
        browser_x = screen_width - browser_width - 20
        browser_y = screen_height - browser_height - 100
        
        # 受注者入力項目ダイアログを提供判定サイトの左側に配置
        # 提供判定サイトの左端から20px、下端から100pxの位置に配置
        x = browser_x - dialog_width - 20
        y = browser_y + browser_height - dialog_height - 100
        
        # 画面の端に近すぎる場合は調整
        if x < 0:
            x = 20
        if y < 0:
            y = 20
        
        # 位置を設定
        self.move(x, y)
        logging.info(f"受注者入力項目ダイアログを配置: x={x}, y={y}")

    def eventFilter(self, obj, event):
        """
        イベントフィルターを実装して、エンターキーと矢印キーの動作をカスタマイズする
        
        Args:
            obj: イベントの発生元オブジェクト
            event: イベントオブジェクト
            
        Returns:
            bool: イベント処理済みかどうか
        """
        if event.type() == QEvent.Type.KeyPress:
            # 現在の入力フィールドのインデックスを取得
            current_index = -1
            for i, field in enumerate(self.input_fields):
                if obj == field:
                    current_index = i
                    break
            
            if current_index != -1:
                if event.key() == Qt.Key.Key_Return:
                    # Enterキーが押された場合
                    if current_index < len(self.input_fields) - 1:
                        # 次の入力フィールドにフォーカスを移動
                        next_field = self.input_fields[current_index + 1]
                        next_field.setFocus()
                        
                        # スクロールエリアを取得
                        scroll_area = None
                        parent = next_field.parent()
                        while parent:
                            if isinstance(parent, QScrollArea):
                                scroll_area = parent
                                break
                            parent = parent.parent()
                        
                        # スクロールエリアが存在する場合、次のフィールドが見えるようにスクロール
                        if scroll_area:
                            # スクロールエリア内のビューポートの座標系に変換
                            viewport = scroll_area.viewport()
                            widget_pos = next_field.mapTo(viewport, QPoint(0, 0))
                            
                            # 現在のスクロール位置を取得
                            current_scroll_y = scroll_area.verticalScrollBar().value()
                            
                            # ビューポートの高さを取得
                            viewport_height = viewport.height()
                            
                            # ウィジェットの位置と高さ
                            widget_y = widget_pos.y()
                            widget_height = next_field.height()
                            
                            # ウィジェットを中央に配置するためのスクロール位置を計算
                            target_scroll_y = current_scroll_y + widget_y - (viewport_height - widget_height) // 2
                            
                            # スクロール位置を設定
                            scroll_area.verticalScrollBar().setValue(target_scroll_y)
                    elif obj == self.relationship_input:  # 最後の名義人入力エリアの場合
                        # 営コメ作成ボタンをクリック
                        self.on_create_clicked()
                    
                    return True
                
                elif event.key() == Qt.Key.Key_Up:
                    # 上矢印キーが押された場合
                    if current_index > 0:
                        # 前の入力フィールドにフォーカスを移動
                        prev_field = self.input_fields[current_index - 1]
                        prev_field.setFocus()
                        
                        # スクロールエリアを取得
                        scroll_area = None
                        parent = prev_field.parent()
                        while parent:
                            if isinstance(parent, QScrollArea):
                                scroll_area = parent
                                break
                            parent = parent.parent()
                        
                        # スクロールエリアが存在する場合、前のフィールドが見えるようにスクロール
                        if scroll_area:
                            # スクロールエリア内のビューポートの座標系に変換
                            viewport = scroll_area.viewport()
                            widget_pos = prev_field.mapTo(viewport, QPoint(0, 0))
                            
                            # 現在のスクロール位置を取得
                            current_scroll_y = scroll_area.verticalScrollBar().value()
                            
                            # ビューポートの高さを取得
                            viewport_height = viewport.height()
                            
                            # ウィジェットの位置と高さ
                            widget_y = widget_pos.y()
                            widget_height = prev_field.height()
                            
                            # ウィジェットを中央に配置するためのスクロール位置を計算
                            target_scroll_y = current_scroll_y + widget_y - (viewport_height - widget_height) // 2
                            
                            # スクロール位置を設定
                            scroll_area.verticalScrollBar().setValue(target_scroll_y)
                        
                        return True
                
                elif event.key() == Qt.Key.Key_Down:
                    # 下矢印キーが押された場合
                    if current_index < len(self.input_fields) - 1:
                        # 次の入力フィールドにフォーカスを移動
                        next_field = self.input_fields[current_index + 1]
                        next_field.setFocus()
                        
                        # スクロールエリアを取得
                        scroll_area = None
                        parent = next_field.parent()
                        while parent:
                            if isinstance(parent, QScrollArea):
                                scroll_area = parent
                                break
                            parent = parent.parent()
                        
                        # スクロールエリアが存在する場合、次のフィールドが見えるようにスクロール
                        if scroll_area:
                            # スクロールエリア内のビューポートの座標系に変換
                            viewport = scroll_area.viewport()
                            widget_pos = next_field.mapTo(viewport, QPoint(0, 0))
                            
                            # 現在のスクロール位置を取得
                            current_scroll_y = scroll_area.verticalScrollBar().value()
                            
                            # ビューポートの高さを取得
                            viewport_height = viewport.height()
                            
                            # ウィジェットの位置と高さ
                            widget_y = widget_pos.y()
                            widget_height = next_field.height()
                            
                            # ウィジェットを中央に配置するためのスクロール位置を計算
                            target_scroll_y = current_scroll_y + widget_y - (viewport_height - widget_height) // 2
                            
                            # スクロール位置を設定
                            scroll_area.verticalScrollBar().setValue(target_scroll_y)
                        
                        return True
                
        # 標準のイベント処理を継続
        return super().eventFilter(obj, event)

    def on_fee_combo_changed(self, text):
        """
        料金認識のコンボボックスが変更された時の処理
        
        Args:
            text (str): 選択されたテキスト
        """
        self.fee_input.setText(text)
        self.check_input_fields()

    def check_birth_date_age(self):
        """
        生年月日から年齢を計算し、80歳以上の場合に赤く表示する
        """
        try:
            # 現在の日付を取得
            now = datetime.datetime.now()
            current_year = now.year
            current_month = now.month
            current_day = now.day
            
            # 生年月日の情報を取得
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
            
            # 年齢を計算
            age = current_year - year
            
            # 誕生日がまだ来ていない場合は年齢を1つ減らす
            if (month > current_month) or (month == current_month and day > current_day):
                age -= 1
            
            # 80歳以上かどうかをチェック
            is_over_80 = age >= 80
            
            # 背景色を設定
            if is_over_80:
                style = "background-color: #FFEBEE;"  # 赤系の背景色
            else:
                style = ""  # デフォルトの背景色
            
            # 各コンボボックスにスタイルを適用
            self.era_combo.setStyleSheet(style)
            self.year_combo.setStyleSheet(style)
            self.month_combo.setStyleSheet(style)
            self.day_combo.setStyleSheet(style)
            
            # 80歳以上の場合にログを出力
            if is_over_80:
                logging.info(f"80歳以上の顧客が検出されました: {age}歳")
            
        except Exception as e:
            logging.error(f"年齢チェック中にエラー: {e}")

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
        
        # 保存データの初期化
        self.saved_data = order_data or {}
        
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
        self.current_line_combo = NoWheelComboBox()
        self.current_line_combo.addItems(["アナログ"])
        if self.saved_data and 'current_line' in self.saved_data:
            self.current_line_combo.setCurrentText(self.saved_data['current_line'])
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
        self.judgment_combo = NoWheelComboBox()
        self.judgment_combo.addItems(["OK", "NG"])
        if self.saved_data and 'judgment' in self.saved_data:
            self.judgment_combo.setCurrentText(self.saved_data['judgment'])
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
        self.cancel_btn.clicked.connect(self.on_cancel_clicked)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
        # シグナルの接続
        self.back_btn.clicked.connect(self.on_back_clicked)
        self.create_comment_btn.clicked.connect(self.create_comment)
    
    def on_back_clicked(self):
        """戻るボタンがクリックされた時の処理"""
        # 現在の入力内容を保存
        self.saved_data = self.get_order_data()
        # 一つ前のダイアログに戻る
        self.done(DIALOG_BACK)

    def get_saved_data(self):
        """保存されたデータを取得"""
        return self.saved_data

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
                self.accept()
            else:
                logging.error("親ウィンドウにgenerate_preview_textメソッドが存在しません")
                QMessageBox.warning(self, "警告", "プレビュー機能が利用できません。")
                self.accept()
        except Exception as e:
            logging.error(f"営コメ作成中にエラー: {e}", exc_info=True)
            QMessageBox.critical(self, "エラー", f"営コメの作成中にエラーが発生しました: {e}")

    def on_cancel_clicked(self):
        """作成中止ボタンがクリックされた時の処理"""
        # 確認ダイアログを表示
        reply = QMessageBox.question(
            self,
            "確認",
            "作成を中止しますか？\n入力したデータは保存されません。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            # メイン画面に戻る
            self.done(QDialog.DialogCode.Rejected)
        else:
            # キャンセルをキャンセル
            pass 