"""
住所情報入力ダイアログと関連コンポーネント

このモジュールは、簡易モードの住所情報入力ダイアログを提供します。
また、提供エリア検索用のスレッドも含まれています。
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
                              QLabel, QLineEdit, QPushButton,
                              QGroupBox, QMessageBox, QWidget, QComboBox, QScrollArea)
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtCore import Qt, QThread, Signal, QEvent, QMetaObject, Q_ARG, QTimer, QPoint, QUrl
from PySide6.QtGui import QFont, QIntValidator
import datetime
import logging
from services.area_search import search_service_area as area_search_service
import threading
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Slot

# グローバル変数 - 提供判定の最終結果を保持
# このグローバル変数は関数やメソッド間でエリア判定結果を共有するために使用されます
_last_search_result = "未検索"

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

def search_service_area(postal_code, address):
    """
    提供エリアを検索する
    
    Args:
        postal_code (str): 郵便番号
        address (str): 住所
        
    Returns:
        str: 提供エリア情報
    """
    try:
        # 郵便番号と住所から提供エリアを判定
        # ここでは簡単な例として、郵便番号の最初の3桁で判定
        area_code = postal_code[:3]
        
        # 提供エリアの判定ロジック
        if area_code in ['100', '101', '102', '103', '104', '105']:
            return "東京23区内"
        elif area_code in ['220', '221', '222', '223', '224', '225']:
            return "横浜市内"
        elif area_code in ['330', '331', '332', '333', '334', '335']:
            return "さいたま市内"
        else:
            return "提供エリア外"
            
    except Exception as e:
        logging.error(f"提供エリア検索中にエラー: {e}")
        return None

class ServiceAreaSearchThread(QThread):
    """提供エリア検索用のスレッド"""
    finished = Signal(dict)  # 検索結果を通知するシグナル
    error = Signal(str)     # エラーを通知するシグナル
    
    def __init__(self, postal_code, address, parent_window):
        super().__init__()
        self.postal_code = postal_code
        self.address = address
        self.parent_window = parent_window
    
    def run(self):
        """スレッドの実行"""
        try:
            # 提供エリア検索を実行
            result = area_search_service(self.postal_code, self.address)
            logging.info(f"★★★ area_search_serviceの結果: {result} ★★★")
            
            # area_search_serviceから返された結果をそのまま使用
            if isinstance(result, dict):
                # 既に辞書型の場合はそのまま使用
                self.finished.emit(result)
            else:
                # 文字列の場合は辞書型に変換
                result_dict = {
                    "status": "available" if "提供可能" in str(result) else "unavailable",
                    "message": str(result),
                    "show_popup": True
                }
                logging.info(f"★★★ 変換後の結果: {result_dict} ★★★")
                self.finished.emit(result_dict)
                
        except Exception as e:
            logging.error(f"提供エリア検索中にエラー: {e}", exc_info=True)
            error_dict = {
                "status": "failure",
                "message": str(e),
                "show_popup": True
            }
            self.finished.emit(error_dict)
            
    def stop(self):
        """スレッドを停止する"""
        if self.isRunning():
            logging.info("提供エリア検索スレッドの停止を要求")
            self.terminate()
            self.wait(500)  # 最大0.5秒待機
            logging.info("提供エリア検索スレッドが停止しました")

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
        self.setMinimumWidth(500)
        
        # 検索スレッドの初期化
        self.search_thread = None
        self.parent_window = parent
        
        # 保存データの初期化
        self.saved_data = address_data or {}
        
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
        if self.saved_data and 'postal_code' in self.saved_data:
            self.postal_code_input.setText(self.saved_data['postal_code'])
        address_layout.addWidget(self.postal_code_input)
        
        # 住所
        address_layout.addWidget(QLabel("住所"))
        self.address_input = QLineEdit()
        self.address_input.setPlaceholderText("例：東京都渋谷区...")
        if self.saved_data and 'address' in self.saved_data:
            self.address_input.setText(self.saved_data['address'])
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
        self.judgment_btn.clicked.connect(self.search_service_area)
        layout.addWidget(self.judgment_btn)
        
        # 提供判定結果表示ラベル
        self.judgment_result = QLabel("提供エリア: 未検索")
        self.judgment_result.setStyleSheet("""
            QLabel {
                font-size: 14px;
                padding: 10px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: #f8f9fa;
            }
        """)
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
        self.next_btn.clicked.connect(self.on_next_clicked)
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
    
    def search_service_area(self):
        """提供エリアを検索"""
        try:
            postal_code = self.postal_code_input.text()
            address = self.address_input.text()
            
            if not postal_code or not address:
                QMessageBox.warning(self, "警告", "郵便番号と住所を入力してください。")
                return
            
            # ダイアログの表示を更新
            self.judgment_result.setText("提供エリア: 検索中...")
            self.judgment_result.setStyleSheet("""
                QLabel {
                    font-size: 14px;
                    padding: 5px;
                    border: 1px solid #FFA500;
                    border-radius: 4px;
                    background-color: #FFF3E0;
                    color: #E65100;
                }
            """)
            
            # 検索ボタンを無効化
            self.judgment_btn.setEnabled(False)
            self.judgment_btn.setText("検索中...")
            
            # メインウィンドウへの参照を確保
            main_window = self.parent_window
            
            # グローバル変数をリセット
            global _last_search_result
            _last_search_result = "検索中..."
            logging.info(f"グローバル変数をリセットしました: {_last_search_result}")
            
            # 既存のスレッドを停止
            if hasattr(self, 'search_thread') and self.search_thread is not None:
                try:
                    self.search_thread.finished.disconnect()
                    self.search_thread.error.disconnect()
                    self.search_thread.stop()
                except Exception as e:
                    logging.warning(f"既存のスレッド切断でエラー: {e}")
            
            # 検索スレッドの作成と開始
            self.search_thread = ServiceAreaSearchThread(postal_code, address, main_window)
            
            # シグナルの接続を確認
            if not self.search_thread.finished.receivers(self.on_search_finished):
                self.search_thread.finished.connect(self.on_search_finished)
                logging.info("finishedシグナルを接続しました")
            
            if not self.search_thread.error.receivers(self.on_search_error):
                self.search_thread.error.connect(self.on_search_error)
                logging.info("errorシグナルを接続しました")
            
            # メインウィンドウにも直接「検索中...」を通知
            if main_window and hasattr(main_window, 'update_judgment_result'):
                try:
                    main_window.update_judgment_result("検索中...")
                    logging.info("メインウィンドウに「検索中...」状態を直接通知しました")
                except Exception as e:
                    logging.error(f"メインウィンドウへの検索中状態通知でエラー: {e}")
            
            # スレッドを開始
            self.search_thread.start()
            logging.info(f"提供エリア検索スレッドを開始しました: postal_code={postal_code}, address={address}")
            
        except Exception as e:
            logging.error(f"提供エリア検索の開始でエラー: {e}", exc_info=True)
            self.judgment_result.setText("提供エリア: エラー")
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
            QMessageBox.critical(self, "エラー", f"提供エリア検索の開始に失敗しました: {e}")
    
    def on_search_finished(self, result):
        """検索完了時の処理"""
        try:
            logging.info(f"★★★ 検索完了: {result} ★★★")
            
            # 結果表示を更新
            status = result.get("status", "failure")
            message = result.get("message", "判定失敗")
            
            # グローバル変数に結果を保存
            global _last_search_result
            
            # 判定結果に応じてラベルを更新
            if status == "available":
                self.judgment_result.setText("提供エリア: 提供可能")
                self.judgment_result.setStyleSheet("""
                    QLabel {
                        font-size: 14px;
                        padding: 10px;
                        border: 1px solid #27AE60;
                        border-radius: 4px;
                        background-color: #E8F5E9;
                        color: #27AE60;
                    }
                """)
                _last_search_result = "提供可能"
                logging.info("★★★ 提供可能と判定されました ★★★")
            elif status == "unavailable":
                self.judgment_result.setText("提供エリア: 未提供")
                self.judgment_result.setStyleSheet("""
                    QLabel {
                        font-size: 14px;
                        padding: 10px;
                        border: 1px solid #E74C3C;
                        border-radius: 4px;
                        background-color: #FFEBEE;
                        color: #E74C3C;
                    }
                """)
                _last_search_result = "未提供"
                logging.info("★★★ 未提供と判定されました ★★★")
            else:
                self.judgment_result.setText(f"提供エリア: {message}")
                self.judgment_result.setStyleSheet("""
                    QLabel {
                        font-size: 14px;
                        padding: 10px;
                        border: 1px solid #F39C12;
                        border-radius: 4px;
                        background-color: #FFF3E0;
                        color: #F39C12;
                    }
                """)
                _last_search_result = "検索失敗"
                logging.info("★★★ 検索失敗と判定されました ★★★")
            
            # UIの更新を確実に実行
            QApplication.processEvents()
            
            # 検索ボタンを有効化
            self.judgment_btn.setEnabled(True)
            self.judgment_btn.setText("提供判定実行")
            
            # 親ウィンドウのプレビューを更新
            if hasattr(self.parent_window, 'generate_preview_text'):
                try:
                    logging.info(f"★★★ 親ウィンドウのプレビューを更新します: {_last_search_result} ★★★")
                    self.parent_window.judgment_combo.setCurrentText(_last_search_result)
                    self.parent_window.generate_preview_text()
                    logging.info("★★★ プレビューの更新が完了しました ★★★")
                except Exception as e:
                    logging.error(f"プレビュー更新でエラー: {e}", exc_info=True)
            
            # 詳細情報がある場合は表示
            if "details" in result and result.get("show_popup", True):
                details = result["details"]
                details_text = "\n".join([f"{k}: {v}" for k, v in details.items()])
                QMessageBox.information(self, "検索結果", details_text)
            
            # 最終的なUIの更新を確実に実行
            QApplication.processEvents()
            logging.info("★★★ 検索完了の処理が終了しました ★★★")
            
        except Exception as e:
            logging.error(f"検索完了時の処理でエラー: {e}", exc_info=True)
            self.judgment_result.setText("提供エリア: エラー")
            self.judgment_result.setStyleSheet("""
                QLabel {
                    font-size: 14px;
                    padding: 10px;
                    border: 1px solid #E74C3C;
                    border-radius: 4px;
                    background-color: #FFEBEE;
                    color: #E74C3C;
                }
            """)
            QMessageBox.critical(self, "エラー", f"検索結果の処理中にエラーが発生しました: {e}")
    
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

# ダイアログの結果コード
DIALOG_BACK = 2  # 戻るボタン用
DIALOG_NEXT = QDialog.DialogCode.Accepted  # 次へボタン用
DIALOG_CANCEL = QDialog.DialogCode.Rejected  # 作成中止ボタン用

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
        self.setWindowTitle("受注者入力項目")
        self.setModal(True)
        self.setMinimumWidth(500)
        
        # エンターキーの挙動を変更（ダイアログを閉じないようにする）
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowContextHelpButtonHint)
        
        # 親ウィンドウへの参照を保持
        self.parent_window = parent
        
        # 保存データの初期化
        self.saved_data = orderer_data or {}
        
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
        orderer_layout.addWidget(self.operator_input)
        
        # 出やすい時間帯
        orderer_layout.addWidget(QLabel("出やすい時間帯"))
        self.available_time_input = QLineEdit()
        self.available_time_input.setPlaceholderText("AMPM希望　固定or携帯　000-0000-0000")
        orderer_layout.addWidget(self.available_time_input)
        
        # 契約者名
        orderer_layout.addWidget(QLabel("契約者名"))
        self.contractor_input = QLineEdit()
        orderer_layout.addWidget(self.contractor_input)
        
        # フリガナ
        furigana_layout = QHBoxLayout()
        furigana_layout.addWidget(QLabel("フリガナ"))
        self.furigana_mode_combo = QComboBox()
        self.furigana_mode_combo.addItems(["自動", "手動"])
        furigana_layout.addWidget(self.furigana_mode_combo)
        orderer_layout.addLayout(furigana_layout)
        self.furigana_input = QLineEdit()
        orderer_layout.addWidget(self.furigana_input)
        
        # 生年月日
        birth_layout = QVBoxLayout()
        birth_layout.addWidget(QLabel("生年月日"))
        
        # 生年月日の入力部分を横並びにする
        birth_input_layout = QHBoxLayout()
        
        self.era_combo = NoWheelComboBox()
        self.era_combo.addItems(["昭和", "平成", "西暦"])
        self.era_combo.setFixedWidth(60)
        birth_input_layout.addWidget(self.era_combo)
        
        self.year_combo = NoWheelComboBox()
        self.year_combo.addItems([str(i) for i in range(1, 65)])
        self.year_combo.setEditable(True)
        self.year_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.year_combo.lineEdit().setMaxLength(4)
        self.year_combo.lineEdit().setValidator(QIntValidator(1, 9999))
        self.year_combo.setFixedWidth(60)
        birth_input_layout.addWidget(self.year_combo)
        birth_input_layout.addWidget(QLabel("年"))
        
        self.month_combo = NoWheelComboBox()
        self.month_combo.addItems([str(i) for i in range(1, 13)])
        self.month_combo.setEditable(True)
        self.month_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.month_combo.lineEdit().setMaxLength(2)
        self.month_combo.lineEdit().setValidator(QIntValidator(1, 12))
        self.month_combo.setFixedWidth(40)
        birth_input_layout.addWidget(self.month_combo)
        birth_input_layout.addWidget(QLabel("月"))
        
        self.day_combo = NoWheelComboBox()
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
        orderer_layout.addWidget(self.order_person_input)
        
        # 社番
        orderer_layout.addWidget(QLabel("社番"))
        self.employee_number_input = QLineEdit()
        orderer_layout.addWidget(self.employee_number_input)
        
        # 料金認識
        orderer_layout.addWidget(QLabel("料金認識"))
        self.fee_input = QLineEdit()
        self.fee_input.setText("2500円～3000円")
        orderer_layout.addWidget(self.fee_input)
        
        # ネット利用
        orderer_layout.addWidget(QLabel("ネット利用"))
        self.net_usage_combo = NoWheelComboBox()
        self.net_usage_combo.addItems(["なし", "あり"])
        orderer_layout.addWidget(self.net_usage_combo)
        
        # 家族了承
        orderer_layout.addWidget(QLabel("家族了承"))
        self.family_approval_combo = NoWheelComboBox()
        self.family_approval_combo.addItems(["ok", "なし"])
        orderer_layout.addWidget(self.family_approval_combo)
        
        # 他番号
        orderer_layout.addWidget(QLabel("他番号"))
        self.other_number_input = QLineEdit()
        self.other_number_input.setText("なし")
        orderer_layout.addWidget(self.other_number_input)
        
        # 電話機
        orderer_layout.addWidget(QLabel("電話機"))
        self.phone_device_input = QLineEdit()
        self.phone_device_input.setText("プッシュホン")
        orderer_layout.addWidget(self.phone_device_input)
        
        # 禁止回線
        orderer_layout.addWidget(QLabel("禁止回線"))
        self.forbidden_line_input = QLineEdit()
        self.forbidden_line_input.setText("なし")
        orderer_layout.addWidget(self.forbidden_line_input)
        
        # ND
        orderer_layout.addWidget(QLabel("ND"))
        self.nd_input = QLineEdit()
        orderer_layout.addWidget(self.nd_input)
        
        # リストとの関係性
        relationship_layout = QHBoxLayout()
        relationship_layout.addWidget(QLabel("備考："))
        self.relationship_input = QLineEdit()
        self.relationship_input.setPlaceholderText("名義人の...")
        relationship_layout.addWidget(self.relationship_input)
        orderer_layout.addLayout(relationship_layout)
        
        # 提供判定エリア
        judgment_layout = QHBoxLayout()
        judgment_layout.addWidget(QLabel("提供判定："))
        self.judgment_combo = NoWheelComboBox()
        self.judgment_combo.addItems(["OK", "NG"])
        judgment_layout.addWidget(self.judgment_combo)
        orderer_layout.addLayout(judgment_layout)
        
        orderer_group.setLayout(orderer_layout)
        scroll_layout.addWidget(orderer_group)
        
        # スクロールエリアにウィジェットを設定
        scroll_area.setWidget(scroll_widget)
        layout.addWidget(scroll_area)
        
        # ボタンエリア
        button_layout = QHBoxLayout()
        
        # 作成ボタン（旧「次へ」ボタンの代わり）
        self.create_btn = QPushButton("作成")
        self.create_btn.setStyleSheet("""
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
        self.create_btn.clicked.connect(self.on_create_clicked)
        button_layout.addWidget(self.create_btn)
        
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
        # 戻るボタンは削除
        
        # エンターキーが押されたときに次の入力項目に移動するための設定
        self.input_fields = [
            self.operator_input, self.available_time_input, self.contractor_input,
            self.furigana_input, self.order_person_input, self.employee_number_input,
            self.fee_input, self.other_number_input, self.phone_device_input,
            self.forbidden_line_input, self.nd_input, self.relationship_input
        ]
        
        # 各入力フィールドにイベントハンドラーを設定
        for i, field in enumerate(self.input_fields):
            if isinstance(field, QLineEdit):
                field.installEventFilter(self)
        
        # ここでデータをセット（先にデータをセット）
        if self.saved_data and 'operator' in self.saved_data:
            self.operator_input.setText(self.saved_data['operator'])
            
        if self.saved_data and 'available_time' in self.saved_data:
            self.available_time_input.setText(self.saved_data['available_time'])
            
        if self.saved_data and 'contractor' in self.saved_data:
            self.contractor_input.setText(self.saved_data['contractor'])
            
        if self.saved_data and 'furigana' in self.saved_data:
            self.furigana_input.setText(self.saved_data['furigana'])
            
        if self.saved_data and 'order_person' in self.saved_data:
            self.order_person_input.setText(self.saved_data['order_person'])
            
        if self.saved_data and 'employee_number' in self.saved_data:
            self.employee_number_input.setText(self.saved_data['employee_number'])
            
        if self.saved_data and 'fee' in self.saved_data:
            self.fee_input.setText(self.saved_data['fee'])
            
        if self.saved_data and 'net_usage' in self.saved_data:
            self.net_usage_combo.setCurrentText(self.saved_data['net_usage'])
            
        if self.saved_data and 'family_approval' in self.saved_data:
            self.family_approval_combo.setCurrentText(self.saved_data['family_approval'])
            
        if self.saved_data and 'other_number' in self.saved_data:
            self.other_number_input.setText(self.saved_data['other_number'])
            
        if self.saved_data and 'phone_device' in self.saved_data:
            self.phone_device_input.setText(self.saved_data['phone_device'])
            
        if self.saved_data and 'forbidden_line' in self.saved_data:
            self.forbidden_line_input.setText(self.saved_data['forbidden_line'])
            
        if self.saved_data and 'nd' in self.saved_data:
            self.nd_input.setText(self.saved_data['nd'])
            
        if self.saved_data and 'relationship' in self.saved_data:
            self.relationship_input.setText(self.saved_data['relationship'])
            
        # 親ウィンドウの提供判定結果を確認し、判定コンボボックスを更新
        try:
            if hasattr(self.parent_window, 'judgment_result_label'):
                judgment_text = self.parent_window.judgment_result_label.text()
                if "提供可能" in judgment_text:
                    self.judgment_combo.setCurrentText("OK")
                elif "提供エリア外" in judgment_text:
                    self.judgment_combo.setCurrentText("NG")
        except Exception as e:
            logging.error(f"提供判定結果の取得でエラー: {e}")
            
        # フリガナ自動生成のシグナルを接続
        self.contractor_input.textChanged.connect(self.auto_generate_furigana)
        self.furigana_mode_combo.currentTextChanged.connect(lambda: self.auto_generate_furigana())
        
        # 初期表示時に一度だけ自動生成を実行
        QTimer.singleShot(100, self.auto_generate_furigana)
    
    def eventFilter(self, obj, event):
        """
        イベントフィルターを実装して、エンターキーの動作をカスタマイズする
        
        Args:
            obj: イベントの発生元オブジェクト
            event: イベントオブジェクト
            
        Returns:
            bool: イベント処理済みかどうか
        """
        if event.type() == QEvent.Type.KeyPress and event.key() == Qt.Key.Key_Return:
            # Enterキーが押された場合
            current_index = -1
            for i, field in enumerate(self.input_fields):
                if obj == field:
                    current_index = i
                    break
            
            if current_index != -1 and current_index < len(self.input_fields) - 1:
                # 次の入力フィールドにフォーカスを移動
                next_index = current_index + 1
                next_field = self.input_fields[next_index]
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
                    
                    # ウィジェットが完全に表示されているかチェック
                    if widget_y < 0:
                        # ウィジェットが上に隠れている場合、上にスクロール
                        scroll_area.verticalScrollBar().setValue(current_scroll_y + widget_y)
                    elif widget_y + widget_height > viewport_height:
                        # ウィジェットが下に隠れている場合、下にスクロール
                        scroll_area.verticalScrollBar().setValue(
                            current_scroll_y + (widget_y + widget_height - viewport_height) + 10
                        )
                
                return True
                
        # 標準のイベント処理を継続
        return super().eventFilter(obj, event)
        
    def keyPressEvent(self, event):
        """
        キー押下イベントを処理する
        
        Args:
            event: キーイベント
        """
        # Enterキーでダイアログを閉じないようにする
        if event.key() == Qt.Key.Key_Return or event.key() == Qt.Key.Key_Enter:
            # Enterキーの標準動作を無効化
            event.accept()
        else:
            # それ以外のキーは通常通り処理
            super().keyPressEvent(event)
    
    def on_create_clicked(self):
        """作成ボタンがクリックされた時の処理"""
        try:
            logging.info("作成ボタンがクリックされました")
            
            # 現在の入力内容を保存
            self.saved_data = self.get_orderer_data()
            logging.info(f"受注者データを取得: {self.saved_data}")
            
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
            
            # 受注情報を作成（現在の日付とユーザー選択の判定結果を使用）
            now = datetime.datetime.now()
            month = now.month  # 0埋めなしの月
            day = now.day      # 0埋めなしの日
            
            order_data = {
                'current_line': 'アナログ',  # デフォルト値
                'order_date': f"{month}/{day}",
                'judgment': self.judgment_combo.currentText()  # ユーザーが選択した判定結果
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

    def auto_generate_furigana(self):
        """契約者名からフリガナを自動生成する"""
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
                
            # 既にフリガナが入力されている場合はスキップする特殊なケース
            current_furigana = self.furigana_input.text()
            if current_furigana and len(current_furigana) > 1 and name in self.saved_data.get('contractor', ''):
                logging.info(f"フリガナ自動生成: 既にフリガナ({current_furigana})が設定されているためスキップします")
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
            else:
                logging.error("親ウィンドウにgenerate_preview_textメソッドが存在しません")
                QMessageBox.warning(self, "警告", "プレビュー機能が利用できません。")
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