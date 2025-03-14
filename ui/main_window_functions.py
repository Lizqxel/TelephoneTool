"""
メインウィンドウの機能

このモジュールは、メインウィンドウの各種機能メソッドを
提供します。
"""

import datetime
import logging
import json
import os
import re
from PySide6.QtWidgets import QMessageBox, QApplication
from PySide6.QtCore import QTimer, QThread, Signal

from ui.settings_dialog import SettingsDialog
from services.area_search import search_service_area
from utils.format_utils import (format_phone_number, format_phone_number_without_hyphen,
                               format_postal_code, convert_to_half_width)
from utils.furigana_utils import convert_to_furigana


class ServiceAreaSearchWorker(QThread):
    """提供エリア検索を実行するワーカースレッド"""
    finished = Signal(dict)
    
    def __init__(self, postal_code, address):
        super().__init__()
        self.postal_code = postal_code
        self.address = address
    
    def run(self):
        """検索を実行して結果を発行"""
        try:
            result = search_service_area(self.postal_code, self.address)
            self.finished.emit(result)
        except Exception as e:
            self.finished.emit({
                "status": "error",
                "message": f"検索中にエラーが発生しました: {str(e)}"
            })

class MainWindowFunctions:
    """メインウィンドウの機能を提供するミックスインクラス"""
    
    def load_settings(self):
        """設定ファイルから設定を読み込む"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    self.format_template = settings.get('format_template', "")
            else:
                # デフォルトのフォーマットテンプレート
                self.format_template = """対応者（お客様の名前）：{operator}
工事希望日
★出やすい時間帯：携帯：{mobile}
★電話取次：アナログ→光電話
★電話OP：
★無線
契約者(書類名義)：{contractor}
フリガナ：{furigana}
生年月日：{birth_date}
郵便番号：{postal_code}
住所：{address}
リスト名：{list_name}
リスト名フリガナ：{list_furigana}
電話番号：{list_phone}
リスト郵便番号：{list_postal_code}
リスト住所：{list_address}
現状回線：{current_line}
受注日：{order_date}
受注者：{order_person}
提供判定：{judgment}

料金認識：{fee}
ネット利用：{net_usage}"""
        except Exception as e:
            logging.error(f"設定の読み込みに失敗しました: {str(e)}")
            QMessageBox.warning(self, "エラー", f"設定の読み込みに失敗しました: {str(e)}")
    
    def format_phone_number(self):
        """電話番号の自動フォーマット処理"""
        sender = self.sender()
        if sender:
            current_text = sender.text()
            formatted_text = format_phone_number(current_text)
            if formatted_text != current_text:
                sender.setText(formatted_text)
    
    def format_phone_number_without_hyphen(self):
        """電話番号の自動フォーマット処理（ハイフンなし）"""
        sender = self.sender()
        if sender:
            current_text = sender.text()
            formatted_text = format_phone_number_without_hyphen(current_text)
            if formatted_text != current_text:
                sender.setText(formatted_text)
    
    def format_postal_code(self):
        """郵便番号の自動フォーマット処理"""
        sender = self.sender()
        if sender:
            current_text = sender.text()
            formatted_text = format_postal_code(current_text)
            if formatted_text != current_text:
                sender.setText(formatted_text)
    
    def convert_to_half_width(self):
        """全角文字を半角に変換する処理"""
        sender = self.sender()
        if sender:
            current_text = sender.text()
            converted_text = convert_to_half_width(current_text)
            if converted_text != current_text:
                sender.setText(converted_text)
    
    def update_year_combo(self, text):
        """元号に応じて年の選択肢を更新"""
        self.year_combo.clear()
        if text == "昭和":
            self.year_combo.addItems([str(i) for i in range(1, 65)])  # 昭和1年～64年
        elif text == "平成":
            self.year_combo.addItems([str(i) for i in range(1, 32)])  # 平成1年～31年
        else:  # 西暦
            self.year_combo.addItems([str(i) for i in range(1926, datetime.datetime.now().year + 1)])
    
    def generate_cti_format(self):
        """CTIフォーマットを生成してクリップボードにコピー"""
        # 生年月日の計算
        era = self.era_combo.currentText()
        year = int(self.year_combo.currentText())
        month = int(self.month_combo.currentText())
        day = int(self.day_combo.currentText())
        
        # 元号から西暦に変換
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
        order_date = f"{month_str}/{day_str}"
        
        # フォーマットデータの準備
        format_data = {
            'operator': self.operator_input.text(),
            'mobile': "なし" if self.mobile_type_combo.currentText() == "なし" else self.mobile_input.text(),
            'contractor': self.contractor_input.text(),
            'furigana': self.furigana_input.text(),
            'birth_date': birth_date,
            'postal_code': self.postal_code_input.text(),
            'address': self.address_input.text(),
            'list_name': self.list_name_input.text(),
            'list_furigana': self.list_furigana_input.text(),
            'list_phone': self.list_phone_input.text(),
            'list_postal_code': self.list_postal_code_input.text(),
            'list_address': self.list_address_input.text(),
            'current_line': self.current_line_combo.currentText(),
            'order_date': order_date,
            'order_person': self.order_person_input.text(),
            'judgment': self.judgment_combo.currentText(),
            'fee': self.fee_input.text(),
            'net_usage': self.net_usage_combo.currentText(),
            'family_approval': self.family_approval_combo.currentText(),
            'remarks': self.remarks_input.toPlainText()
        }
        
        # フォーマットテンプレートに値を埋め込む
        try:
            formatted_text = self.format_template.format(**format_data)
            
            # プレビューに表示
            self.preview_text.setText(formatted_text)
            
            # プレビューテキストの色を確保
            self.preview_text.setStyleSheet("""
                QTextEdit {
                    background-color: #f8f8f8;
                    color: #333333;
                    border: 1px solid #ddd;
                    border-radius: 4px;
                    padding: 8px;
                    font-family: 'MS Gothic', monospace;
                }
            """)
            
            # クリップボードにコピー
            clipboard = QApplication.clipboard()
            clipboard.setText(formatted_text)
            
            QMessageBox.information(self, "成功", "CTIフォーマットをクリップボードにコピーしました。")
        except KeyError as e:
            QMessageBox.warning(self, "エラー", f"フォーマットテンプレートに不明なプレースホルダーがあります: {e}")
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"フォーマットの生成に失敗しました: {str(e)}")
    
    def clear_all_inputs(self):
        """全ての入力フィールドをクリア"""
        # テキスト入力フィールドのクリア
        self.operator_input.clear()
        self.mobile_type_combo.setCurrentIndex(0)  # 携帯電話番号の選択を「入力」に戻す
        self.mobile_input.setEnabled(True)  # 入力フィールドを有効化
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
        
        # 提供エリア検索結果をリセット
        self.area_result_label.setText("提供エリア: 未検索")
        self.area_result_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                padding: 5px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: #f8f8f8;
            }
        """)
        
        # コンボボックスを初期値に戻す
        self.era_combo.setCurrentIndex(0)
        self.year_combo.setCurrentIndex(0)
        self.month_combo.setCurrentIndex(0)
        self.day_combo.setCurrentIndex(0)
        self.current_line_combo.setCurrentIndex(0)
        self.judgment_combo.setCurrentIndex(0)
        self.net_usage_combo.setCurrentIndex(0)
        self.family_approval_combo.setCurrentIndex(0)
        
        # 受注日を今日の日付に更新
        self.order_date_input.setText(datetime.datetime.now().strftime("%Y/%m/%d"))
        
        # 料金認識を初期値に戻す
        self.fee_input.setText("3000円～3500円")
        
        # 備考欄をクリア
        self.remarks_input.clear()
        
        # プレビューエリアをクリア
        self.preview_text.clear()
    
    def setup_google_sheets(self):
        """Google Sheetsの設定"""
        # この部分は実装しない（必要に応じて実装）
        pass
    
    def write_to_spreadsheet(self):
        """スプレッドシートにデータを書き込む"""
        # この部分は実装しない（必要に応じて実装）
        QMessageBox.information(self, "情報", "スプレッドシート連携機能は実装されていません。")
    
    def show_settings(self):
        """設定ダイアログを表示"""
        dialog = SettingsDialog(self)
        if dialog.exec():
            # ダイアログがOKで閉じられた場合、設定を再読み込み
            self.load_settings()
    
    def toggle_mobile_input(self, text):
        """携帯電話番号入力フィールドの有効/無効を切り替え"""
        self.mobile_input.setEnabled(text == "入力")
        if text == "なし":
            self.mobile_input.clear()
    
    def toggle_clipboard_monitor(self):
        """クリップボード監視の開始/停止を切り替え"""
        if self.clipboard_toggle_btn.isChecked():
            self.clipboard_timer.start(1000)  # 1秒ごとにチェック
            QMessageBox.information(self, "クリップボード監視", "クリップボード監視を開始しました。\n他のアプリからコピーした情報を自動で取得します。")
        else:
            self.clipboard_timer.stop()
            QMessageBox.information(self, "クリップボード監視", "クリップボード監視を停止しました。")
    
    def check_clipboard(self):
        """クリップボードの内容をチェック"""
        text = self.clipboard.text()
        if text and text != self.last_clipboard_text:
            self.last_clipboard_text = text
            self.analyze_clipboard_content(text)
    
    def analyze_clipboard_content(self, text):
        """クリップボードの内容を解析して適切なフィールドに入力"""
        # 電話番号（ハイフンあり/なし）のパターン
        phone_pattern = re.compile(r'(\d{2,4}[-\s]?\d{2,4}[-\s]?\d{4})')
        phone_matches = phone_pattern.finditer(text)
        
        # 郵便番号（ハイフンあり/なし）のパターン
        postal_pattern = re.compile(r'(\d{3}[-\s]?\d{4})')
        postal_match = postal_pattern.search(text)
        
        # 電話番号の処理
        for match in phone_matches:
            phone_number = match.group(1)
            # 携帯電話番号の判定（070, 080, 090で始まる番号）
            if phone_number.replace('-', '').replace(' ', '').startswith(('070', '080', '090')):
                self.mobile_input.setText(phone_number)
                self.mobile_type_combo.setCurrentText("入力")
            else:
                self.list_phone_input.setText(phone_number)
        
        # 郵便番号の処理
        if postal_match:
            postal_code = postal_match.group(1)
            self.postal_code_input.setText(postal_code)
            self.list_postal_code_input.setText(postal_code)
        
        # 住所らしき文字列（漢字とカタカナが含まれる長い文字列）
        if len(text) > 10 and any(ord(c) >= 0x4E00 and ord(c) <= 0x9FFF for c in text):
            self.address_input.setText(text)
            self.list_address_input.setText(text)
        
        # カタカナのみの文字列（フリガナとして扱う）
        if all(ord(c) >= 0x30A0 and ord(c) <= 0x30FF or c.isspace() for c in text):
            self.list_furigana_input.setText(text)
        
        # その他の文字列（名前として扱う）
        if len(text) <= 20 and any(ord(c) >= 0x4E00 and ord(c) <= 0x9FFF for c in text):
            self.list_name_input.setText(text)
    
    def search_service_area(self):
        """提供エリア検索を実行"""
        postal_code = self.postal_code_input.text()
        address = self.address_input.text()
        
        if not postal_code or not address:
            QMessageBox.warning(self, "エラー", "郵便番号と住所を入力してください。")
            return
        
        # 検索ボタンを無効化し、進捗状態を表示
        self.area_search_btn.setEnabled(False)
        self.area_search_btn.setText("検索中...")
        
        # ワーカースレッドを作成して開始
        self.search_worker = ServiceAreaSearchWorker(postal_code, address)
        self.search_worker.finished.connect(self.on_search_completed)
        self.search_worker.start()
    
    def on_search_completed(self, result):
        """検索完了時の処理"""
        # UIを元の状態に戻す
        self.area_search_btn.setEnabled(True)
        self.area_search_btn.setText("検索")
        
        # スクリーンショットボタンを更新
        self.update_screenshot_button(result.get("screenshot"))
        
        # 結果表示を更新
        if result["status"] == "success":
            # 提供可能の場合
            self.area_result_label.setText("提供エリア: 提供可能")
            self.area_result_label.setStyleSheet("""
                QLabel {
                    font-size: 14px;
                    padding: 5px;
                    border: 1px solid #27AE60;
                    border-radius: 4px;
                    background-color: #E8F5E9;
                    color: #27AE60;
                }
            """)
            self.judgment_combo.setCurrentText("○")
        else:
            # 提供不可の場合
            self.area_result_label.setText("提供エリア: 提供不可")
            self.area_result_label.setStyleSheet("""
                QLabel {
                    font-size: 14px;
                    padding: 5px;
                    border: 1px solid #E74C3C;
                    border-radius: 4px;
                    background-color: #FFEBEE;
                    color: #E74C3C;
                }
            """)
            self.judgment_combo.setCurrentText("×")
        
        # 詳細情報がある場合は表示
        if "details" in result:
            details = result["details"]
            details_text = "\n".join([f"{k}: {v}" for k, v in details.items()])
            QMessageBox.information(self, "検索結果", details_text)
    
    def update_screenshot_button(self, screenshot_path):
        """スクリーンショットボタンを更新"""
        # このメソッドは実装しない（必要に応じて実装）
        pass
    
    def update_screenshot_button(self):
        """スクリーンショットボタンの状態を更新"""
        # 実装は必要に応じて追加
        pass
        
    def show_screenshot(self):
        """スクリーンショットを表示する"""
        try:
            # スクリーンショットファイルのパスを取得
            screenshot_path = "debug_screenshot.png"
            
            # ファイルが存在するか確認
            if not os.path.exists(screenshot_path):
                QMessageBox.warning(
                    self,
                    "エラー",
                    "スクリーンショットファイルが見つかりません。"
                )
                return
                
            # QPixmapを使用して画像を表示
            from PySide6.QtGui import QPixmap
            from PySide6.QtWidgets import QLabel, QDialog, QVBoxLayout
            
            dialog = QDialog(self)
            dialog.setWindowTitle("スクリーンショット")
            layout = QVBoxLayout(dialog)
            
            label = QLabel()
            pixmap = QPixmap(screenshot_path)
            label.setPixmap(pixmap)
            layout.addWidget(label)
            
            dialog.setLayout(layout)
            dialog.exec()
            
        except Exception as e:
            logging.error(f"スクリーンショット表示エラー: {str(e)}")
            QMessageBox.critical(
                self,
                "エラー",
                f"スクリーンショットの表示中にエラーが発生しました: {str(e)}"
            )
            
    def open_street_view(self):
        """住所からGoogleマップを開く"""
        try:
            # 住所を取得
            address = self.address_input.text()
            if not address:
                QMessageBox.warning(self, "エラー", "住所が入力されていません。")
                return
                
            # URLエンコード
            from urllib.parse import quote
            encoded_address = quote(address)
            
            # Google Mapの検索URL（ストリートビューではなく通常の検索結果）
            url = f"https://www.google.com/maps/search/{encoded_address}"
            
            # ブラウザで開く
            from PySide6.QtCore import QUrl
            from PySide6.QtGui import QDesktopServices
            QDesktopServices.openUrl(QUrl(url))
            
            logging.info(f"Googleマップを開きました: {address}")
            
        except Exception as e:
            logging.error(f"Googleマップ表示エラー: {str(e)}")
            QMessageBox.critical(
                self,
                "エラー",
                f"Googleマップの表示中にエラーが発生しました: {str(e)}"
            )
            
    def apply_font_size(self):
        """フォントサイズを適用する"""
        try:
            from PySide6.QtGui import QFont
            
            # 設定ファイルからフォントサイズを取得
            font_size = 10  # デフォルト値
            
            if hasattr(self, 'settings') and 'font_size' in self.settings:
                font_size = self.settings['font_size']
            
            # アプリケーション全体のフォントを設定
            app = QApplication.instance()
            font = app.font()
            font.setPointSize(font_size)
            app.setFont(font)
            
            logging.info(f"フォントサイズを {font_size} に設定しました")
            
        except Exception as e:
            logging.error(f"フォントサイズ適用エラー: {str(e)}")
            
    def auto_generate_furigana(self):
        """契約者名からフリガナを自動生成する"""
        # 自動モードの場合のみ処理
        if self.furigana_mode_combo.currentText() != "自動":
            return
            
        # 契約者名が空の場合は何もしない
        name = self.contractor_input.text()
        if not name:
            return
            
        try:
            # フリガナ変換APIを使用
            furigana = convert_to_furigana(name)
            if furigana:
                self.furigana_input.setText(furigana)
                logging.info(f"フリガナを自動生成しました: {name} → {furigana}")
        except Exception as e:
            logging.error(f"フリガナ自動生成エラー: {str(e)}")
            
    def auto_generate_list_furigana(self):
        """リスト名からフリガナを自動生成する"""
        # 自動モードの場合のみ処理
        if self.list_furigana_mode_combo.currentText() != "自動":
            return
            
        # リスト名が空の場合は何もしない
        name = self.list_name_input.text()
        if not name:
            return
            
        try:
            # フリガナ変換APIを使用
            furigana = convert_to_furigana(name)
            if furigana:
                self.list_furigana_input.setText(furigana)
                logging.info(f"リストフリガナを自動生成しました: {name} → {furigana}")
        except Exception as e:
            logging.error(f"リストフリガナ自動生成エラー: {str(e)}") 