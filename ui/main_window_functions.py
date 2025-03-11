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
from PySide6.QtCore import QTimer

from ui.settings_dialog import SettingsDialog
from services.area_search import search_service_area
from utils.format_utils import (format_phone_number, format_phone_number_without_hyphen,
                               format_postal_code, convert_to_half_width)


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
        """提供エリア検索機能"""
        # 郵便番号と住所を取得
        postal_code = self.postal_code_input.text().strip()
        address = self.address_input.text().strip()
        
        # 入力チェック
        if not postal_code or not address:
            QMessageBox.warning(self, "入力エラー", "郵便番号と住所を入力してください。")
            return
        
        # 処理中メッセージを表示
        self.area_result_label.setText("提供エリア: 検索中...")
        self.area_result_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                padding: 5px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: #fffde7;
            }
        """)
        QApplication.processEvents()  # UIを更新
        
        try:
            # 提供エリア検索を実行
            result = search_service_area(postal_code, address)
            
            # 結果に基づいてUIを更新
            if result["status"] == "success":
                self.area_result_label.setText(f"提供エリア: {result['message']}")
                self.area_result_label.setStyleSheet("""
                    QLabel {
                        font-size: 14px;
                        padding: 5px;
                        border: 1px solid #4CAF50;
                        border-radius: 4px;
                        background-color: #E8F5E9;
                        color: #2E7D32;
                    }
                """)
                # 提供判定コンボボックスを更新
                self.judgment_combo.setCurrentText("OK")
            else:
                self.area_result_label.setText(f"提供エリア: {result['message']}")
                self.area_result_label.setStyleSheet("""
                    QLabel {
                        font-size: 14px;
                        padding: 5px;
                        border: 1px solid #F44336;
                        border-radius: 4px;
                        background-color: #FFEBEE;
                        color: #C62828;
                    }
                """)
                # 提供判定コンボボックスを更新
                self.judgment_combo.setCurrentText("NG")
        
        except Exception as e:
            logging.error(f"提供エリア検索中にエラーが発生しました: {str(e)}")
            self.area_result_label.setText(f"提供エリア: エラー ({str(e)})")
            self.area_result_label.setStyleSheet("""
                QLabel {
                    font-size: 14px;
                    padding: 5px;
                    border: 1px solid #FF9800;
                    border-radius: 4px;
                    background-color: #FFF3E0;
                    color: #E65100;
                }
            """)
            QMessageBox.critical(self, "エラー", f"提供エリア検索中にエラーが発生しました:\n{str(e)}") 