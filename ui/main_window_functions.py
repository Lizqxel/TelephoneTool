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
from PySide6.QtWidgets import QMessageBox, QApplication, QWidget
from PySide6.QtCore import QTimer, QThread, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QMessageBox, QApplication

from ui.settings_dialog import SettingsDialog
from services import area_search
from services import oneclick
from utils.format_utils import (format_phone_number, format_phone_number_without_hyphen,
                               format_postal_code, convert_to_half_width)
from utils.furigana_utils import convert_to_furigana


class ServiceAreaSearchWorker(QThread):
    """
    提供エリア検索を非同期で実行するワーカースレッド
    
    Attributes:
        finished (Signal): 検索完了時に発火する信号
        postal_code (str): 検索対象の郵便番号
        address (str): 検索対象の住所
    """
    finished = Signal(dict)
    
    def __init__(self, postal_code, address):
        """
        コンストラクタ
        
        Args:
            postal_code (str): 検索対象の郵便番号
            address (str): 検索対象の住所
        """
        super().__init__()
        self.postal_code = postal_code
        self.address = address
        self.driver = None
    
    def run(self):
        """
        検索を実行
        """
        try:
            # 住所を分割して検索を実行
            result = area_search.search_service_area(self.postal_code, self.address)
            self.finished.emit(result)
        except Exception as e:
            logging.error(f"提供エリア検索中にエラーが発生: {str(e)}")
            self.finished.emit({
                "status": "failure",
                "details": {"error": str(e)},
                "show_popup": True
            })
        finally:
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass


class MainWindowFunctions:
    """メインウィンドウの機能を提供するミックスインクラス"""
    
    def load_settings(self):
        """設定ファイルから設定を読み込む"""
        try:
            # 初期設定を設定
            self.settings = {
                'format_template': "",
                'font_size': 10  # デフォルトのフォントサイズ
            }
            
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    self.format_template = settings.get('format_template', "")
                    # settingsオブジェクトを更新
                    self.settings = settings
            else:
                # デフォルトのフォーマットテンプレート
                self.format_template = """対応者（お客様の名前）：{operator}
工事希望日
★出やすい時間帯：{available_time} 携帯：{mobile}
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
                # デフォルト設定をsettingsに保存
                self.settings['format_template'] = self.format_template
                
                # デフォルト設定をファイルに保存
                with open(self.settings_file, 'w', encoding='utf-8') as f:
                    json.dump(self.settings, f, ensure_ascii=False, indent=2)
                    
            logging.info(f"設定を読み込みました: フォントサイズ={self.settings.get('font_size', 10)}")
                
        except Exception as e:
            logging.error(f"設定の読み込みに失敗しました: {str(e)}")
            QMessageBox.warning(self, "エラー", f"設定の読み込みに失敗しました: {str(e)}")
            
            # エラーが発生した場合でもデフォルト設定を使用
            self.settings = {
                'format_template': "",
                'font_size': 10
            }
            logging.info("エラーが発生したため、デフォルト設定を使用します")
    
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
    
    def copy_cti_to_clipboard(self):
        """
        営業コメントを生成してクリップボードにコピーする
        """
        # 既存のgenerate_cti_formatメソッドを使用して営業コメントを生成
        formatted_text = self.generate_cti_format()
        
        # 生成に成功した場合のみクリップボードにコピーして通知
        if formatted_text:
            # クリップボードにコピー
            clipboard = QApplication.clipboard()
            clipboard.setText(formatted_text)
            
            # 成功メッセージをステータスバーに表示
            self.statusBar().showMessage("営業コメントをクリップボードにコピーしました", 5000)
    
    def generate_cti_format(self):
        """CTIフォーマットを生成するだけで、クリップボードへのコピーは行わない"""
        # 必須項目の検証
        required_fields = {
            'operator': self.operator_input,
            'contractor': self.contractor_input,
            'furigana': self.furigana_input,
            'address': self.address_input,
            'postal_code': self.postal_code_input,  # 郵便番号を追加
            'list_name': self.list_name_input,  # リスト名を追加
            'list_furigana': self.list_furigana_input,  # リストフリガナを追加
            'list_phone': self.list_phone_input,
            'list_postal_code': self.list_postal_code_input,  # リスト郵便番号を追加
            'list_address': self.list_address_input,  # リスト住所を追加
            'order_person': self.order_person_input,
            'available_time': self.available_time_input,  # 出やすい時間帯も必須項目に追加
            'fee': self.fee_input,  # 料金認識を追加
            'relationship': self.relationship_input,  # 名義人との関係性
            'employee_number': self.employee_number_input,  # 社番を追加
            'nd': self.nd_input  # NDを追加
        }
        
        # すべてのフィールドの背景色をリセット
        for field in required_fields.values():
            field.setStyleSheet("")
        
        # 未入力のフィールドをチェック
        missing_fields = []
        for name, field in required_fields.items():
            if not field.text().strip():
                # 背景を赤くする
                field.setStyleSheet("background-color: #FFE4E1;")  # 薄い赤色
                missing_fields.append(name)
        
        # 未入力フィールドがある場合
        has_empty_fields = len(missing_fields) > 0
        
        # 未入力項目があることを警告し、続行するか確認
        if has_empty_fields:
            # 日本語のフィールド名に変換
            field_names_ja = {
                'operator': '対応者名',
                'contractor': '契約者名',
                'furigana': 'フリガナ',
                'address': '住所',
                'postal_code': '郵便番号',  # 追加
                'list_name': 'リスト名',  # 追加
                'list_furigana': 'リストフリガナ',  # 追加
                'list_phone': '電話番号',
                'list_postal_code': 'リスト郵便番号',  # 追加
                'list_address': 'リスト住所',  # 追加
                'order_person': '受注者名',
                'available_time': '出やすい時間帯',
                'fee': '料金認識',  # 追加
                'relationship': '備考：名義人の...',  # 名義人との関係性に変更
                'employee_number': '社番',  # 社番を追加
                'nd': 'ND'  # NDを追加
            }
            
            # 未入力項目の日本語名のリスト
            missing_fields_ja = [field_names_ja[field] for field in missing_fields]
            
            # 確認ダイアログを表示（日本語ボタン）
            message_box = QMessageBox()
            message_box.setWindowTitle("未入力項目があります")
            message_box.setText(f"以下の項目が未入力です:\n\n{', '.join(missing_fields_ja)}\n\n営コメを作成しますか？")
            message_box.setIcon(QMessageBox.Question)
            yes_button = message_box.addButton("はい", QMessageBox.YesRole)
            no_button = message_box.addButton("いいえ", QMessageBox.NoRole)
            message_box.setDefaultButton(no_button)
            
            message_box.exec()
            
            # いいえが選択された場合は処理を中止
            if message_box.clickedButton() == no_button:
                return None
        
        # 日付の書式設定
        order_date = self.order_date_input.text()
        
        # 生年月日の取得
        birth_date = ""
        era = self.era_combo.currentText()
        year = self.year_combo.currentText()
        month = self.month_combo.currentText()
        day = self.day_combo.currentText()
        
        if era and year and month and day:
            # 和暦から西暦への変換
            era_year_map = {"令和": 2018, "平成": 1988, "昭和": 1925, "大正": 1911, "明治": 1867, "西暦": 0}
            if era in era_year_map and year.isdigit() and month.isdigit() and day.isdigit():
                try:
                    jp_year = int(year)
                    if era == "西暦":
                        western_year = jp_year
                    else:
                        western_year = era_year_map[era] + jp_year
                    birth_date = f"{western_year}/{month}/{day}"
                except (ValueError, TypeError):
                    logging.warning("生年月日の変換に失敗しました")
                    
        # フォーマットデータの準備
        format_data = {
            'operator': self.operator_input.text(),
            'mobile': "",  # 携帯電話番号を空に
            'available_time': self.available_time_input.text(),  # 出やすい時間帯を追加
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
            'employee_number': self.employee_number_input.text(),  # 社番を追加
            'other_number': self.other_number_input.text(),  # 他番号を追加
            'phone_device': self.phone_device_input.text(),  # 電話機を追加
            'forbidden_line': self.forbidden_line_input.text(),  # 禁止回線を追加
            'nd': self.nd_input.text(),  # NDを追加
            'relationship': self.relationship_input.text()  # 名義人との関係性
        }
        
        # フォーマットテンプレートに値を埋め込む
        try:
            formatted_text = self.format_template.format(**format_data)

            # GoogleマップのURLを追加
            maps_url = self.get_google_maps_url()
            if maps_url:
                formatted_text += f"\n\nGoogleマップ URL: {maps_url}"
            
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
            
            # フォーマットしたテキストを返す
            return formatted_text
        except KeyError as e:
            logging.error(f"フォーマットテンプレートに不明なプレースホルダーがあります: {e}")
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"フォーマットの生成に失敗しました: {str(e)}")
        
        return None
    
    def clear_all_inputs(self):
        """全ての入力フィールドをクリア"""
        # テキスト入力フィールドのクリア
        self.operator_input.clear()
        # 携帯電話番号入力エリアの参照を削除
        self.available_time_input.clear()  # 出やすい時間帯をクリア
        self.contractor_input.clear()
        self.furigana_input.clear()
        self.postal_code_input.clear()
        self.address_input.clear()
        self.list_name_input.clear()
        self.list_furigana_input.clear()
        self.list_phone_input.clear()
        self.list_postal_code_input.clear()
        self.list_address_input.clear()
        # 受注者名はクリアしない（保持する）
        # self.order_person_input.clear()
        
        # 社番はクリアしない（保持する）
        # self.employee_number_input.clear()
        
        # 他番号、電話機、禁止回線には初期値を設定
        self.other_number_input.setText("なし")
        self.phone_device_input.setText("プッシュホン")
        self.forbidden_line_input.setText("なし")
        
        self.nd_input.clear()               # ND
        self.relationship_input.clear()     # 名義人との関係性
        
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
        self.net_usage_combo.setCurrentIndex(0)  # "なし"が選択される
        self.family_approval_combo.setCurrentIndex(0)  # okがインデックス0になる
        
        # 受注日を今日の日付に更新（0埋めなし）
        now = datetime.datetime.now()
        month = str(now.month)  # 0埋めなしの月
        day = str(now.day)      # 0埋めなしの日
        self.order_date_input.setText(f"{month}/{day}")
        
        # 料金認識は初期値に戻さない（保持する）
        # self.fee_input.setText("2500円～3000円")
        
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
            # フォントサイズを適用
            self.apply_font_size()
            # ウィジェットを更新
            self.update()
            # 全てのウィジェットを再描画
            for widget in self.findChildren(QWidget):
                widget.update()
            logging.info("設定を更新しました")
    
    def toggle_mobile_input(self, text):
        """携帯電話番号入力フィールドの有効/無効を切り替え - 現在は使用しない"""
        # 携帯電話番号入力エリアの削除に伴い、このメソッドは使用しません
        pass
    
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
            # 一般電話番号として扱う
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
        # 郵便番号と住所を取得
        postal_code = self.postal_code_input.text().strip()
        address = self.address_input.text().strip()
        
        # 入力チェック
        if not postal_code or not address:
            QMessageBox.warning(self, "入力エラー", "郵便番号と住所を入力してください。")
            return
        
        # 検索中の表示
        self.area_result_label.setText("提供エリア: 検索中...")
        self.area_result_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                padding: 5px;
                border: 1px solid #3498DB;
                border-radius: 4px;
                background-color: #E3F2FD;
                color: #3498DB;
            }
        """)
        QApplication.processEvents()
        
        # 検索ボタンを無効化
        if hasattr(self, 'area_search_btn'):
            self.area_search_btn.setEnabled(False)
            self.area_search_btn.setText("検索中...")
        
        # ワーカースレッドを作成して開始
        self.search_worker = ServiceAreaSearchWorker(postal_code, address)
        self.search_worker.finished.connect(self.on_search_completed)
        self.search_worker.start()
    
    def on_search_completed(self, result):
        """検索完了時の処理"""
        # 検索ボタンを有効化
        if hasattr(self, 'area_search_btn'):
            self.area_search_btn.setEnabled(True)
            self.area_search_btn.setText("検索")
        
        # 結果表示を更新
        status = result.get("status", "failure")
        
        if status == "available":
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
        elif status == "unavailable":
            # 未提供の場合
            self.area_result_label.setText("提供エリア: 未提供")
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
        else:
            # 判定失敗の場合
            self.area_result_label.setText("提供エリア: 判定失敗")
            self.area_result_label.setStyleSheet("""
                QLabel {
                    font-size: 14px;
                    padding: 5px;
                    border: 1px solid #F39C12;
                    border-radius: 4px;
                    background-color: #FFF3E0;
                    color: #F39C12;
                }
            """)
            self.judgment_combo.setCurrentText("")
        
        # 詳細情報がある場合は表示
        if "details" in result and result.get("show_popup", True):
            details = result["details"]
            details_text = "\n".join([f"{k}: {v}" for k, v in details.items()])
            QMessageBox.information(self, "検索結果", details_text)
        
        # スクリーンショットパスを更新
        if "screenshot" in result:
            self.update_screenshot_button(result["screenshot"])
    
    def update_screenshot_button(self, screenshot_path=None):
        """スクリーンショットボタンを更新"""
        if screenshot_path:
            self.screenshot_path = screenshot_path
            self.screenshot_btn.setEnabled(True)
        else:
            self.screenshot_btn.setEnabled(False)
    
    def show_screenshot(self):
        """スクリーンショットを表示する"""
        try:
            # スクリーンショットファイルのパスを取得
            if hasattr(self, 'screenshot_path') and self.screenshot_path:
                screenshot_path = self.screenshot_path
            else:
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
            dialog.setWindowTitle("スクリーンショット - 提供判定結果")
            dialog.setMinimumSize(800, 600)
            layout = QVBoxLayout(dialog)
            
            label = QLabel()
            pixmap = QPixmap(screenshot_path)
            label.setPixmap(pixmap)
            label.setScaledContents(True)
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
            
    def get_google_maps_url(self):
        """住所からGoogleマップのURLを取得する"""
        try:
            # 住所を取得
            address = self.address_input.text()
            if not address:
                return ""
                
            # URLエンコード
            from urllib.parse import quote
            encoded_address = quote(address)
            
            # Google Mapの検索URL
            url = f"https://www.google.com/maps/search/{encoded_address}"
            return url
            
        except Exception as e:
            logging.error(f"Googleマップ URL取得エラー: {str(e)}")
            return ""
            
    def apply_font_size(self):
        """フォントサイズを適用する"""
        try:
            # 設定ファイルからフォントサイズを取得
            font_size = 10  # デフォルト値
            
            if hasattr(self, 'settings') and 'font_size' in self.settings:
                font_size = self.settings['font_size']
            else:
                logging.warning("設定が見つからないため、デフォルトのフォントサイズ(10)を使用します")
                # 設定が存在しない場合は初期化
                if not hasattr(self, 'settings'):
                    self.settings = {'font_size': font_size}
                else:
                    self.settings['font_size'] = font_size
            
            # アプリケーション全体のフォントを設定
            app = QApplication.instance()
            font = QFont()
            font.setPointSize(font_size)
            app.setFont(font)
            
            # スタイルシートを使用してフォントサイズを設定
            self.setStyleSheet(f"* {{ font-size: {font_size}pt; }}")
            
            # メインウィンドウの全てのウィジェットに対してフォントを再設定
            for widget in self.findChildren(QWidget):
                widget_font = widget.font()
                widget_font.setPointSize(font_size)
                widget.setFont(widget_font)
            
            logging.info(f"フォントサイズを {font_size} に設定しました")
            
        except Exception as e:
            logging.error(f"フォントサイズ適用エラー: {str(e)}")
            QMessageBox.warning(self, "エラー", f"フォントサイズの適用に失敗しました: {str(e)}")
            
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
    
    def generate_preview_text(self):
        """
        プレビュー表示用のテキストを生成します。
        入力値を使用して書式設定された営業コメントのプレビューを生成します。
        
        Returns:
            str: 書式設定されたテキスト、または失敗した場合はNone
        """
        try:
            # 入力値の取得（空でも許可）
            operator = self.operator_input.text()
            contractor = self.contractor_input.text() 
            address = self.address_input.text()
            postal_code = self.postal_code_input.text()
            
            # 日付の計算
            birth_date = ""
            era = self.era_combo.currentText()
            year = self.year_combo.currentText()
            month = self.month_combo.currentText()
            day = self.day_combo.currentText()
            
            if era and year and year != "年" and month and month != "月" and day and day != "日":
                # 和暦から西暦への変換
                era_year_map = {"令和": 2018, "平成": 1988, "昭和": 1925, "大正": 1911, "明治": 1867}
                if era in era_year_map:
                    try:
                        jp_year = int(year)
                        western_year = era_year_map[era] + jp_year
                        birth_date = f"{western_year}/{month}/{day}"
                    except ValueError:
                        pass
            
            # フォーマットデータの準備（generate_cti_formatと同じ処理）
            format_data = {
                'operator': self.operator_input.text(),
                'mobile': "",  # 携帯電話番号を空に
                'available_time': self.available_time_input.text(),  # 出やすい時間帯を追加
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
                'order_date': self.order_date_input.text(),
                'order_person': self.order_person_input.text(),
                'judgment': self.judgment_combo.currentText(),
                'fee': self.fee_input.text(),
                'net_usage': self.net_usage_combo.currentText(),
                'family_approval': self.family_approval_combo.currentText(),
                'employee_number': self.employee_number_input.text(),  # 社番を追加
                'other_number': self.other_number_input.text(),  # 他番号を追加
                'phone_device': self.phone_device_input.text(),  # 電話機を追加
                'forbidden_line': self.forbidden_line_input.text(),  # 禁止回線を追加
                'nd': self.nd_input.text(),  # NDを追加
                'relationship': self.relationship_input.text()  # 名義人との関係性
            }
            
            # フォーマットテンプレートに値を埋め込む
            try:
                formatted_text = self.format_template.format(**format_data)
                
                # GoogleマップのURLを追加
                maps_url = self.get_google_maps_url()
                if maps_url:
                    formatted_text += f"\n\nGoogleマップ URL: {maps_url}"
                
                return formatted_text
            except KeyError as e:
                logging.error(f"テンプレート書式エラー: {str(e)}")
                return None
            
        except Exception as e:
            logging.error(f"プレビュー生成中にエラー: {e}")
            return None 