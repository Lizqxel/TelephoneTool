"""
スタイルシート定義

このモジュールは、アプリケーション全体で使用される
スタイルシートを定義します。
"""

MAIN_STYLE = """
/* グローバルスタイル */
* {
    font-family: 'Yu Gothic UI', 'Meiryo UI', sans-serif;
    font-size: 12px;
}

/* メインウィンドウ */
QMainWindow {
    background-color: #F0F0F0;
}

/* タブウィジェット */
QTabWidget::pane {
    border: 1px solid #C0C0C0;
    background-color: #FFFFFF;
}

QTabWidget::tab-bar {
    alignment: left;
}

QTabBar::tab {
    background-color: #E1E1E1;
    color: #000000;
    padding: 8px 12px;
    border: 1px solid #C0C0C0;
    border-bottom: none;
    min-width: 80px;
    margin-right: 2px;
}

QTabBar::tab:selected {
    background-color: #0078D7;
    color: #FFFFFF;
}

QTabBar::tab:hover:!selected {
    background-color: #E5F3FF;
}

/* グループボックス */
QGroupBox {
    border: 1px solid #C0C0C0;
    border-radius: 4px;
    margin-top: 12px;
    padding: 12px;
    background-color: #FFFFFF;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 8px;
    padding: 0 3px;
    background-color: #FFFFFF;
    font-weight: bold;
}

/* 入力フィールド */
QLineEdit, QTextEdit, QComboBox {
    border: 1px solid #C0C0C0;
    border-radius: 2px;
    padding: 4px;
    background-color: #FFFFFF;
    height: 23px;
}

QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
    border: 1px solid #0078D7;
}

QLineEdit:hover, QTextEdit:hover, QComboBox:hover {
    border: 1px solid #0078D7;
}

QLineEdit:disabled, QTextEdit:disabled, QComboBox:disabled {
    background-color: #F0F0F0;
    color: #666666;
    border: 1px solid #C0C0C0;
}

/* コンボボックス */
QComboBox {
    padding-right: 20px;
}

QComboBox::drop-down {
    border: none;
    width: 20px;
}

QComboBox::down-arrow {
    image: url(resources/down_arrow.png);
    width: 12px;
    height: 12px;
}

/* ラベル */
QLabel {
    color: #333333;
    padding: 2px 0;
}

/* フォームレイアウト */
QFormLayout {
    spacing: 6px;
}

/* ボタン */
QPushButton {
    background-color: #0078D7;
    color: white;
    border: none;
    padding: 6px 12px;
    border-radius: 2px;
    min-width: 80px;
}

QPushButton:hover {
    background-color: #1E88E5;
}

QPushButton:pressed {
    background-color: #0056B3;
}

QPushButton:disabled {
    background-color: #CCCCCC;
    color: #666666;
}

/* プレビューエリア */
#preview_area {
    background-color: #FFFFFF;
    border: 1px solid #C0C0C0;
    border-radius: 2px;
}

/* スクロールバー */
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

QScrollBar::add-line:vertical,
QScrollBar::sub-line:vertical {
    border: none;
    background: none;
}

/* 日付入力コンボボックス */
QComboBox[dateField="true"] {
    max-width: 60px;
}

/* プレビューラベル */
QLabel#preview_label {
    font-weight: bold;
    color: #000000;
    background-color: #F0F0F0;
    padding: 4px;
    border-bottom: 1px solid #C0C0C0;
}

/* CTIフォーマットプレビュー */
QTextEdit#cti_preview {
    font-family: 'MS Gothic', monospace;
    font-size: 12px;
    border: 1px solid #C0C0C0;
    background-color: #FFFFFF;
}

QProgressBar {
    border: none;
    background-color: #E3F2FD;
    text-align: center;
}

QProgressBar::chunk {
    background-color: #0078D7;
}

QToolTip {
    background-color: #FFFFFF;
    color: #333333;
    border: 1px solid #C0C0C0;
    padding: 4px;
}

QMenu {
    background-color: #FFFFFF;
    border: 1px solid #C0C0C0;
}

QMenu::item {
    padding: 6px 24px;
}

QMenu::item:selected {
    background-color: #0078D7;
    color: #FFFFFF;
}

QMenuBar {
    background-color: #F0F0F0;
}

QMenuBar::item {
    padding: 6px 12px;
}

QMenuBar::item:selected {
    background-color: #0078D7;
    color: #FFFFFF;
}
"""

# 提供エリア検索結果のスタイル
AREA_SEARCH_AVAILABLE = """
    color: #000000;
    background-color: #FFFFFF;
    padding: 4px;
    border-radius: 2px;
    border: 1px solid #CCCCCC;
    margin: 2px 0;
"""

AREA_SEARCH_UNAVAILABLE = """
    color: #000000;
    background-color: #FFFFFF;
    padding: 4px;
    border-radius: 2px;
    border: 1px solid #CCCCCC;
    margin: 2px 0;
"""

AREA_SEARCH_ERROR = """
    color: #000000;
    background-color: #FFFFFF;
    padding: 4px;
    border-radius: 2px;
    border: 1px solid #CCCCCC;
    margin: 2px 0;
""" 