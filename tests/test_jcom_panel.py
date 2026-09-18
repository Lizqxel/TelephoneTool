import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from datetime import datetime

from PySide6.QtGui import QColor, QPixmap
from PySide6.QtWidgets import QApplication

from services.jcom_simulation_models import (
    AgeBracket,
    JcomSimulationResult,
    JST,
    ResidenceType,
    SimulationStatus,
)
from ui.jcom_simulation_panel import JcomSimulationPanel


def app():
    return QApplication.instance() or QApplication([])


def test_panel_defaults_to_editable_apartment_and_no_result():
    application = app()
    panel = JcomSimulationPanel()
    assert panel.residence_type() is ResidenceType.APARTMENT
    assert panel.search_button.text() == "料金シミュレーション検索"
    assert not panel.tabs.isVisible()
    panel.close()


def test_panel_displays_zero_yen_and_missing_image_separately():
    application = app()
    panel = JcomSimulationPanel()
    result = JcomSimulationResult(
        request_id="request-1",
        generation=1,
        acquired_at=datetime.now(JST),
        input_address="入力住所",
        confirmed_address="確定住所",
        residence_type=ResidenceType.APARTMENT,
        age=22,
        calculated_age_bracket=AgeBracket.UNDER_22,
        applied_age_bracket=AgeBracket.UNDER_22,
        age_selection_used=True,
        selected_service="ネットのみ",
        line_type="J:COM NET 光(N)",
        course="光(N) 1Gコース",
        next_month_price_yen=0,
        base_price_yen=5258,
        raw_text="料金結果の原文",
        status=SimulationStatus.SUCCESS,
    )
    panel.show_result(result)
    assert "0円（税込）" in panel.result_text.toPlainText()
    assert panel.raw_text.toPlainText() == "料金結果の原文"
    assert "画像未取得" in panel.image_label.text()
    panel.close()


def test_panel_shows_loading_until_delayed_screenshot_arrives(tmp_path):
    application = app()
    image_path = tmp_path / "delayed-result.png"
    image = QPixmap(320, 640)
    image.fill(QColor("white"))
    assert image.save(str(image_path), "PNG")

    panel = JcomSimulationPanel()
    result = JcomSimulationResult(
        request_id="request-loading",
        generation=1,
        acquired_at=datetime.now(JST),
        input_address="入力住所",
        confirmed_address="確定住所",
        residence_type=ResidenceType.APARTMENT,
        age=30,
        calculated_age_bracket=AgeBracket.OVER_27,
        applied_age_bracket=None,
        age_selection_used=False,
        selected_service="ネットのみ",
        line_type="J:COM NET 光(N)",
        course="光(N) 1Gコース",
        screenshot_pending=True,
        status=SimulationStatus.SUCCESS,
    )

    panel.show_result(result)
    panel.tabs.setCurrentIndex(2)
    assert "生成しています" in panel.image_label.text()
    assert not panel.open_image_button.isEnabled()

    result.screenshot_pending = False
    result.screenshot_path = str(image_path)
    panel.update_screenshot(result)

    assert panel.tabs.currentIndex() == 2
    assert not panel._source_pixmap.isNull()
    assert panel.open_image_button.isEnabled()
    panel.close()


def test_partial_result_warns_that_price_uses_site_confirmed_address():
    application = app()
    panel = JcomSimulationPanel()
    result = JcomSimulationResult(
        request_id="request-partial",
        generation=1,
        acquired_at=datetime.now(JST),
        input_address="入力住所",
        confirmed_address="町域のみ",
        residence_type=ResidenceType.DETACHED,
        age=30,
        calculated_age_bracket=AgeBracket.OVER_27,
        applied_age_bracket=None,
        age_selection_used=False,
        selected_service="ネットのみ",
        line_type="J:COM NET 光(N)",
        course="光(N) 1Gコース",
        partial_address=True,
        status=SimulationStatus.PARTIAL,
    )
    panel.show_result(result)
    assert "住所の一部未確認" in panel.progress_label.text()
    panel.close()


def test_site_image_preview_fits_full_image_and_has_zoom_controls(tmp_path):
    application = app()
    image_path = tmp_path / "full-result.png"
    image = QPixmap(800, 1600)
    image.fill(QColor("white"))
    assert image.save(str(image_path), "PNG")

    panel = JcomSimulationPanel()
    panel.resize(900, 700)
    result = JcomSimulationResult(
        request_id="request-image",
        generation=1,
        acquired_at=datetime.now(JST),
        input_address="input",
        confirmed_address="confirmed",
        residence_type=ResidenceType.DETACHED,
        age=30,
        calculated_age_bracket=AgeBracket.OVER_27,
        applied_age_bracket=None,
        age_selection_used=False,
        selected_service="ネットのみ",
        line_type="J:COM NET 光(N)",
        course="光(N) 1Gコース",
        screenshot_path=str(image_path),
        status=SimulationStatus.SUCCESS,
    )
    panel.show_result(result)
    panel.show()
    panel.tabs.setCurrentIndex(2)
    application.processEvents()
    panel._fit_image_to_width()

    assert panel._source_pixmap.height() == 1600
    assert panel.image_label.pixmap().height() > panel.image_scroll.viewport().height()
    assert panel.fit_image_button.isEnabled()
    assert panel.zoom_in_button.isEnabled()
    assert panel.open_image_button.isEnabled()
    panel.close()
