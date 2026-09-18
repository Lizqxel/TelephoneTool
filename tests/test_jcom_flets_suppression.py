from types import SimpleNamespace

from ui.main_window import MainWindow


class VisibleWidget:
    def __init__(self):
        self.visible = None

    def setVisible(self, visible):
        self.visible = visible


def test_jcom_hides_all_flets_widgets_and_shows_jcom_panel():
    owner = SimpleNamespace(
        current_product="jcom",
        judgment_label=VisibleWidget(),
        judgment_combo=VisibleWidget(),
        area_search_btn=VisibleWidget(),
        area_result_container=VisibleWidget(),
        screenshot_btn=VisibleWidget(),
        jcom_panel=VisibleWidget(),
    )

    MainWindow._update_product_selector(owner)

    assert owner.jcom_panel.visible is True
    assert owner.judgment_label.visible is False
    assert owner.judgment_combo.visible is False
    assert owner.area_search_btn.visible is False
    assert owner.area_result_container.visible is False
    assert owner.screenshot_btn.visible is False


def test_self_collabo_restores_flets_widgets_and_hides_jcom_panel():
    owner = SimpleNamespace(
        current_product="self_collabo",
        judgment_label=VisibleWidget(),
        judgment_combo=VisibleWidget(),
        area_search_btn=VisibleWidget(),
        area_result_container=VisibleWidget(),
        screenshot_btn=VisibleWidget(),
        jcom_panel=VisibleWidget(),
    )

    MainWindow._update_product_selector(owner)

    assert owner.jcom_panel.visible is False
    assert owner.judgment_label.visible is True
    assert owner.judgment_combo.visible is True
    assert owner.area_search_btn.visible is True
    assert owner.area_result_container.visible is True
    assert owner.screenshot_btn.visible is True


def test_jcom_blocks_direct_flets_search():
    owner = SimpleNamespace(current_product="jcom", is_auto_processing=True)

    MainWindow.search_service_area(owner)

    assert owner.is_auto_processing is False


def test_jcom_cti_callback_fetches_customer_but_does_not_schedule_flets_search():
    calls = []
    owner = SimpleNamespace(
        current_product="jcom",
        is_auto_processing=False,
        fetch_cti_data=lambda: calls.append("fetch"),
    )

    MainWindow.on_cti_dialing_to_talking(owner)

    assert calls == ["fetch"]
    assert owner.is_auto_processing is False


def test_jcom_blocks_delayed_auto_search_entry_point():
    owner = SimpleNamespace(current_product="jcom", is_auto_processing=True)

    MainWindow.auto_search_service_area(owner)

    assert owner.is_auto_processing is False
