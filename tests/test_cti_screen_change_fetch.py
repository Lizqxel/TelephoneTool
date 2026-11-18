from types import SimpleNamespace

from ui.main_window import MainWindow


CTI_NAME = "ヤマモト美容室"
CTI_ADDRESS = "愛知県西尾市一色町野田堤外２３"


class TextField:
    def __init__(self, text=""):
        self.value = text

    def text(self):
        return self.value

    def setText(self, text):
        self.value = text


class UndoStack:
    def beginMacro(self, _name):
        pass

    def endMacro(self):
        pass


class StatusBar:
    def showMessage(self, *_args):
        pass


def _owner(product="self_collabo"):
    owner = SimpleNamespace(
        current_product=product,
        current_mode="simple",
        list_name_input=TextField(),
        contractor_input=TextField(),
        address_input=TextField(),
        list_address_input=TextField(),
        list_phone_input=TextField(),
        postal_code_input=TextField(),
        list_postal_code_input=TextField(),
        unlisted_address_detail_input=TextField(),
        undo_stack=UndoStack(),
        fetch_count=0,
    )
    owner._push_text_change = lambda widget, text: widget.setText(text)
    owner._is_unlisted_self_collabo = lambda: owner.current_product == "self_collabo_unlisted"
    owner._prefecture_only = MainWindow._prefecture_only
    owner._clear_unlisted_address_candidates = lambda: None
    owner.statusBar = lambda: StatusBar()

    def fetch_cti_data():
        owner.fetch_count += 1
        data = SimpleNamespace(
            customer_name=CTI_NAME,
            address=CTI_ADDRESS,
            phone="0563734511",
            postal_code="444-0404",
        )
        MainWindow.update_form_with_data(owner, data)

    owner.fetch_cti_data = fetch_cti_data
    return owner


def _monitor(owner, name=CTI_NAME, address=CTI_ADDRESS):
    MainWindow._handle_cti_screen_values(owner, name, address)


def test_unlisted_address_built_from_postal_code_is_not_reverted():
    owner = _owner("self_collabo_unlisted")
    _monitor(owner)
    assert owner.fetch_count == 1
    assert owner.address_input.text() == "愛知県"

    # 郵便番号の候補と番地から住所を組み立てる
    owner.address_input.setText("愛知県西尾市一色町野田堤外２３")
    for _ in range(5):
        _monitor(owner)

    assert owner.fetch_count == 1
    assert owner.address_input.text() == "愛知県西尾市一色町野田堤外２３"


def test_manually_edited_address_is_kept_in_normal_product():
    owner = _owner("self_collabo")
    _monitor(owner)
    assert owner.fetch_count == 1

    owner.address_input.setText("愛知県西尾市一色町野田堤外２３　山本ビル")
    owner.list_name_input.setText("ヤマモト美容院")
    _monitor(owner)

    assert owner.fetch_count == 1
    assert owner.address_input.text() == "愛知県西尾市一色町野田堤外２３　山本ビル"
    assert owner.list_name_input.text() == "ヤマモト美容院"


def test_fetches_when_cti_shows_another_customer():
    owner = _owner("self_collabo_unlisted")
    _monitor(owner)
    owner.address_input.setText("愛知県西尾市一色町野田堤外２３")

    _monitor(owner, name="桃華", address="静岡県掛川市杉谷南１丁目１４−１８")

    assert owner.fetch_count == 2


def test_fetches_on_first_detection_and_when_ui_is_empty():
    owner = _owner("self_collabo")
    _monitor(owner)
    assert owner.fetch_count == 1

    owner.list_name_input.setText("")
    owner.address_input.setText("")
    _monitor(owner)

    assert owner.fetch_count == 2


def test_temporarily_empty_cti_field_is_not_treated_as_a_change():
    owner = _owner("self_collabo")
    _monitor(owner)

    _monitor(owner, name="", address=CTI_ADDRESS)
    _monitor(owner, name=CTI_NAME, address="")
    _monitor(owner, name="", address="")

    assert owner.fetch_count == 1


def _product_owner(product):
    return SimpleNamespace(
        current_product=product,
        jcom_generation=0,
        _cti_values_at_last_fetch=(CTI_NAME, CTI_ADDRESS),
        cancel_jcom_search=lambda: None,
        _persist_selected_product=lambda _product: True,
        _update_product_selector=lambda: None,
    )


def test_switching_to_or_from_unlisted_refetches_on_next_monitor_tick():
    for before, after in (
        ("self_collabo", "self_collabo_unlisted"),
        ("self_collabo_unlisted", "self_collabo"),
    ):
        owner = _product_owner(before)
        MainWindow.set_product(owner, after)
        assert owner._cti_values_at_last_fetch is None


def test_switching_between_listed_products_keeps_fetched_values():
    owner = _product_owner("self_collabo")
    MainWindow.set_product(owner, "jcom")
    assert owner._cti_values_at_last_fetch == (CTI_NAME, CTI_ADDRESS)
