from types import SimpleNamespace

from ui.main_window import MainWindow


class TextField:
    def __init__(self, text=""):
        self.value = text

    def text(self):
        return self.value

    def setText(self, text):
        self.value = text


class ComboBox:
    def __init__(self):
        self.value = ""

    def setCurrentText(self, text):
        self.value = text


def test_unlisted_product_sets_default_fee_when_input_is_empty():
    owner = SimpleNamespace(
        current_product="self_collabo_unlisted",
        fee_input=TextField(),
        fee_combo=ComboBox(),
    )

    MainWindow._apply_unlisted_fee_default(owner)

    assert owner.fee_combo.value == "2500円～3000円"
    assert owner.fee_input.text() == "2500円～3000円"


def test_unlisted_product_does_not_overwrite_entered_fee():
    owner = SimpleNamespace(
        current_product="self_collabo_unlisted",
        fee_input=TextField("3200円"),
        fee_combo=ComboBox(),
    )

    MainWindow._apply_unlisted_fee_default(owner)

    assert owner.fee_combo.value == ""
    assert owner.fee_input.text() == "3200円"


def test_other_product_does_not_set_default_fee():
    owner = SimpleNamespace(
        current_product="self_collabo",
        fee_input=TextField(),
        fee_combo=ComboBox(),
    )

    MainWindow._apply_unlisted_fee_default(owner)

    assert owner.fee_combo.value == ""
    assert owner.fee_input.text() == ""
