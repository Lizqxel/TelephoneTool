from types import SimpleNamespace

from ui.main_window import MainWindow


class VisibleWidget:
    def __init__(self):
        self.visible = None

    def setVisible(self, visible):
        self.visible = visible


def _owner(product):
    return SimpleNamespace(
        current_product=product,
        list_name_label=VisibleWidget(),
        list_name_input=VisibleWidget(),
        list_address_label=VisibleWidget(),
        list_address_input=VisibleWidget(),
    )


def test_unlisted_product_hides_list_name_and_address_fields():
    owner = _owner("self_collabo_unlisted")

    MainWindow._update_unlisted_list_inputs(owner)

    assert owner.list_name_label.visible is False
    assert owner.list_name_input.visible is False
    assert owner.list_address_label.visible is False
    assert owner.list_address_input.visible is False


def test_other_products_show_list_name_and_address_fields():
    for product in ("self_collabo", "jcom"):
        owner = _owner(product)

        MainWindow._update_unlisted_list_inputs(owner)

        assert owner.list_name_label.visible is True
        assert owner.list_name_input.visible is True
        assert owner.list_address_label.visible is True
        assert owner.list_address_input.visible is True
