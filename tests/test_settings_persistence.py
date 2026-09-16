import json
import os
from types import SimpleNamespace

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QWidget

from ui.main_window import MainWindow
from ui.settings_dialog import SettingsDialog


def app():
    return QApplication.instance() or QApplication([])


def _settings(template="CUSTOM TEMPLATE", version="old-version"):
    return {
        "mode": "simple",
        "show_mode_selection": False,
        "format_template": template,
        "format_template_simple": template,
        "format_template_corporate": "CORPORATE TEMPLATE {call_preference}",
        "template_defaults_version": version,
        "selected_product": "jcom",
    }


def test_selected_product_is_saved_without_losing_other_settings(tmp_path):
    path = tmp_path / "settings.json"
    path.write_text(json.dumps({"other_setting": 123}), encoding="utf-8")
    owner = SimpleNamespace(settings_file=str(path), settings={})

    assert MainWindow._persist_selected_product(owner, "jcom")

    saved = json.loads(path.read_text(encoding="utf-8"))
    assert saved["selected_product"] == "jcom"
    assert saved["other_setting"] == 123


def test_load_settings_restores_product_and_live_comment_template(tmp_path):
    path = tmp_path / "settings.json"
    path.write_text(json.dumps(_settings()), encoding="utf-8")
    owner = SimpleNamespace(
        settings_file=str(path),
        settings={},
        current_product="self_collabo",
        format_template="STALE TEMPLATE",
    )

    MainWindow.load_settings(owner)

    assert owner.current_product == "jcom"
    assert owner.format_template == "CUSTOM TEMPLATE"
    assert owner.settings["format_template"] == "CUSTOM TEMPLATE"
    assert json.loads(path.read_text(encoding="utf-8"))["format_template_simple"] == "CUSTOM TEMPLATE"


def test_settings_dialog_uses_parent_settings_path(tmp_path):
    application = app()
    path = tmp_path / "settings.json"
    path.write_text(json.dumps(_settings(version="3.8.7")), encoding="utf-8")
    parent = QWidget()
    parent.settings_file = str(path)
    parent.current_mode = "simple"

    dialog = SettingsDialog(parent)

    assert dialog.settings_file == str(path.resolve())
    dialog.format_edit.setPlainText("EDITED IN DIALOG")
    assert dialog.save_settings()
    assert json.loads(path.read_text(encoding="utf-8"))["format_template_simple"] == "EDITED IN DIALOG"

    owner = SimpleNamespace(
        settings_file=str(path),
        settings={},
        current_product="self_collabo",
        format_template="STALE TEMPLATE",
    )
    MainWindow.load_settings(owner)
    assert owner.format_template == "EDITED IN DIALOG"
    dialog.close()
    parent.close()
