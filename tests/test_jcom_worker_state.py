import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import threading
import time

from PySide6.QtWidgets import QApplication, QWidget

from services.jcom_simulation_models import AddressCandidate, AddressCandidateRequest
from ui.jcom_simulation_worker import JcomSimulationWorker


def test_cancel_sets_dedicated_event():
    worker = JcomSimulationWorker(criteria=None)
    assert not worker.cancel_event.is_set()
    worker.cancel()
    assert worker.cancel_event.is_set()


def test_worker_factory_receives_candidate_prompt_resolver():
    observed = {}

    class FakeResult:
        pass

    class FakeService:
        def run(self, criteria):
            observed["criteria"] = criteria
            return FakeResult()

    def factory(cancel_event, progress, candidate_resolver):
        observed["cancel_event"] = cancel_event
        observed["candidate_resolver"] = candidate_resolver
        return FakeService()

    worker = JcomSimulationWorker(criteria="snapshot", service_factory=factory)
    worker.run()
    assert isinstance(observed["cancel_event"], threading.Event)
    assert callable(observed["candidate_resolver"])
    assert observed["criteria"] == "snapshot"


def test_candidate_reply_checks_stage_and_cancel_releases_waiter():
    worker = JcomSimulationWorker(criteria=None)
    request = AddressCandidateRequest(
        request_id="request-1",
        generation=4,
        stage_id="address-2",
        selected_address="選択済み",
        remaining_address="",
        candidates=[],
    )
    answers = []
    thread = threading.Thread(
        target=lambda: answers.append(worker._resolve_candidate(request))
    )
    thread.start()
    for _ in range(100):
        with worker._candidate_condition:
            if worker._pending_stage is not None:
                break
        time.sleep(0.01)
    assert not worker.submit_candidate("request-1", 4, "address-1", "old")
    assert worker.submit_candidate("request-1", 4, "address-2", "current")
    thread.join(timeout=2)
    assert answers == ["current"]

    answers.clear()
    thread = threading.Thread(
        target=lambda: answers.append(worker._resolve_candidate(request))
    )
    thread.start()
    for _ in range(100):
        with worker._candidate_condition:
            if worker._pending_stage is not None:
                break
        time.sleep(0.01)
    worker.cancel()
    thread.join(timeout=2)
    assert answers == [None]


def test_building_dialog_submits_only_user_selected_candidate():
    from ui.main_window import MainWindow

    application = QApplication.instance() or QApplication([])

    class Owner(QWidget):
        current_product = "jcom"
        jcom_active_request_id = "request-1"
        jcom_generation = 4

    class Worker:
        def __init__(self):
            self.submissions = []
            self.cancelled = False

        def isRunning(self):
            return True

        def submit_candidate(self, *parts):
            self.submissions.append(parts)

        def cancel(self):
            self.cancelled = True

    owner = Owner()
    worker = Worker()
    owner.jcom_worker = worker
    request = AddressCandidateRequest(
        request_id="request-1",
        generation=4,
        stage_id="address-2",
        selected_address="府中市美好町３丁目 → ３０番地",
        remaining_address="38",
        candidates=[
            AddressCandidate("2:0", "サニーハイツ"),
            AddressCandidate("2:1", "メゾン・ド・クラ２"),
        ],
    )
    MainWindow._on_jcom_candidate_requested(owner, worker, request)
    dialog = owner._jcom_candidate_dialog
    assert "サニーハイツ" in dialog._jcom_candidate_combo.currentText()
    dialog._jcom_candidate_combo.setCurrentIndex(1)
    dialog._jcom_select_button.click()
    application.processEvents()
    assert worker.submissions == [("request-1", 4, "address-2", "2:1")]
    assert not worker.cancelled
    owner.close()


def test_room_dialog_is_labeled_for_user_selection():
    from ui.main_window import MainWindow

    application = QApplication.instance() or QApplication([])

    class Owner(QWidget):
        current_product = "jcom"
        jcom_active_request_id = "request-room"
        jcom_generation = 5

    class Worker:
        def __init__(self):
            self.submissions = []
            self.cancelled = False

        def isRunning(self):
            return True

        def submit_candidate(self, *parts):
            self.submissions.append(parts)

        def cancel(self):
            self.cancelled = True

    owner = Owner()
    worker = Worker()
    owner.jcom_worker = worker
    request = AddressCandidateRequest(
        request_id="request-room",
        generation=5,
        stage_id="address-3",
        selected_address="富岡町３丁目 → ２７番地 → コーポＳＡＨ",
        remaining_address="206号",
        candidates=[
            AddressCandidate("3:0", "２０５号"),
            AddressCandidate("3:1", "２０６号"),
            AddressCandidate(
                "3:2", "【表示中の住所】で次へ", "next_with_current"
            ),
        ],
        selection_kind="room",
    )
    MainWindow._on_jcom_candidate_requested(owner, worker, request)
    dialog = owner._jcom_candidate_dialog
    assert dialog.windowTitle() == "J:COM 部屋番号の選択"
    assert "該当する部屋番号" in dialog._jcom_candidate_label.text()
    assert dialog._jcom_cancel_button.text() == "検索を中止"
    assert dialog._jcom_next_button.text() == "該当なし（表示中の住所で次へ）"
    dialog._jcom_candidate_combo.setCurrentIndex(1)
    dialog._jcom_select_button.click()
    application.processEvents()
    assert worker.submissions == [("request-room", 5, "address-3", "3:1")]
    assert not worker.cancelled
    owner.close()


def test_no_matching_room_submits_site_next_action_without_cancelling_search():
    from ui.main_window import MainWindow

    application = QApplication.instance() or QApplication([])

    class Owner(QWidget):
        current_product = "jcom"
        jcom_active_request_id = "request-next"
        jcom_generation = 6

    class Worker:
        submissions = []
        cancelled = False

        def isRunning(self):
            return True

        def submit_candidate(self, *parts):
            self.submissions.append(parts)

        def cancel(self):
            self.cancelled = True

    owner = Owner()
    worker = Worker()
    owner.jcom_worker = worker
    request = AddressCandidateRequest(
        request_id="request-next",
        generation=6,
        stage_id="address-4",
        selected_address="富岡町３丁目 → ２７番地 → コーポＳＡＨ",
        remaining_address="999号",
        candidates=[
            AddressCandidate("4:0", "２０６号"),
            AddressCandidate(
                "4:1", "【表示中の住所】で次へ", "next_with_current"
            ),
        ],
        selection_kind="room",
    )
    MainWindow._on_jcom_candidate_requested(owner, worker, request)
    dialog = owner._jcom_candidate_dialog
    assert dialog._jcom_cancel_button.text() == "検索を中止"
    dialog._jcom_next_button.click()
    application.processEvents()
    assert worker.submissions == [("request-next", 6, "address-4", "4:1")]
    assert not worker.cancelled
    owner.close()
