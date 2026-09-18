"""J:COM Selenium処理をGUIスレッドから分離するQThread。"""

from __future__ import annotations

import copy
import threading

from PySide6.QtCore import QThread, Signal, Slot

from services.jcom_simulation_models import JcomSearchCriteria
from services.jcom_simulation_service import JcomSimulationService


class JcomSimulationWorker(QThread):
    progress = Signal(str)
    result_ready = Signal(object)
    screenshot_ready = Signal(object)
    candidate_requested = Signal(object)

    def __init__(
        self,
        criteria: JcomSearchCriteria,
        parent=None,
        service_factory=None,
        headless=False,
    ):
        super().__init__(parent)
        self.criteria = criteria
        self.headless = bool(headless)
        self.cancel_event = threading.Event()
        self._service_factory = service_factory
        self._candidate_condition = threading.Condition()
        self._pending_stage = None
        self._candidate_answer = None

    def _resolve_candidate(self, request):
        stage = (request.request_id, request.generation, request.stage_id)
        with self._candidate_condition:
            self._pending_stage = stage
            self._candidate_answer = None
        self.candidate_requested.emit(request)
        with self._candidate_condition:
            while self._candidate_answer is None and not self.cancel_event.is_set():
                self._candidate_condition.wait(0.2)
            answer = self._candidate_answer
            self._pending_stage = None
            self._candidate_answer = None
            return None if self.cancel_event.is_set() else answer

    @Slot(str, int, str, str)
    def submit_candidate(self, request_id, generation, stage_id, candidate_id):
        with self._candidate_condition:
            if self._pending_stage != (request_id, generation, stage_id):
                return False
            self._candidate_answer = candidate_id
            self._candidate_condition.notify_all()
            return True

    def run(self):
        early_result_emitted = False

        def emit_result_early(result):
            nonlocal early_result_emitted
            early_result_emitted = True
            self.result_ready.emit(copy.deepcopy(result))

        def emit_screenshot_ready(result):
            self.screenshot_ready.emit(copy.deepcopy(result))

        if self._service_factory:
            service = self._service_factory(
                self.cancel_event, self.progress.emit, self._resolve_candidate
            )
        else:
            service = JcomSimulationService(
                cancel_event=self.cancel_event,
                progress=self.progress.emit,
                candidate_resolver=self._resolve_candidate,
                headless=self.headless,
            )
        service.result_ready_callback = emit_result_early
        service.screenshot_ready_callback = emit_screenshot_ready
        result = service.run(self.criteria)
        if not early_result_emitted:
            self.result_ready.emit(result)

    @Slot()
    def cancel(self):
        self.cancel_event.set()
        with self._candidate_condition:
            self._candidate_condition.notify_all()
