"""J:COM Selenium処理をGUIスレッドから分離するQThread。"""

from __future__ import annotations

import threading

from PySide6.QtCore import QThread, Signal, Slot

from services.jcom_simulation_models import JcomSearchCriteria
from services.jcom_simulation_service import JcomSimulationService


class JcomSimulationWorker(QThread):
    progress = Signal(str)
    result_ready = Signal(object)
    candidate_requested = Signal(object)

    def __init__(self, criteria: JcomSearchCriteria, parent=None, service_factory=None):
        super().__init__(parent)
        self.criteria = criteria
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
        if self._service_factory:
            service = self._service_factory(
                self.cancel_event, self.progress.emit, self._resolve_candidate
            )
        else:
            service = JcomSimulationService(
                cancel_event=self.cancel_event,
                progress=self.progress.emit,
                candidate_resolver=self._resolve_candidate,
            )
        self.result_ready.emit(service.run(self.criteria))

    @Slot()
    def cancel(self):
        self.cancel_event.set()
        with self._candidate_condition:
            self._candidate_condition.notify_all()
