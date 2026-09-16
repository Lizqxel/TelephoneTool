"""J:COM Selenium処理をGUIスレッドから分離するQThread。"""

from __future__ import annotations

import threading

from PySide6.QtCore import QThread, Signal, Slot

from services.jcom_simulation_models import JcomSearchCriteria
from services.jcom_simulation_service import JcomSimulationService


class JcomSimulationWorker(QThread):
    progress = Signal(str)
    result_ready = Signal(object)

    def __init__(self, criteria: JcomSearchCriteria, parent=None, service_factory=None):
        super().__init__(parent)
        self.criteria = criteria
        self.cancel_event = threading.Event()
        self._service_factory = service_factory

    def run(self):
        if self._service_factory:
            service = self._service_factory(
                self.cancel_event, self.progress.emit, None
            )
        else:
            service = JcomSimulationService(
                cancel_event=self.cancel_event,
                progress=self.progress.emit,
                candidate_resolver=None,
            )
        self.result_ready.emit(service.run(self.criteria))

    @Slot()
    def cancel(self):
        self.cancel_event.set()
