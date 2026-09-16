import threading
from ui.jcom_simulation_worker import JcomSimulationWorker


def test_cancel_sets_dedicated_event():
    worker = JcomSimulationWorker(criteria=None)
    assert not worker.cancel_event.is_set()
    worker.cancel()
    assert worker.cancel_event.is_set()


def test_worker_factory_receives_no_candidate_prompt_resolver():
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
    assert observed["candidate_resolver"] is None
    assert observed["criteria"] == "snapshot"
