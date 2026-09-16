import base64
import os
from pathlib import Path
import threading
import time
import uuid

from services.jcom_simulation_service import JcomSimulationService


class CaptureDriver:
    def __init__(self):
        self.clip = None

    def execute_script(self, script, *args):
        if "getBoundingClientRect" in script:
            return {"x": 10, "y": 20, "width": 769, "height": 1450}
        return None

    def execute_cdp_cmd(self, command, params):
        self.clip = params["clip"]
        png = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
        )
        return {"data": base64.b64encode(png).decode("ascii")}


class CaptureElement:
    def screenshot(self, path):
        raise AssertionError("CDP capture should be used before the fallback")


def test_screenshot_retention_removes_old_and_excess_files():
    tmp_path = Path.cwd() / f".test-jcom-screens-{uuid.uuid4().hex}"
    tmp_path.mkdir()
    service = JcomSimulationService(
        threading.Event(), lambda message: None, lambda request: None,
        screenshot_dir=tmp_path,
    )
    try:
        for index in range(23):
            path = tmp_path / f"request-{index}.png"
            path.write_bytes(b"png")
            stamp = time.time() - index
            os.utime(path, (stamp, stamp))
        old = tmp_path / "old-request.png"
        old.write_bytes(b"png")
        old_stamp = time.time() - 8 * 24 * 60 * 60
        os.utime(old, (old_stamp, old_stamp))

        service._prune_screenshots()

        assert not old.exists()
        assert len(list(tmp_path.glob("*.png"))) == 20
    finally:
        for path in tmp_path.glob("*.png"):
            path.unlink()
        tmp_path.rmdir()


def test_result_screenshot_captures_full_height_beyond_viewport(tmp_path):
    service = JcomSimulationService(threading.Event(), lambda message: None)
    service.driver = CaptureDriver()
    path = tmp_path / "result.png"

    assert service._capture_result_screenshot(CaptureElement(), path)
    assert path.stat().st_size > 0
    assert service.driver.clip["width"] == 769.0
    assert service.driver.clip["height"] == 1450.0
