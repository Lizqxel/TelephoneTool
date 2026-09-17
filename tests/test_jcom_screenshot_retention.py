import base64
from io import BytesIO
import os
from pathlib import Path
import threading
import time
import uuid

from PIL import Image

from services.jcom_simulation_service import JcomSimulationService


class CaptureDriver:
    def __init__(self, with_address=False):
        self.clips = []
        self.with_address = with_address
        self.address_element = object()

    def execute_script(self, script, *args):
        if "selected-address-box" in script:
            return ([{"element": self.address_element, "text": "東京都稲城市大丸２２１３番地\n再設定"}]
                    if self.with_address else [])
        if "getBoundingClientRect" in script:
            return {
                "result": {"x": 10, "y": 20, "width": 769, "height": 1450},
                "address": (
                    {"x": 20, "y": 5, "width": 729, "height": 72}
                    if self.with_address and args[2] is self.address_element else None
                ),
            }
        return None

    def execute_cdp_cmd(self, command, params):
        clip = params["clip"]
        self.clips.append(clip)
        color = "blue" if clip["height"] == 72 else "red"
        image = Image.new("RGB", (int(clip["width"]), int(clip["height"])), color)
        output = BytesIO()
        image.save(output, "PNG")
        return {"data": base64.b64encode(output.getvalue()).decode("ascii")}


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
    assert service.driver.clips[0]["width"] == 769.0
    assert service.driver.clips[0]["height"] == 1450.0
    assert not service._screenshot_address_included


def test_screenshot_places_site_address_above_full_price_result(tmp_path):
    service = JcomSimulationService(threading.Event(), lambda message: None)
    service.driver = CaptureDriver(with_address=True)
    path = tmp_path / "address-and-price.png"

    assert service._capture_result_screenshot(
        CaptureElement(), path, "東京都稲城市大丸２２１３番地"
    )
    assert service._screenshot_address_included
    assert len(service.driver.clips) == 2
    with Image.open(path) as image:
        assert image.size == (769, 72 + 16 + 1450)
        assert image.getpixel((10, 10)) == (0, 0, 255)
        assert image.getpixel((10, 100)) == (255, 0, 0)


def test_screenshot_does_not_attach_a_different_site_address(tmp_path):
    service = JcomSimulationService(threading.Event(), lambda message: None)
    service.driver = CaptureDriver(with_address=True)
    path = tmp_path / "result-only.png"

    assert service._capture_result_screenshot(
        CaptureElement(), path, "東京都稲城市大丸９９９９番地"
    )
    assert not service._screenshot_address_included
    assert len(service.driver.clips) == 1
    with Image.open(path) as image:
        assert image.size == (769, 1450)
