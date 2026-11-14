import pytest

from version import VERSION, is_newer_version, version_key


@pytest.mark.parametrize("latest, current", [
    ("3.9.5", "3.9.4"),
    ("3.10.0", "3.9.4"),   # 文字の並びで比べると逆になっていた
    ("3.10.0", "3.9.9"),
    ("4.0.0", "3.99.99"),
    ("3.9.10", "3.9.9"),
    ("3.9.5", "3.9"),
])
def test_newer_versions_are_compared_as_numbers(latest, current):
    assert is_newer_version(latest, current)
    assert not is_newer_version(current, latest)


def test_same_or_older_version_is_not_offered_as_an_update():
    assert not is_newer_version("3.9.5", "3.9.5")
    assert not is_newer_version("3.9.5", "3.9.5.0")
    assert not is_newer_version("3.9.4", "3.9.5")


def test_tag_prefixes_and_suffixes_are_ignored():
    assert version_key("v3.10.0") == (3, 10, 0)
    assert version_key("3.9.5-beta") == (3, 9, 5)

