"""
* Tests: Photoshop Version Utils

Covers the logic which decides how Proxyshop finds and evaluates a Photoshop install.
Photoshop itself can't be reached from CI, so these tests verify everything up to the
COM connection: the release mappings, the accepted `PS_VERSION` formats, the ordering of
detected installs, and the feature version gates.
"""
# Standard Library
from datetime import date
import importlib.util
from pathlib import Path
import sys

# Third Party
import pytest

# The module is loaded by path because importing `src.utils` executes `src/__init__.py`,
# which loads the entire Photoshop application stack and only runs on Windows.
MODULE_PATH = Path(__file__).parents[1] / 'src' / 'utils' / 'ps_version.py'
_spec = importlib.util.spec_from_file_location('ps_version', MODULE_PATH)
ps_version = importlib.util.module_from_spec(_spec)
sys.modules['ps_version'] = ps_version
_spec.loader.exec_module(ps_version)


"""
* Version Mappings
"""


@pytest.mark.parametrize('year, app_id', [
    # Mappings bundled with `photoshop-python-api`, the generated values must match
    ('2017', '110'), ('2018', '120'), ('2019', '130'), ('2020', '140'), ('2021', '150'),
    ('2022', '160'), ('2023', '170'), ('2024', '180'), ('2025', '190'),
    # Releases newer than any the bundled mappings cover
    ('2026', '200'), ('2027', '210')])
def test_version_mappings_match_known_program_ids(year, app_id):
    """Each release year maps to the COM program ID Photoshop registers."""
    assert ps_version.get_photoshop_version_mappings(int(year)).get(year) == app_id


def test_version_mappings_cover_next_year():
    """A release which ships before Proxyshop updates is still mapped."""
    assert str(date.today().year + 1) in ps_version.PS_VERSION_MAPPINGS


def test_version_mappings_start_at_oldest_supported_release():
    """Releases older than the supported range aren't mapped."""
    assert '2016' not in ps_version.PS_VERSION_MAPPINGS
    assert ps_version.PS_VERSION_MAPPINGS['2017'] == '110'


"""
* Version Normalization
"""


@pytest.mark.parametrize('value, expected', [
    # Release year
    ('2026', '2026'), ('2025', '2025'), (2025, '2025'), (' 2024 ', '2024'), ('CC 2019', '2019'),
    # Internal version number
    ('27', '2026'), ('26', '2025'), ('26.1.0', '2025'), ('25.0', '2024'), ('18', '2017'),
    # COM program ID
    ('200', '2026'), ('190', '2025'), ('110', '2017'),
    # Empty or unrecognized, falls back to automatic detection
    (None, None), ('', None), ('null', None), ('latest', None), ('CS6', None),
    ('2016', None), ('2099', None)])
def test_normalize_photoshop_version(value, expected):
    """A version is accepted as a release year, internal version, or COM program ID."""
    assert ps_version.normalize_photoshop_version(value) == expected


"""
* Version Discovery
"""


def test_sort_app_ids_orders_newest_release_first():
    """Program ID's are ordered numerically, not as strings.

    Notes:
        Ordering as strings ranks the Photoshop CS6 ID '60' above the 2025 ID '190',
        which would connect Proxyshop to the older install.
    """
    assert ps_version.sort_photoshop_app_ids(['60', '180', '200', '190']) == ['200', '190', '180', '60']


def test_sort_app_ids_discards_junk_and_duplicates():
    """Registry keys which aren't version numbers are ignored."""
    assert ps_version.sort_photoshop_app_ids(['190', 'junk', '190', '', '180']) == ['190', '180']


def test_sort_app_ids_with_nothing_found():
    """No detected installs yields no program ID's."""
    assert ps_version.sort_photoshop_app_ids([]) == []


"""
* Version Parsing
"""


@pytest.mark.parametrize('value, expected', [
    ('27.10.0', '27.10.0'),
    ('26.0', '26.0'),
    # Photoshop can report build info alongside the version number
    ('26.1.0 20241021.r.55', '26.1.0'),
    (' 25.0.0 ', '25.0.0')])
def test_parse_photoshop_version(value, expected):
    """A reported version is parsed, ignoring any build info."""
    assert str(ps_version.parse_photoshop_version(value)) == expected


@pytest.mark.parametrize('value', [None, '', 'not-a-version'])
def test_parse_photoshop_version_unreadable(value):
    """An unreadable version yields None rather than raising."""
    assert ps_version.parse_photoshop_version(value) is None


@pytest.mark.parametrize('value, year', [
    ('27.10.0', '2026'), ('26.1.0', '2025'), ('25.0.0', '2024'), ('18.0.0', '2017')])
def test_get_photoshop_version_year(value, year):
    """A version number resolves to its release year."""
    assert ps_version.get_photoshop_version_year(ps_version.parse_photoshop_version(value)) == year


def test_get_photoshop_version_year_unknown():
    """An unknown version has no release year."""
    assert ps_version.get_photoshop_version_year(None) is None


"""
* Feature Requirements
"""


@pytest.mark.parametrize('value, text_replace, webp, generative_fill', [
    ('27.10.0', True, True, True),
    ('26.1.0 20241021.r.55', True, True, True),
    ('24.6.0', True, True, True),
    ('24.0.0', True, True, False),
    ('23.2.0', True, True, False),
    ('22.5.1', True, False, False),
    ('21.0.0', False, False, False)])
def test_feature_requirements(value, text_replace, webp, generative_fill):
    """Each feature is gated on the release which introduced it."""
    version = ps_version.parse_photoshop_version(value)
    check = ps_version.check_version_requirement
    assert check(version, ps_version.PS_VERSION_TARGET_TEXT_REPLACE) is text_replace
    assert check(version, ps_version.PS_VERSION_WEBP) is webp
    assert check(version, ps_version.PS_VERSION_GENERATIVE_FILL) is generative_fill


def test_unknown_version_assumes_feature_support():
    """An undeterminable version is treated as a modern release.

    Notes:
        A future release reporting an unexpected version string must not have its
        supported features switched off.
    """
    assert ps_version.check_version_requirement(None, ps_version.PS_VERSION_WEBP) is True
