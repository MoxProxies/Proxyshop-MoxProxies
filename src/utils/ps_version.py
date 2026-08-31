"""
* Utils: Photoshop Versions

Identifying, comparing, and normalizing Photoshop version numbers. Kept free of any
Photoshop or Windows imports so this logic can be verified on any platform.
"""
# Standard Library
from contextlib import suppress
from datetime import date
import re
from typing import Any, Iterable, Optional

# Third Party
from packaging.version import parse, Version

"""
* Version Constants
"""

# Oldest Photoshop release Proxyshop supports.
PS_YEAR_MIN = 2017

# COM program ID of the oldest supported release, e.g. Photoshop.Application.110
PS_APP_ID_MIN = 110

# Each yearly Photoshop release raises the COM program ID by this amount, e.g.
# 2017 -> 110, 2020 -> 140, 2024 -> 180, 2025 -> 190, 2026 -> 200
PS_APP_ID_STEP = 10

# Difference between a Photoshop release year and its internal version number,
# e.g. Photoshop 2025 -> 26.0.0, Photoshop 2026 -> 27.0.0
PS_YEAR_OFFSET = 1999

# Newest release known at the time of writing, guarantees a mapping for it even
# if the system clock is set behind.
PS_YEAR_KNOWN = 2026

"""
* Feature Requirements
"""

# Minimum Photoshop version required for targeted text replacement
PS_VERSION_TARGET_TEXT_REPLACE = '22.0.0'

# Minimum Photoshop version required to open WEBP files
PS_VERSION_WEBP = '23.2.0'

# Minimum Photoshop version required for Generative Fill
PS_VERSION_GENERATIVE_FILL = '24.6.0'

"""
* Version Mappings
"""


def get_photoshop_app_id(year: int) -> str:
    """Calculates the COM program ID suffix used by a given Photoshop release.

    Args:
        year: Photoshop release year, e.g. 2025.

    Returns:
        COM program ID suffix, e.g. '190' for `Photoshop.Application.190`.
    """
    return str(PS_APP_ID_MIN + ((year - PS_YEAR_MIN) * PS_APP_ID_STEP))


def get_photoshop_version_mappings(year_max: Optional[int] = None) -> dict[str, str]:
    """Maps every supported Photoshop release year to its COM program ID.

    Args:
        year_max: Newest release year to generate a mapping for. Uses next year if not
            provided, so releases which ship ahead of a Proxyshop update are covered.

    Returns:
        Dict mapping release year to COM program ID, e.g. {'2025': '190'}.
    """
    year_max = max(year_max or (date.today().year + 1), PS_YEAR_KNOWN)
    return {str(year): get_photoshop_app_id(year) for year in range(PS_YEAR_MIN, year_max + 1)}


# Photoshop release year -> COM program ID, e.g. '2025' -> '190'
PS_VERSION_MAPPINGS: dict[str, str] = get_photoshop_version_mappings()

# COM program ID -> Photoshop release year, e.g. '190' -> '2025'
PS_APP_ID_MAPPINGS: dict[str, str] = {v: k for k, v in PS_VERSION_MAPPINGS.items()}


"""
* Version Utils
"""


def normalize_photoshop_version(value: Any) -> Optional[str]:
    """Normalizes a user provided Photoshop version to a release year, the key format
    used by the `photoshop-python-api` version mappings.

    Args:
        value: Photoshop version provided by the user. Accepts a release year, e.g.
            '2025' or 'CC 2019', an internal version number, e.g. '26' or '26.1.0',
            or a COM program ID, e.g. '190'.

    Returns:
        Photoshop release year, e.g. '2025', or None if the value was empty or
            couldn't be recognized as a supported release.
    """
    if value in (None, ''):
        return None
    value = str(value).strip()

    # Release year, e.g. '2025' or 'CC 2019'
    if found := re.search(r'\b(20\d{2})\b', value):
        year = found.group(1)
        return year if year in PS_VERSION_MAPPINGS else None

    # Leading number, e.g. '190' or '26.1.0'
    if found := re.match(r'^(\d+)', value):
        number = found.group(1)

        # COM program ID, e.g. '190'
        if number in PS_APP_ID_MAPPINGS:
            return PS_APP_ID_MAPPINGS[number]

        # Internal version number, e.g. '26'
        year = str(int(number) + PS_YEAR_OFFSET)
        if year in PS_VERSION_MAPPINGS:
            return year

    # Unrecognized version, fall back to automatic detection
    return None


def sort_photoshop_app_ids(app_ids: Iterable[Any]) -> list[str]:
    """Orders Photoshop COM program ID's from newest release to oldest, discarding
    duplicates and any value which isn't a number.

    Notes:
        The ID's must be ordered numerically. Ordering them as strings can rank an older
        install above a newer one, e.g. the Photoshop CS6 ID '60' above the 2025 ID '190'.

    Args:
        app_ids: COM program ID's to order, e.g. those found in the Windows registry.

    Returns:
        Ordered list of COM program ID's, newest release first.
    """
    unique = {str(app_id) for app_id in app_ids if str(app_id).isdigit()}
    return sorted(unique, key=int, reverse=True)


def parse_photoshop_version(value: Any) -> Optional[Version]:
    """Parses a version number reported by the Photoshop application.

    Args:
        value: Version reported by Photoshop, which can carry build info alongside the
            version number, e.g. '26.1.0 20241021.r.55'.

    Returns:
        Parsed version number, or None if the value couldn't be parsed.
    """
    with suppress(Exception):
        return parse(str(value).strip().split(' ')[0])
    return None


def get_photoshop_version_year(version: Optional[Version]) -> Optional[str]:
    """Gets the release year matching a Photoshop version number.

    Args:
        version: Parsed Photoshop version number, e.g. 26.1.0.

    Returns:
        Release year, e.g. '2025', or None if a version wasn't provided.
    """
    if version is None:
        return None
    return str(version.major + PS_YEAR_OFFSET)


def check_version_requirement(version: Optional[Version], required: str) -> bool:
    """Checks whether a Photoshop version meets or exceeds a required version.

    Args:
        version: Parsed Photoshop version number, or None if it couldn't be determined.
        required: Minimum version required, e.g. '23.2.0'.

    Returns:
        True if the version meets the requirement, otherwise False. An unknown version is
            assumed to meet it, so that a future release reporting an unexpected version
            string isn't mistaken for an unsupported one.
    """
    if version is None:
        return True
    return version >= parse(required)
