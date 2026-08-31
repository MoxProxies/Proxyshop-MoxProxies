"""
* Utils: Adobe Photoshop
"""
# Standard Library
from _ctypes import COMError, ArgumentError
from contextlib import suppress
from ctypes import c_uint32
from datetime import date
from functools import cache, cached_property
from os import environ
import re
from typing import Union, Any, Optional, TypedDict, Callable
import winreg

# Third Party
from comtypes.client.lazybind import Dispatch
from packaging.version import parse, Version
from photoshop.api import (
    ActionDescriptor,
    ActionReference,
    Application,
    DialogModes,
    PhotoshopPythonAPIError,
    Units)
from photoshop.api._artlayer import ArtLayer
from photoshop.api._core import Photoshop
from photoshop.api._document import Document
from photoshop.api._layerSet import LayerSet
from photoshop.api.constants import PHOTOSHOP_VERSION_MAPPINGS
from win32api import FormatMessage

# Local Imports
from src._state import AppEnvironment

"""
* Types & Definitions
"""

# Common Layer Objects
LayerContainer = LayerSet, Document, Dispatch
LayerObject = LayerSet, ArtLayer, Dispatch

# Common Layer Types
LayerContainerTypes = Union[LayerSet, Document, Dispatch]
LayerObjectTypes = Union[ArtLayer, LayerSet, Dispatch]

# Common Photoshop Exceptions
PS_EXCEPTIONS = (
    PhotoshopPythonAPIError,
    ArgumentError,
    COMError,
    AttributeError,
    IndexError,
    KeyError,
    ValueError,
    TypeError,
    OSError
)

PS_ERROR_CODES: dict[int: str] = {
    # --> COMError Messages that contain a message string
    # Response: "The message filter indicated that the application is busy."
    -2147417846: "Photoshop is currently busy, close any dialog boxes and stop any pending actions.",
    # Response: "The remote procedure call failed."
    -2147023170: "Unable to make connection with Photoshop, please check the FAQ for solutions.",
    # Response: "Invalid index."
    -2147352565: "Failed to load a PSD template or other file, ensure template file isn't corrupted "
                 "and that you have allocated enough scratch disk space and RAM to Photoshop.",
    # Response: "Exception occurred."
    -2147352567: "Photoshop does not appear to be installed. If Photoshop is installed, check the FAQ for solutions.",

    # --> COMError Messages that don't contain a message string, but have been investigated
    # Reference: https://docs.google.com/document/d/1j5xkWCWeHEFUZUaVtF59ccvAm9zTFsZ1qJcmFsXvkCM
    -2147220261: "Invalid data type passed to action descriptor function.",
    # Reference: https://docs.google.com/document/d/1LWXWyMa1kXAcGp4mDBZlpvqIgaR5jlk-H6j7uaKvOvI
    -2147213497: "Tried to transform, select, or translate an empty layer.",
    # Reference: https://docs.google.com/document/d/1mMeqi2lSaq2oUm1khl9rC0k9QLSI556a6m-MmLk19nw
    -2147212704: "Action descriptor or layer object key/property is missing.",
    # Reference: https://docs.google.com/document/d/1Oz69nNO0jR9qBbhjv3SVlMmk8iRnZY1LG7VaX7pqB-U
    -2147220262: "Photoshop tried to load a PSD template or file that doesn't exist.",

    # --> COMError Messages that don't contain a message string, but have been identified with testing
    # Test case: Pass a value to layer.textItem.color that isn't a SolidColor object
    -2147220279: "Wrong type of value passed to a Photoshop object property.",
    # Test case: Try to access the textItem property of a layer that isn't a TextLayer
    # Also: Observed when accessing the textItem property of a layer that contains an uninstalled font
    -2147213327: "Tried to interact with a text layer that is rasterized or has an uninstalled font.",
    # Test case: Delete a layer object, then try to delete it again.
    -2147213404: "Tried to delete a layer that doesn't exist."
}


# Layer bounds: left, top, right, bottom
LayerBounds = tuple[int, int, int, int]


class LayerDimensions(TypedDict):
    """Calculated layer dimension info for a layer."""
    width: int
    height: int
    center_x: int
    center_y: int
    left: int
    right: int
    top: int
    bottom: int


"""
* Photoshop Version Support
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

# Extend the mappings bundled with `photoshop-python-api`, which only cover the releases
# known when that package was published. Existing entries are left untouched.
for _year, _app_id in PS_VERSION_MAPPINGS.items():
    PHOTOSHOP_VERSION_MAPPINGS.setdefault(_year, _app_id)


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


"""
* Util Classes
"""


class ApplicationHandler(Application):
    """Wrapper for the Photoshop Application class."""

    def __init__(self, env: Optional[AppEnvironment] = None):

        # Normalize the requested version, checking the app environment first
        version = normalize_photoshop_version(env.PS_VERSION if env else None)
        version = version or normalize_photoshop_version(environ.get('PS_VERSION'))

        # `photoshop-python-api` reads this variable before the version passed to it
        if version:
            environ['PS_VERSION'] = version

        super().__init__(version=version)
        self._env = env

        # Set error dialog state
        with suppress(Exception):
            self.displayDialogs = DialogModes.DisplayErrorDialogs if (
                env.PS_ERROR_DIALOG
            ) else DialogModes.DisplayNoDialogs

    """
    * Handler Properties
    """

    @cached_property
    def _env(self) -> Optional[AppEnvironment]:
        """AppEnvironment: Global app environment object."""
        return

    def is_error_dialog_enabled(self) -> bool:
        """bool: Whether to allow error dialogs, defined in app environment object."""
        if self._env:
            return self._env.PS_ERROR_DIALOG
        return False

    """
    * Version Discovery
    """

    def _get_photoshop_versions(self) -> list[str]:
        """Retrieves the COM program ID of each Photoshop install listed in the Windows
        registry, ordered from newest release to oldest.

        Notes:
            Overrides the `photoshop-python-api` implementation, which orders the ID's as
            strings. String ordering can rank an older install above a newer one, e.g. the
            Photoshop CS6 ID '60' would be chosen over the Photoshop 2025 ID '190'.

        Returns:
            List of Photoshop COM program ID's, with a blank ID as a final fallback.
        """
        versions: list[str] = []
        with suppress(Exception):
            key = self._open_key(self._reg_path)
            for i in range(winreg.QueryInfoKey(key)[0]):
                app_id = winreg.EnumKey(key, i).split('.')[0]
                if app_id.isdigit() and app_id not in versions:
                    versions.append(app_id)
        if not versions:
            self._logger.debug('Unable to find Photoshop version number in HKEY_LOCAL_MACHINE registry!')

        # Newest release first, version-less program ID as a final fallback
        return [*sorted(versions, key=int, reverse=True), '']


class PhotoshopHandler(ApplicationHandler):
    """Wrapper for a single global Photoshop Application object equipped with soft loading,
    caching mechanisms, environment settings, and more."""
    DIMS_1200 = (3264, 4440)
    DIMS_800 = (2176, 2960)
    DIMS_600 = (1632, 2220)
    _version: Optional[Version] = None
    _instance = None

    def __new__(cls, env: Optional[Any] = None) -> 'PhotoshopHandler':
        """Always return the same Photoshop Application instance on successive calls.

        Args:
            env (AppEnvironment): Global app environment containing relevant env variables.

        Returns:
            The existing or newly created PhotoshopHandler instance.
        """
        # Use existing Photoshop instance or create new one
        if cls._instance is None:
            try:
                cls._instance = super().__new__(cls)
            except PS_EXCEPTIONS:
                cls._instance = super(Photoshop, cls).__new__(cls)

        # Establish the app environment object
        return cls._instance

    """
    * Managing the application object
    """

    def refresh_app(self):
        """Replace the existing Photoshop Application instance with a new one."""
        if not self.is_running():
            try:
                # Load Photoshop and default preferences
                self._version = None
                super(PhotoshopHandler, self).__init__(env=self._env)
                self.preferences.rulerUnits = Units.Pixels
                self.preferences.typeUnits = Units.Points
            except Exception as e:
                # Photoshop is either busy or unresponsive
                return OSError(get_photoshop_error_message(e))
        return

    """
    * Class Methods
    """

    @classmethod
    def is_running(cls) -> bool:
        """Check if the current Photoshop Application instance is still valid."""
        with suppress(Exception):
            _ = cls._instance.version
            return True
        return False

    """
    * Action Descriptor ID Conversions
    """

    @cache
    def charIDToTypeID(self, index: str) -> int:
        """Caching handler for charIDToTypeID.

        Args:
            index: Char ID to convert to Type ID.

        Returns:
            Type ID converted from Char ID.
        """
        return super().charIDToTypeID(index)

    @cache
    def CharIDToTypeID(self, index: str) -> int:
        """Uppercase redirect for charIDToTypeID."""
        return self.charIDToTypeID(index)

    @cache
    def cID(self, index: str) -> int:
        """Shorthand redirect for charIDToTypeID."""
        return self.charIDToTypeID(index)

    @cache
    def typeIDToCharID(self, index: int) -> str:
        """Caching handler for typeIDToCharID.

        Args:
            index: Type ID to convert to Char ID.

        Returns:
            Character representation of Type ID.
        """
        return super().typeIDToCharID(index)

    @cache
    def t2c(self, index: int) -> str:
        """Shorthand redirect for typeIDToCharID."""
        return self.typeIDToCharID(index)

    """
    * String ID Conversions
    """

    @cache
    def stringIDToTypeID(self, index: str) -> int:
        """Caching handler for stringIDToTypeID.

        Args:
            index: String ID to convert to Type ID.

        Returns:
            Type ID converted from string ID.
        """
        return super().stringIDToTypeID(index)

    @cache
    def StringIDToTypeID(self, index: str) -> int:
        """Uppercase redirect for stringIDTotypeID."""
        return self.stringIDToTypeID(index)

    @cache
    def sID(self, index: str) -> int:
        """Shorthand redirect for stringIDToTypeID."""
        return self.stringIDToTypeID(index)

    @cache
    def typeIDToStringID(self, index: int) -> str:
        """Caching handler for typeIDToStringID.

        Args:
            index: Type ID to convert to String ID.

        Returns:
            str: String representation of Type ID.
        """
        return super().typeIDToStringID(index)

    @cache
    def t2s(self, index: int) -> str:
        """Shorthand redirect for typeIDToStringID."""
        return self.typeIDToStringID(index)

    """
    * String / Char ID Conversions
    """

    @cache
    def charIDToStringID(self, index: int) -> str:
        """Converts a Char ID to a String ID.

        Args:
            index: Char ID to convert to String ID.

        Returns:
            str: String representation of Char ID.
        """
        return self.typeIDToStringID(
            self.charIDToTypeID(index))

    @cache
    def stringIDToCharID(self, index: int) -> str:
        """Converts a String ID to a Char ID.

        Args:
            index: String ID to convert to Char ID.

        Returns:
            str: Character representation of String ID.
        """
        return self.typeIDToCharID(
            self.stringIDToTypeID(index))

    """
    * Executing Action Descriptors
    """

    def executeAction(
        self, event_id: int,
        descriptor: ActionDescriptor,
        dialogs: DialogModes = DialogModes.DisplayNoDialogs
    ) -> Any:
        """Middleware to allow all dialogs when an error occurs upon calling executeAction in development mode.

        Args:
            event_id: Action descriptor event ID.
            descriptor: Main action descriptor tree to execute.
            dialogs: DialogMode which governs whether to display dialogs.

        Returns:
            Result of the action descriptor execution.
        """
        if self.is_error_dialog_enabled():
            # Allow error dialogs if enabled in the app environment
            return super().executeAction(event_id, descriptor, DialogModes.DisplayErrorDialogs)
        return super().executeAction(event_id, descriptor, dialogs)

    def ExecuteAction(
            self, event_id: int,
            descriptor: ActionDescriptor,
            dialogs: DialogModes = DialogModes.DisplayNoDialogs
    ) -> Any:
        """Utility definition rerouting to original `executeAction`."""
        self.executeAction(event_id, descriptor, dialogs)

    """
    * Version Checks
    """

    def get_version_number(self) -> Optional[Version]:
        """Gets the version number reported by the running Photoshop application.

        Returns:
            Parsed version number, e.g. 26.1.0 for Photoshop 2025, or None if Photoshop
                isn't reachable or reported a version that couldn't be parsed.
        """
        if self._version is None:
            with suppress(Exception):
                # Photoshop can report build info alongside the version, e.g. "26.1.0 20241021.r.55"
                self._version = parse(str(self.version).strip().split(' ')[0])
        return self._version

    def get_version_year(self) -> Optional[str]:
        """Gets the release year of the running Photoshop application.

        Returns:
            Release year, e.g. '2025', or None if the version couldn't be determined.
        """
        version = self.get_version_number()
        if version is None:
            return None
        return str(version.major + PS_YEAR_OFFSET)

    def supports_target_text_replace(self) -> bool:
        """bool: Checks if Photoshop version supports targeted text replacement."""
        return self.version_meets_requirement('22.0.0')

    def supports_webp(self) -> bool:
        """bool: Checks if Photoshop version supports WEBP files."""
        return self.version_meets_requirement('23.2.0')

    def supports_generative_fill(self) -> bool:
        """bool: Checks if Photoshop version supports Generative Fill."""
        return self.version_meets_requirement('24.6.0')

    def version_meets_requirement(self, value: str) -> bool:
        """Checks if Photoshop version meets or exceeds required value.

        Args:
            value: Minimum version string required, e.g. '23.2.0'.

        Returns:
            True if the running version meets the requirement, otherwise False. Assumes
                the requirement is met when the version can't be determined, so that a
                future release reporting an unexpected version string isn't mistaken
                for an unsupported one.
        """
        version = self.get_version_number()
        if version is None:
            return True
        return version >= parse(value)

    """
    * Dimensions
    """

    @cache
    def scale_by_dpi(self, value: Union[int, float]) -> int:
        """Scales a value by comparing document DPI to ideal DPI.

        Args:
            value: Integer or float value to adjust by DPI ratio.

        Returns:
            Adjusted value as an integer.
        """
        return int((self.activeDocument.width / 3264) * value)


class ReferenceLayer(ArtLayer):
    """A static ArtLayer whose properties such as width or height are not going to change. Most often
    used as a reference to position or size other layers."""

    def __init__(self, parent: Any = None, app: PhotoshopHandler = None):
        self._global_app = app if app else PhotoshopHandler()
        super().__init__(parent=parent)

    """
    * API Methods
    """

    def duplicate(self, relativeObject=None, insertionLocation=None):
        """Duplicates the layer and returns it as a `ReferenceLayer` object."""
        return ReferenceLayer(self.app.duplicate(relativeObject, insertionLocation))

    """
    * Cached Conversions
    """

    @cache
    def sID(self, index: str) -> int:
        """Caching handler for stringIDToTypeID on the global application object.

        Args:
            index: String ID to convert to Type ID.

        Returns:
            Type ID converted from string ID.
        """
        return self._global_app.stringIDToTypeID(index)

    """
    * Layer Properties
    """

    @cached_property
    def id(self) -> int:
        """int: This layer's ID (cached)."""
        return self.app.id

    @cached_property
    def action_getter(self) -> ActionReference:
        """Gets action descriptor info object for this layer.

        Returns:
            Action descriptor info object about the layer.
        """
        ref = ActionReference()
        ref.putIdentifier(self.sID('layer'), self.id)
        return self._global_app.executeActionGet(ref)

    """
    * Layer Bounds
    """

    @cached_property
    def bounds(self) -> LayerBounds:
        """LayerBounds: Bounds of the layer (left, top, right, bottom)."""
        return self.app.bounds

    @cached_property
    def bounds_no_effects(self) -> LayerBounds:
        """LayerBounds: Bounds of the layer (left, top, right, bottom) without layer effects applied."""
        with suppress(Exception):
            d = self.action_getter
            try:
                # Try getting bounds no effects
                bounds = d.getObjectValue(self.sID('boundsNoEffects'))
            except PS_EXCEPTIONS:
                # Try getting bounds
                bounds = d.getObjectValue(self.sID('bounds'))
            return (
                bounds.getInteger(self.sID('left')),
                bounds.getInteger(self.sID('top')),
                bounds.getInteger(self.sID('right')),
                bounds.getInteger(self.sID('bottom')))
        # Fallback to layer object bounds property
        return self.bounds

    """
    * Layer Dimensions
    """

    @cached_property
    def dims(self) -> type[LayerDimensions]:
        """LayerDimensions: Returns dimensions of the layer (cached), including:
            - bounds (left, right, top, bottom)
            - height
            - width
            - center_x
            - center_y
        """
        return self.get_dimensions_from_bounds(self.bounds)

    @cached_property
    def dims_no_effects(self) -> type[LayerDimensions]:
        """LayerDimensions: Returns dimensions of the layer (cached) without layer effects applied, including:
            - bounds (left, right, top, bottom)
            - height
            - width
            - center_x
            - center_y
        """
        return self.get_dimensions_from_bounds(self.bounds_no_effects)

    """
    * Utility Methods
    """

    @staticmethod
    def get_dimensions_from_bounds(bounds) -> type[LayerDimensions]:
        """Compute width and height based on a set of bounds given.

        Args:
            bounds: List of bounds given.

        Returns:
            Dict containing height, width, and positioning locations.
        """
        width = int(bounds[2] - bounds[0])
        height = int(bounds[3] - bounds[1])
        return LayerDimensions(
            width=width,
            height=height,
            center_x=round((width / 2) + bounds[0]),
            center_y=round((height / 2) + bounds[1]),
            left=int(bounds[0]), right=int(bounds[2]),
            top=int(bounds[1]), bottom=int(bounds[3]))


"""
* Utility Decorators
"""


def try_photoshop(func) -> Callable:
    """Decorator to handle trying to run a Photoshop action but allowing exceptions to fail silently.

    Args:
        func: Function being wrapped.

    Returns:
        The wrapped function.
    """
    def wrapper(self, *args, **kwargs):
        try:
            result = func(self, *args, **kwargs)
            return result
        except PS_EXCEPTIONS:
            return
    return wrapper


"""
* Utility Funcs
"""


def get_photoshop_error_message(err: Exception) -> str:
    """Gets a user-facing error message based on a given Photoshop access exception.

    Args:
        err: Exception object containing the reason an action failed.

    Returns:
        Proper user response for this exception.
    """
    return (
        "Photoshop is currently busy, close any dialogs and stop any actions.\n"
    ) if 'busy' in str(err).lower() else (
        "Photoshop does not appear to be installed on your system.\n"
        "Please close Proxyshop and install a fresh copy of Photoshop,\n"
        "if Photoshop is installed, view the FAQ for troubleshooting.\n"
    )


def get_com_error(signed_int: int) -> str:
    """Check for an error message for both the signed and unsigned version of a COMError code (HRESULT).

    Args:
        signed_int: Signed integer representing a COMError exception.

    Returns:
        The string error message associated with this COMError code.
    """
    try:
        err = FormatMessage(signed_int)
    except Exception as e:
        try:
            unsigned_int = c_uint32(signed_int).value
            err = FormatMessage(unsigned_int) or e.args[2]
        except Exception as e:
            err = e.args[2]
    return err
