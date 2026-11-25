"""
* Tool: Photoshop COM Check

Reports how `comtypes` is talking to Photoshop on this machine.

`comtypes` has two ways of calling a COM object. Given a type library it uses
`lazybind.Dispatch`, which knows which names are methods and which are properties.
Without one it falls back to `dynamic._Dispatch`, which guesses, and guesses wrong on
Photoshop in several ways: collections can't be indexed by name, and a method fetched
as a property is run by Photoshop with no arguments, producing either a COM error or a
quiet `None` that fails later as "'NoneType' object is not callable".

Proxyshop works around all of that, but the workarounds are only needed on the second
kind. This says which kind you have.

Usage:
    poetry run python tools/check_photoshop_com.py
"""
# Standard Library
import os
import sys
import winreg

# Third Party
from comtypes.automation import IDispatch
from comtypes.client import CreateObject
from comtypes.typeinfo import LoadTypeLibEx, REGKIND_NONE
from comtypes.client.dynamic import _Dispatch as FullyDynamicDispatch
from comtypes.client.lazybind import Dispatch as LazyDispatch


def get_registered_app_ids() -> list[str]:
    """Collects the Photoshop COM program ID's listed in the Windows registry.

    Returns:
        List of program ID's, newest release first.
    """
    app_ids: list[str] = []
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Adobe\\Photoshop')
        app_ids = [winreg.EnumKey(key, i).split('.')[0]
                   for i in range(winreg.QueryInfoKey(key)[0])]
    except OSError as e:
        print(f'  Could not read the registry: {e}')
    return sorted((i for i in app_ids if i.isdigit()), key=int, reverse=True)


def read_key(root: int, path: str, value: str = '') -> str:
    """Reads a value from the Windows registry.

    Args:
        root: Registry root key.
        path: Path to the key.
        value: Name of the value to read, blank for the key's default value.

    Returns:
        The value, or a blank string if it couldn't be read.
    """
    try:
        with winreg.OpenKey(root, path) as key:
            return str(winreg.QueryValueEx(key, value)[0])
    except OSError:
        return ''


def get_install_path(app_id: str) -> str:
    """Finds where a Photoshop release is installed.

    Args:
        app_id: Photoshop COM program ID, e.g. '210'.

    Returns:
        Path to the install directory, or a blank string.
    """
    return read_key(
        winreg.HKEY_LOCAL_MACHINE,
        f'SOFTWARE\\Adobe\\Photoshop\\{app_id}.0', 'ApplicationPath')


def report_type_library(app_id: str) -> None:
    """Reports where Photoshop's type library is registered and whether it loads.

    Notes:
        `comtypes` guesses at method and property names only when it can't read a type
        library. 'Error loading type library/DLL' means the library is registered but
        the file behind it can't be opened, which is worth telling apart from it not
        being registered at all.

    Args:
        app_id: Photoshop COM program ID, e.g. '210'.
    """
    prog_id = f'Photoshop.Application.{app_id}'
    clsid = read_key(winreg.HKEY_CLASSES_ROOT, f'{prog_id}\\CLSID')
    print(f'  CLSID      : {clsid or "not registered"}')
    if not clsid:
        return

    lib_id = read_key(winreg.HKEY_CLASSES_ROOT, f'CLSID\\{clsid}\\TypeLib')
    version = read_key(winreg.HKEY_CLASSES_ROOT, f'CLSID\\{clsid}\\Version')
    print(f'  TypeLib    : {lib_id or "no type library registered for this class"}')

    # The path the registry points at, per architecture
    registered: list[str] = []
    if lib_id:
        for arch in ('win64', 'win32'):
            path = read_key(
                winreg.HKEY_CLASSES_ROOT,
                f'TypeLib\\{lib_id}\\{version or "1.0"}\\0\\{arch}')
            if path:
                exists = 'exists' if os.path.isfile(path) else 'MISSING FROM DISK'
                print(f'  {arch:<11}: {path}  [{exists}]')
                registered.append(path)
        if not registered:
            print(f'  path       : no path registered under version {version or "1.0"}')

    # Candidates to load the library straight from a file, bypassing the registry
    install = get_install_path(app_id)
    candidates = [*registered]
    if install:
        candidates += [os.path.join(install, n) for n in ('Photoshop.exe', 'Photoshop.tlb')]
    for path in dict.fromkeys(candidates):
        if not os.path.isfile(path):
            continue
        try:
            LoadTypeLibEx(path, REGKIND_NONE)
            print(f'  LOADS OK   : {path}')
        except Exception as e:
            print(f'  cannot load: {path} -- {type(e).__name__}: {e}')


def describe_dispatch(app_id: str) -> None:
    """Reports which dispatch `comtypes` returns for a Photoshop program ID.

    Args:
        app_id: Photoshop COM program ID, e.g. '210'.
    """
    prog_id = f'Photoshop.Application.{app_id}' if app_id else 'Photoshop.Application'
    print(f'\n{prog_id}')
    try:
        app = CreateObject(prog_id, dynamic=True)
    except Exception as e:
        print(f'  Not available: {type(e).__name__}: {e}')
        return

    if isinstance(app, LazyDispatch):
        print('  Dispatch   : lazybind.Dispatch  <-- Photoshop provided a type library')
    elif isinstance(app, FullyDynamicDispatch):
        print('  Dispatch   : dynamic._Dispatch  <-- NO type library, comtypes is guessing')
    else:
        print(f'  Dispatch   : {type(app).__module__}.{type(app).__name__}')

    # Ask the object for its type library directly, which is what comtypes tried
    try:
        app.QueryInterface(IDispatch).GetTypeInfo(0)
        print('  GetTypeInfo: works')
    except Exception as e:
        print(f'  GetTypeInfo: fails -- {type(e).__name__}: {e}')

    for name in ('version', 'name'):
        try:
            print(f'  {name:<11}: {getattr(app, name)}')
        except Exception as e:
            print(f'  {name:<11}: could not be read -- {type(e).__name__}')

    if app_id and isinstance(app, FullyDynamicDispatch):
        report_type_library(app_id)


def main() -> int:
    """Reports on every Photoshop install found.

    Returns:
        Process exit code.
    """
    print(f'Python  : {sys.version.split()[0]}')
    try:
        import comtypes
        print(f'comtypes: {comtypes.__version__}')
    except Exception:
        pass

    app_ids = get_registered_app_ids()
    print(f'\nPhotoshop program ID\'s in the registry: {app_ids or "none found"}')
    for app_id in [*app_ids, '']:
        describe_dispatch(app_id)
    print('\nNote: creating an application object may start Photoshop.')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
