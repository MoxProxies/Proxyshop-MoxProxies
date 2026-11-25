"""
* Tests: JavaScript Execution

`Document.activeLayer` asks Photoshop for the active layer's type by running a line of
JavaScript, so every lookup of the active layer goes through `eval_javascript`.

Photoshop doesn't implement `IDispatch::GetTypeInfo`. Where `comtypes` can't get a type
library it falls back to its fully dynamic dispatch, which can't tell a method from a
property and so fetches `doJavaScript` as a property. Photoshop runs it with no
arguments and reports the failure as a plain 'Exception occurred', which `comtypes`
doesn't recognise as 'that name is a method', so the whole render dies with:

    Illegal argument - argument 1
    - Required value is missing

`comtypes` can be told a name is a method, and `photoshop-python-api` does that, but
only for the object it was handed. A Document builds a second Application object of its
own and calls JavaScript through that one, which nothing ever flags.

Photoshop can't run here, so both dispatch types are modelled: one resolves a name
whatever it is, the way `lazybind` does with a type library, one refuses an unflagged
name the way the dynamic dispatch does. `eval_javascript` must work against both.
"""
# Standard Library
import ast
from pathlib import Path
import sys
import types
from typing import Any

# Third Party
import pytest

REPO = Path(__file__).parents[1]
ADOBE = REPO / 'src' / 'utils' / 'adobe.py'

# The failure Photoshop reports when a method is evaluated as a property
COM_ERROR = (
    -2147352567, 'Exception occurred.',
    ('Illegal argument - argument 1\n- Required value is missing',
     'Adobe Photoshop', None, 0, -2147220262))


"""
* Modelling Photoshop's COM objects
"""


class COMError(Exception):
    """`_ctypes.COMError`, which `comtypes` raises straight through."""


class Photoshop:
    """The Photoshop application behind a dispatch, recording what it was asked."""

    def __init__(self):
        self.calls: list[tuple] = []

    def doJavaScript(self, *args) -> str:
        if not args:
            # Photoshop evaluates a property fetch as a call with no arguments
            raise COMError(*COM_ERROR)
        self.calls.append(args)
        return 'ArtLayer'


class LazyDispatch:
    """`comtypes.client.lazybind.Dispatch`, which has a type library to read."""

    def __init__(self, target: Photoshop):
        self.target = target

    def __getattr__(self, name: str):
        return getattr(self.target, name)


class DynamicDispatch:
    """`comtypes.client.dynamic._Dispatch`, which has no type library to read."""

    def __init__(self, target: Photoshop):
        self.target, self.methods = target, set()

    def _FlagAsMethod(self, *names: str) -> None:
        self.methods.update(names)

    def __getattr__(self, name: str):
        if name in self.methods:
            return getattr(self.target, name)
        # Fetched as a property, which Photoshop runs with no arguments. The error
        # isn't one comtypes treats as 'this is really a method', so it is raised.
        return getattr(self.target, name)()


class PhotoshopObject:
    """A `photoshop-python-api` object, e.g. an Application or a Document.

    Args:
        dispatch: Dispatch type the object's COM objects use.
        has_parent: True for a child object such as a Document, which calls JavaScript
            through an Application object of its own rather than through itself.
    """

    def __init__(self, dispatch: type, has_parent: bool = True):
        self.target = Photoshop()
        self.app = dispatch(self.target if not has_parent else Photoshop())
        self.adobe = dispatch(self.target) if has_parent else None
        self._has_parent = has_parent


"""
* Loading the implementation without Photoshop
"""


def load_eval_javascript() -> Any:
    """Loads the `eval_javascript` replacement against the modelled dispatch types.

    Notes:
        `src.utils.adobe` imports the whole Windows-only Photoshop stack, so the
        function is compiled on its own from the module's source.

        `FullyDynamicDispatch` always names the dynamic dispatch, as it does in the
        module, so a `lazybind` dispatch is left alone rather than flagged.

    Returns:
        The `eval_javascript` implementation.
    """
    tree = ast.parse(ADOBE.read_text(encoding='utf-8'), filename=str(ADOBE))
    found = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == '_eval_javascript']
    assert found, '_eval_javascript is no longer defined in adobe.py'

    module = types.ModuleType('adobe_eval_javascript')
    module.__dict__.update({
        'Any': Any, 'Photoshop': PhotoshopObject,
        'FullyDynamicDispatch': DynamicDispatch})
    sys.modules['adobe_eval_javascript'] = module
    exec(compile(ast.Module(body=found, type_ignores=[]), str(ADOBE), 'exec'), module.__dict__)
    return module._eval_javascript


"""
* Tests
"""


def test_the_model_reproduces_the_original_failure():
    """An unflagged name on the dynamic dispatch must fail, else these tests prove nothing."""
    dispatch = DynamicDispatch(Photoshop())
    with pytest.raises(COMError) as err:
        _ = dispatch.doJavaScript
    assert 'Required value is missing' in err.value.args[2][0]


@pytest.mark.parametrize('dispatch', [LazyDispatch, DynamicDispatch],
                         ids=['lazybind', 'dynamic-dispatch'])
def test_javascript_runs_on_either_dispatch(dispatch: type):
    """A Document must be able to run JavaScript however its dispatch behaves.

    Notes:
        This is the lookup which failed in the field: `Document.activeLayer` asking
        Photoshop for the active layer's type while importing the card's artwork.
    """
    obj = PhotoshopObject(dispatch, has_parent=True)
    result = load_eval_javascript()(obj, 'app.activeDocument.activeLayer.typename')
    assert result == 'ArtLayer'
    assert obj.target.calls == [('app.activeDocument.activeLayer.typename',)]


@pytest.mark.parametrize('dispatch', [LazyDispatch, DynamicDispatch],
                         ids=['lazybind', 'dynamic-dispatch'])
def test_an_application_runs_javascript_through_itself(dispatch: type):
    """An object with no parent, i.e. the Application, calls JavaScript on itself."""
    obj = PhotoshopObject(dispatch, has_parent=False)
    result = load_eval_javascript()(obj, 'app.name')
    assert result == 'ArtLayer'
    assert obj.target.calls == [('app.name',)]


@pytest.mark.parametrize('dispatch', [LazyDispatch, DynamicDispatch],
                         ids=['lazybind', 'dynamic-dispatch'])
def test_optional_arguments_are_omitted_not_passed_as_none(dispatch: type):
    """Arguments which weren't given must not reach Photoshop.

    Notes:
        `comtypes` sends `None` as a COM null, which is a value Photoshop is being
        handed, not an argument it was never given.
    """
    obj = PhotoshopObject(dispatch, has_parent=True)
    load_eval_javascript()(obj, 'app.name', None, None)
    assert obj.target.calls == [('app.name',)]


@pytest.mark.parametrize('dispatch', [LazyDispatch, DynamicDispatch],
                         ids=['lazybind', 'dynamic-dispatch'])
def test_given_arguments_are_passed_through(dispatch: type):
    """Arguments which were given must still reach Photoshop, in order."""
    obj = PhotoshopObject(dispatch, has_parent=True)
    load_eval_javascript()(obj, 'app.name', ['arg'], 2)
    assert obj.target.calls == [('app.name', ['arg'], 2)]


def test_the_replacement_is_installed_on_the_base_class():
    """Defining the function achieves nothing unless it replaces the original.

    Notes:
        Every Photoshop object inherits `eval_javascript` from the one base class, so
        the assignment is what puts the fix in front of `Document.activeLayer`.
    """
    tree = ast.parse(ADOBE.read_text(encoding='utf-8'), filename=str(ADOBE))
    assigned = [
        node for node in ast.walk(tree)
        if isinstance(node, ast.Assign)
        and any(isinstance(t, ast.Attribute) and t.attr == 'eval_javascript'
                and isinstance(t.value, ast.Name) and t.value.id == 'Photoshop'
                for t in node.targets)]
    assert assigned, (
        'Photoshop.eval_javascript is never reassigned in adobe.py, so the '
        'photoshop-python-api implementation is still the one being used.')
