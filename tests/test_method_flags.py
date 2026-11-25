"""
* Tests: Photoshop Method Flags

`comtypes` reads a type library to tell a method from a property. Photoshop doesn't
always provide one, and `comtypes` then falls back to a dispatch which guesses: it
fetches the name as a property first, and concludes the name is a method only if that
fails in one of a few specific ways. Photoshop instead runs the method with no
arguments, so the guess lands anywhere between a COM error and a quiet `None`:

    TypeError: 'NoneType' object is not callable

`comtypes` can be told which names are methods, and `photoshop-python-api` does that,
but each class flags only its own and the lists have holes. `ArtLayer` never flags
`resize`, `rotate` or `translate`, all of which this project calls on layers, so
framing the artwork died on `layer.resize(...)`.

Photoshop can't run here, so a guessing dispatch is modelled and checked against the
flagging helper.
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

# Methods this project calls on Photoshop objects, which must be flagged
REQUIRED = (
    'doJavaScript', 'resize', 'rotate', 'translate', 'move', 'duplicate', 'remove',
    'delete', 'link', 'unlink', 'merge', 'rasterize', 'select', 'invert', 'close',
    'executeAction', 'charIDToTypeID', 'stringIDToTypeID')


"""
* Modelling a guessing dispatch
"""


class Photoshop:
    """The Photoshop object behind a dispatch."""

    def __init__(self):
        self.calls: list[str] = []
        self.name = 'Art Frame Tall'

    def resize(self, *args) -> str:
        self.calls.append('resize')
        return 'resized'


class DynamicDispatch:
    """`comtypes.client.dynamic._Dispatch`, which guesses without a type library."""

    def __init__(self, target: Photoshop):
        self.target, self.methods = target, set()

    def _FlagAsMethod(self, *names: str) -> None:
        self.methods.update(names)

    def __getattr__(self, name: str):
        if name.startswith('_'):
            raise AttributeError(name)
        if not hasattr(self.target, name):
            # An unknown name fails on GetIDsOfNames, before any flag is consulted
            raise AttributeError(name)
        if name in self.methods:
            return getattr(self.target, name)
        value = getattr(self.target, name)
        if not callable(value):
            return value
        # Fetched as a property: Photoshop runs it and hands back nothing
        self.target.calls.append(f'{name} (as a property)')
        return None


class LazyDispatch:
    """`comtypes.client.lazybind.Dispatch`, which reads the type library."""

    def __init__(self, target: Photoshop):
        self.target = target

    def __getattr__(self, name: str):
        return getattr(self.target, name)


"""
* Loading the helper without Photoshop
"""


def load_adobe() -> types.ModuleType:
    """Loads the flagging helper and its name list against a modelled dispatch.

    Notes:
        `src.utils.adobe` imports the whole Windows-only Photoshop stack, so the helper
        and the name list are compiled on their own from the module's source.

    Returns:
        A module holding `flag_photoshop_methods` and `PS_METHOD_NAMES`.
    """
    tree = ast.parse(ADOBE.read_text(encoding='utf-8'), filename=str(ADOBE))
    wanted = [
        n for n in tree.body
        if (isinstance(n, ast.FunctionDef) and n.name == 'flag_photoshop_methods')
        or (isinstance(n, ast.AnnAssign) and getattr(n.target, 'id', None) == 'PS_METHOD_NAMES')]
    assert len(wanted) == 2, 'flag_photoshop_methods or PS_METHOD_NAMES is missing from adobe.py'

    module = types.ModuleType('adobe_method_flags')
    module.__dict__.update({
        'Any': Any, 'FullyDynamicDispatch': DynamicDispatch,
        'suppress': __import__('contextlib').suppress})
    sys.modules['adobe_method_flags'] = module
    exec(compile(ast.Module(body=wanted, type_ignores=[]), str(ADOBE), 'exec'), module.__dict__)
    return module


"""
* Tests
"""


def test_the_model_reproduces_the_original_failure():
    """An unflagged method must come back as None, else these tests prove nothing."""
    dispatch = DynamicDispatch(Photoshop())
    assert dispatch.resize is None
    with pytest.raises(TypeError, match='not callable'):
        dispatch.resize(50, 50)


def test_flagging_makes_a_method_callable():
    """The call which killed the render must go through once the name is flagged.

    Notes:
        This is `frame_layer` scaling the imported artwork, which read `layer.resize`
        as a property, got nothing back, and raised on the call.
    """
    target = Photoshop()
    dispatch = DynamicDispatch(target)
    load_adobe().flag_photoshop_methods(dispatch)
    assert dispatch.resize(50, 50) == 'resized'
    assert target.calls == ['resize']


def test_properties_are_still_properties():
    """Flagging must not turn a property into something which has to be called."""
    dispatch = DynamicDispatch(Photoshop())
    load_adobe().flag_photoshop_methods(dispatch)
    assert dispatch.name == 'Art Frame Tall'


def test_other_dispatch_types_are_left_alone():
    """A dispatch which reads the type library needs no flags, and has no way to take them."""
    dispatch = LazyDispatch(Photoshop())
    assert load_adobe().flag_photoshop_methods(dispatch) is dispatch
    assert dispatch.resize(50, 50) == 'resized'


def test_anything_else_passes_through():
    """Flagging is applied wherever a layer turns up, so it must tolerate non-COM values."""
    flag = load_adobe().flag_photoshop_methods
    for value in (None, 'a layer name', 42):
        assert flag(value) == value


@pytest.mark.parametrize('name', REQUIRED)
def test_methods_this_project_calls_are_flagged(name: str):
    """Every Photoshop method called on an object here must be flagged as one.

    Notes:
        A missing name doesn't fail where it is called. It is read as a property,
        returns None, and raises on the call, or raises a COM error from inside
        Photoshop, neither of which names the method.
    """
    assert name in load_adobe().PS_METHOD_NAMES


def test_flags_are_applied_to_every_photoshop_object():
    """Flagging a helper achieves nothing unless something calls it.

    Notes:
        The initializer is wrapped rather than `_flag_as_method`, so that objects which
        never flag anything of their own are covered too, including the second
        Application object each child object quietly builds for itself.
    """
    tree = ast.parse(ADOBE.read_text(encoding='utf-8'), filename=str(ADOBE))
    assigned = [
        node for node in ast.walk(tree)
        if isinstance(node, ast.Assign)
        and any(isinstance(t, ast.Attribute) and t.attr == '__init__'
                and isinstance(t.value, ast.Name) and t.value.id == 'Photoshop'
                for t in node.targets)]
    assert assigned, (
        'Photoshop.__init__ is never wrapped in adobe.py, so no Photoshop object is '
        'given its method flags.')
