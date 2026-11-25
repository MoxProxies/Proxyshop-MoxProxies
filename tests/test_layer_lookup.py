"""
* Tests: Layer Name Lookups

Indexing a Photoshop collection by name, e.g. `doc.layerSets['Text and Icons']`, reaches
`comtypes`' fully dynamic dispatch. Its `__getitem__` accepts only an integer index, so a
string raises, in every released version of comtypes:

    TypeError: '>' not supported between instances of 'str' and 'int'

Which dispatch Photoshop returns varies by version and by the state of the typelib cache,
so the same lookup works on one machine and fails on another. On a failure the group
resolved to the whole document, every nested lookup then missed, and the render died far
from the cause with `'NoneType' object has no attribute 'textItem'`.

Photoshop can't run here, so the collections are modelled: one indexes by name the way
`lazybind` does, one raises on a string the way the dynamic dispatch does. The lookup
helpers must give the same answer against both.
"""
# Standard Library
import ast
from pathlib import Path
import sys
import types
from typing import Any, Iterable, Optional, Union

# Third Party
import pytest

REPO = Path(__file__).parents[1]
LAYERS = REPO / 'src' / 'helpers' / 'layers.py'


"""
* Modelling Photoshop's collections
"""


class Item:
    """A layer or group, holding nested collections."""

    def __init__(self, name: str, layer_sets=None, art_layers=None, strict: bool = False):
        self.name = name
        self.layerSets = Collection(layer_sets or [], strict=strict)
        self.artLayers = Collection(art_layers or [], strict=strict)


class Collection:
    """A Photoshop collection.

    Args:
        items: Items held by the collection.
        strict: True models `comtypes`' dynamic dispatch, which raises TypeError when
            indexed by name. False models `lazybind`, which resolves the name.
    """

    def __init__(self, items: list, strict: bool = False):
        self._items, self._strict = items, strict

    def __iter__(self):
        return iter(self._items)

    def __getitem__(self, key):
        if self._strict and isinstance(key, str):
            # The exact failure comtypes produces
            raise TypeError("'>' not supported between instances of 'str' and 'int'")
        for item in self._items:
            if item.name == key:
                return item
        raise KeyError(key)


def build_document(strict: bool) -> Item:
    """Builds a document shaped like `borderless-vector.psd`.

    Args:
        strict: Whether its collections refuse to be indexed by name.

    Returns:
        The document.
    """
    tall = Item('Tall', art_layers=[
        Item('Rules Text - Noncreature', strict=strict),
        Item('Rules Text - Creature', strict=strict)], strict=strict)
    text_and_icons = Item('Text and Icons', layer_sets=[tall], strict=strict)
    return Item('borderless-vector.psd', layer_sets=[text_and_icons], strict=strict)


"""
* Loading the helpers without Photoshop
"""


def load_helpers(document: Item) -> types.ModuleType:
    """Loads the lookup helpers against a stubbed Photoshop application.

    Notes:
        `src.helpers.layers` imports the whole application stack, which is Windows only,
        so the helper functions are compiled on their own from the module's source.

    Args:
        document: Document the stubbed application reports as active.

    Returns:
        A module holding the lookup helpers.
    """
    tree = ast.parse(LAYERS.read_text(encoding='utf-8'), filename=str(LAYERS))
    wanted = {'find_in_collection', 'resolve_container', 'describe_group'}
    found = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in wanted]
    missing = wanted - {n.name for n in found}
    assert not missing, f'helpers no longer defined in layers.py: {sorted(missing)}'

    module = types.ModuleType('layers_helpers')
    module.__dict__.update({
        'Any': Any, 'Optional': Optional, 'Union': Union, 'Iterable': Iterable,
        'LayerContainerTypes': Any,
        'suppress': __import__('contextlib').suppress,
        'APP': types.SimpleNamespace(activeDocument=document),
        'LayerContainer': Item,
    })
    sys.modules['layers_helpers'] = module
    exec(compile(ast.Module(body=found, type_ignores=[]), str(LAYERS), 'exec'), module.__dict__)
    return module


"""
* Tests
"""


@pytest.mark.parametrize('strict', [False, True], ids=['lazybind', 'dynamic-dispatch'])
def test_find_in_collection_resolves_by_name(strict: bool):
    """A name resolves whether or not the collection can be indexed by one."""
    doc = build_document(strict)
    helpers = load_helpers(doc)
    found = helpers.find_in_collection(doc.layerSets, 'Text and Icons')
    assert found is not None, 'group was not found'
    assert found.name == 'Text and Icons'


@pytest.mark.parametrize('strict', [False, True], ids=['lazybind', 'dynamic-dispatch'])
def test_find_in_collection_missing_name_returns_none(strict: bool):
    """A name which isn't present yields None rather than raising."""
    doc = build_document(strict)
    helpers = load_helpers(doc)
    assert helpers.find_in_collection(doc.layerSets, 'Not A Group') is None


@pytest.mark.parametrize('strict', [False, True], ids=['lazybind', 'dynamic-dispatch'])
def test_resolve_container_walks_a_nested_path(strict: bool):
    """The group path which failed in the field resolves on both dispatch types.

    Notes:
        This is the exact lookup from `BorderlessVectorTemplate.size`: the 'Tall' group
        inside 'Text and Icons'. Indexing by name sent it to the document root, where
        'Tall' does not exist.
    """
    doc = build_document(strict)
    helpers = load_helpers(doc)
    text_group = helpers.find_in_collection(doc.layerSets, 'Text and Icons')
    container = helpers.resolve_container([text_group, 'Tall'])
    assert container is not None, "'Tall' was not reached"
    assert container.name == 'Tall'

    layer = helpers.find_in_collection(container.artLayers, 'Rules Text - Noncreature')
    assert layer is not None, 'rules layer was not found'


def test_the_model_reproduces_the_original_failure():
    """The strict collection must reject a name, else these tests prove nothing."""
    doc = build_document(strict=True)
    with pytest.raises(TypeError, match="not supported between instances of 'str' and 'int'"):
        _ = doc.layerSets['Text and Icons']


@pytest.mark.parametrize('group, expected', [
    (None, 'active document'),
    ('Text and Icons', 'Text and Icons'),
    (['Text and Icons', 'Tall'], 'Text and Icons -> Tall')])
def test_describe_group_reads_clearly(group, expected):
    """A failed lookup names the group path it searched."""
    helpers = load_helpers(build_document(strict=False))
    assert helpers.describe_group(group) == expected
