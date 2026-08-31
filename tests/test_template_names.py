"""
* Tests: Template Names

A template's displayed name comes from the first entry of `AppTemplate.all_names`.
Building that list from a set makes the name depend on Python's per-process string
hashing, so a template can be listed under a different one of its names on every
launch — `normal.psd` showing as 'Fullart' one run and 'Stargazing' the next.

A single run of such code always looks correct, so the ordering is checked directly
against the manifest, and the collection idiom is checked in the source.
"""
# Standard Library
import ast
from pathlib import Path

# Third Party
import pytest
import yaml

REPO = Path(__file__).parents[1]
MANIFEST = REPO / 'src' / 'data' / 'manifest.yml'
LOADER = REPO / 'src' / '_loader.py'

# Properties whose order reaches the user, so they must not be built from a set
ORDERED_PROPERTIES = ('all_names', 'types_supported', 'all_classes')


def get_manifest() -> dict:
    """Loads the built-in template manifest.

    Returns:
        Parsed manifest data.
    """
    return yaml.safe_load(MANIFEST.read_text(encoding='utf-8'))


def get_property(name: str) -> ast.FunctionDef:
    """Finds a property definition in the loader module.

    Args:
        name: Name of the property to find.

    Returns:
        Parsed function definition.
    """
    tree = ast.parse(LOADER.read_text(encoding='utf-8'))
    for node in ast.walk(tree):
        if isinstance(node, ast.FunctionDef) and node.name == name:
            return node
    raise AssertionError(f'{name} is no longer defined in {LOADER.name}')


@pytest.mark.parametrize('name', ORDERED_PROPERTIES)
def test_ordered_properties_are_not_built_from_a_set(name: str):
    """These lists must preserve manifest order, so no set may appear in them.

    Notes:
        A set comprehension and a `set()` call both scramble the order under Python's
        randomized string hashing, and neither fails on any single run.
    """
    node = get_property(name)
    for child in ast.walk(node):
        assert not isinstance(child, ast.SetComp), (
            f'{name} builds its list from a set comprehension, which makes the order, '
            f'and the template name taken from it, vary between launches.')
        assert not (isinstance(child, ast.Call) and isinstance(child.func, ast.Name)
                    and child.func.id == 'set'), (
            f'{name} builds its list from set(), which makes the order, and the '
            f'template name taken from it, vary between launches.')


def test_template_named_normal_is_listed_first():
    """`normal.psd` must be named after its 'Normal' entry.

    Notes:
        `generate_template_name` names a template after `all_names[0]`, so the manifest
        order decides what the download manager calls it. This is the template users
        look for by name.
    """
    names = list(get_manifest()['normal.psd']['templates'].keys())
    assert names[0] == 'Normal', (
        f"normal.psd would be listed as '{names[0]}' rather than 'Normal'.")


def test_manifest_preserves_the_order_names_are_written_in():
    """Parsing the manifest must keep each template's names in written order.

    Notes:
        The fix relies on manifest order deciding the displayed name. A loader which
        sorted keys would reorder the names just as silently as the set did.
    """
    text = MANIFEST.read_text(encoding='utf-8')
    multi = {
        file_name: list(data['templates'].keys())
        for file_name, data in get_manifest().items()
        if len(data.get('templates', {})) > 1}

    # Guards the test itself: these are the templates the ordering actually affects
    assert multi, 'no multi-name templates found, this test is checking nothing'
    for file_name, names in multi.items():
        written = sorted(names, key=lambda n: text.index(f'{n}:', text.index(f'{file_name}:')))
        assert names == written, (
            f'{file_name} parsed its names as {names}, but they are written as {written}.')
