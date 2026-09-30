"""
* Tests: Photoshop ID Conversion Aliases

Modules alias Photoshop's ID conversion methods at import time:

    sID, cID = APP.stringIDToTypeID, APP.charIDToTypeID

A misspelled alias is not reported by Python or by the linter. `PhotoshopHandler` has no
such attribute, so the lookup falls through to the COM object, where the name was never
flagged as a method and Photoshop evaluates it as a property with no arguments:

    Illegal argument - argument 1
    - Required value is missing

Because these aliases are bound at module scope, that kills the app during import,
before a window appears. `src/helpers/selection.py` and `src/commands/test/utility.py`
both carried `charIDtoTypeID`, with a lowercase 't', which is what this guards.

The spellings are compared against what `PhotoshopHandler` actually defines, read from
the source, so the two stay in step without importing anything Windows-only.
"""
# Standard Library
import ast
from pathlib import Path
import re

# Third Party
import pytest

REPO = Path(__file__).parents[1]
HANDLER = REPO / 'src' / 'utils' / 'adobe.py'
SOURCE_DIRS = ('src', 'plugins')

# Matches any ID conversion method, whatever its casing, e.g. charIDToTypeID
CONVERSION = re.compile(r'(char|string|type)id_?to_?(char|string|type)id', re.IGNORECASE)


def get_defined_conversions() -> set[str]:
    """Collects the ID conversion methods `PhotoshopHandler` defines.

    Returns:
        Set of method names, exactly as spelled in the handler.
    """
    tree = ast.parse(HANDLER.read_text(encoding='utf-8'), filename=str(HANDLER))
    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef) and node.name == 'PhotoshopHandler':
            return {
                n.name for n in node.body
                if isinstance(n, ast.FunctionDef) and CONVERSION.fullmatch(n.name)}
    raise AssertionError(f'PhotoshopHandler is no longer defined in {HANDLER.name}')


def get_source_files() -> list[Path]:
    """Collects every Python module in the project.

    Returns:
        List of paths to Python modules.
    """
    return sorted(p for d in SOURCE_DIRS for p in (REPO / d).rglob('*.py'))


def get_conversion_attributes(tree: ast.AST) -> list[ast.Attribute]:
    """Collects every attribute access which names an ID conversion method.

    Args:
        tree: Parsed module to search.

    Returns:
        List of attribute nodes, whatever object they are accessed on.
    """
    return [
        node for node in ast.walk(tree)
        if isinstance(node, ast.Attribute) and CONVERSION.fullmatch(node.attr)]


def test_handler_defines_the_expected_conversions():
    """The handler must define the conversions, otherwise this test checks nothing."""
    defined = get_defined_conversions()
    assert 'charIDToTypeID' in defined
    assert 'stringIDToTypeID' in defined


@pytest.mark.parametrize('path', get_source_files(), ids=lambda p: str(p.relative_to(REPO)))
def test_conversion_aliases_are_spelled_as_defined(path: Path):
    """Every ID conversion referenced must match a spelling the handler defines.

    Notes:
        A spelling the handler does not define reaches Photoshop through the COM
        object, which raises at import time rather than reporting a missing attribute.
    """
    defined = get_defined_conversions()
    tree = ast.parse(path.read_text(encoding='utf-8'), filename=str(path))
    for node in get_conversion_attributes(tree):
        assert node.attr in defined, (
            f'{path.relative_to(REPO)}:{node.lineno} refers to {node.attr!r}, which '
            f'PhotoshopHandler does not define. Expected one of: {sorted(defined)}.')
