"""
* Tests: Type Casts

`typing.cast(typ, val)` returns `val` untouched, so swapping its arguments silently
returns the type annotation instead of the value. Nothing raises until the annotation
is used as if it were the value, often deep inside a render:

    AttributeError: '_SpecialForm' object has no attribute 'visible'

Neither the interpreter nor a linter catches it, so the source is checked directly.
"""
# Standard Library
import ast
from pathlib import Path

# Third Party
import pytest

REPO = Path(__file__).parents[1]
SOURCE_DIRS = ('src', 'plugins')


def get_source_files() -> list[Path]:
    """Collects every Python module in the project.

    Returns:
        List of paths to Python modules.
    """
    return sorted(p for d in SOURCE_DIRS for p in (REPO / d).rglob('*.py'))


def get_cast_calls(tree: ast.AST) -> list[ast.Call]:
    """Collects every call to `cast` in a parsed module.

    Args:
        tree: Parsed module to search.

    Returns:
        List of `cast` call nodes.
    """
    return [
        node for node in ast.walk(tree)
        if isinstance(node, ast.Call) and (
            (isinstance(node.func, ast.Name) and node.func.id == 'cast') or
            (isinstance(node.func, ast.Attribute) and node.func.attr == 'cast'))]


@pytest.mark.parametrize('path', get_source_files(), ids=lambda p: str(p.relative_to(REPO)))
def test_cast_arguments_are_not_swapped(path: Path):
    """The first argument given to `cast` must be a type expression, never a value.

    Notes:
        A type expression never contains a call, whereas the values passed to `cast`
        in this project reliably do, e.g. `super().background_group`.
    """
    tree = ast.parse(path.read_text(encoding='utf-8'), filename=str(path))
    for node in get_cast_calls(tree):
        if not node.args:
            continue
        calls = [n for n in ast.walk(node.args[0]) if isinstance(n, ast.Call)]
        assert not calls, (
            f'{path.relative_to(REPO)}:{node.lineno} passes a value as the first argument '
            f'to cast(). The type comes first: cast(TheType, the_value).')
