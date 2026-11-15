"""Load YAML and refuse a mapping that repeats a key.

PyYAML keeps the last of two equal keys without a word, so a second
`permissions:` or `timeout-minutes:` in a workflow would replace the first
unnoticed. This loader raises instead. It is the safe loader otherwise: no
Python objects are constructed.

The key `on` of a workflow is read by YAML 1.1 as the boolean true; callers
ask for it with `events(document)`.
"""
from __future__ import annotations

from typing import Any

import yaml


class DuplicateKeyError(ValueError):
    pass


class _Loader(yaml.SafeLoader):
    pass


def _mapping(loader: yaml.SafeLoader, node: yaml.MappingNode, deep: bool = False) -> dict[Any, Any]:
    seen: set[Any] = set()
    for key_node, _ in node.value:
        key = loader.construct_object(key_node, deep=deep)
        if key in seen:
            mark = key_node.start_mark
            raise DuplicateKeyError(f"duplicate key {key!r} at line {mark.line + 1}, column {mark.column + 1}")
        seen.add(key)
    return loader.construct_mapping(node, deep=deep)


_Loader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, _mapping)


def loads(text: str) -> Any:
    return yaml.load(text, Loader=_Loader)  # noqa: S506 - a SafeLoader subclass


def events(document: dict[Any, Any]) -> Any:
    """The `on:` block of a workflow, whichever way YAML read its key."""
    return document.get("on", document.get(True))
