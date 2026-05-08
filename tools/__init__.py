"""
MCP Tools Package

Tool modules are automatically discovered and loaded by the document server.
Each module should define a class ending in 'Tools' that provides tool_ methods.
"""

import logging
from importlib import import_module
from inspect import getmembers, isclass
from pathlib import Path
from pkgutil import iter_modules

_logger = logging.getLogger(__name__)


DEFAULT_TOOL_MODULES = [
    "azure_pricing_tools",
    "discovery_tools",
    "excel_advanced_tools",
    "excel_tools",
    "office_unified_tools",
    "pptx_advanced_tools",
    "pptx_slide_transfer_tools",
    "pptx_tools",
    "web_tools",
    "word_advanced_tools",
    "word_tools",
]


def _discover_tool_module_names() -> list[str]:
    """Discover tool modules in normal and frozen execution modes."""
    module_names: list[str] = []

    tools_dir = Path(__file__).parent
    if tools_dir.exists():
        module_names.extend(
            sorted(
                module_path.stem
                for module_path in tools_dir.glob("*_tools.py")
            )
        )

    if not module_names:
        module_names.extend(
            sorted(
                name
                for _, name, _ in iter_modules(__path__)
                if name.endswith("_tools")
            )
        )

    if not module_names:
        module_names.extend(DEFAULT_TOOL_MODULES)

    deduped: list[str] = []
    seen = set()
    for name in module_names:
        if name not in seen:
            seen.add(name)
            deduped.append(name)
    return deduped

# Discover all tool classes in this package
TOOL_CLASSES = []
_seen_classes = set()

for _module_name in _discover_tool_module_names():
    try:
        _module = import_module(f".{_module_name}", package=__name__)
        for _name, _cls in getmembers(_module, isclass):
            # Only include classes that:
            # 1. End with 'Tools' (but not just 'Tools')
            # 2. Are defined in this module (not imported)
            # 3. Haven't been seen before
            if (_name.endswith("Tools")
                and _name != "Tools"
                and _cls.__module__ == _module.__name__
                and _name not in _seen_classes):
                TOOL_CLASSES.append(_cls)
                _seen_classes.add(_name)
    except Exception as exc:  # noqa: BLE001
        # Log and skip — on Windows, missing native libraries (lxml, etc.)
        # may raise ImportError, OSError, or DLL-load failures.
        _logger.warning("Failed to load tool module %s: %s: %s",
                        _module_name, type(exc).__name__, exc)

_logger.info("Loaded %d tool classes: %s", len(TOOL_CLASSES),
             [c.__name__ for c in TOOL_CLASSES])

__all__ = [
    "TOOL_CLASSES",
]
