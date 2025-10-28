"""
Dynamic loader for Word formatting recipes.
Each recipe lives in formatter/recipes_word/<name>.py
"""

import importlib
import os
import pkgutil

__all__ = ["get_word_recipe", "discover_word_recipes"]


def discover_word_recipes():
    """
    Scan the formatter/recipes_word/ directory for available recipe modules.
    Automatically imports any module that defines a callable ending in `_py`.
    Returns a dict of {recipe_name: callable}.
    """
    from myapp.processor import recipes_word as this_pkg

    recipes = {}
    package_path = os.path.dirname(this_pkg.__file__)

    for _, module_name, _ in pkgutil.iter_modules([package_path]):
        # Skip internal files or packages
        if module_name.startswith("_") or module_name in {"__init__"}:
            continue

        try:
            module = importlib.import_module(f"myapp.processor.recipes_word.{module_name}")
            # Find the first callable ending in "_py"
            func = next(
                getattr(module, attr)
                for attr in dir(module)
                if callable(getattr(module, attr)) and attr.endswith("_py")
            )
            recipes[module_name] = func
        except Exception as e:
            print(f"[WARN] Failed to load recipe '{module_name}': {e}")

    return recipes


def get_word_recipe(name: str):
    """Dynamically import and return a Word recipe function by name."""
    module_path = f"myapp.processor.recipes_word.{name}"
    try:
        module = importlib.import_module(module_path)
        func = next(
            getattr(module, attr)
            for attr in dir(module)
            if callable(getattr(module, attr)) and attr.endswith("_py")
        )
        return func
    except (ImportError, StopIteration) as e:
        raise ImportError(f"Recipe '{name}' not found or invalid: {e}")
