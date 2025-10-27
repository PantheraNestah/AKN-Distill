"""
formatter/recipes_word.py
Dynamic loader for Word formatting recipes.
Each recipe lives in formatter/recipes_word/<name>.py
"""

import importlib


def get_word_recipe(name: str):
    """Dynamically import a Word recipe by name."""
    module_path = f"formatter.recipes_word.{name}"
    try:
        module = importlib.import_module(module_path)
        # Find first callable ending in _py
        func = next(
            getattr(module, attr)
            for attr in dir(module)
            if callable(getattr(module, attr)) and attr.endswith("_py")
        )
        return func
    except (ImportError, StopIteration) as e:
        raise ImportError(f"Recipe '{name}' not found or invalid: {e}")
