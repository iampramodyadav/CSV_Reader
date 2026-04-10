"""
Plugin Manager — loads user_plugins.py from a hardcoded path at runtime.
Works in both development and compiled .exe mode.
"""
import importlib.util
import traceback

# ── HARDCODED PLUGIN FILE PATH ────────────────────────────────────────────────
# Change this to wherever your team stores the plugin file.
# PLUGIN_FILE_PATH = r"K:\Users\pramod.kumar\shared\tool\01table_data_editor\user_plugins.py"
from config import PLUGIN_FILE_PATH
# ─────────────────────────────────────────────────────────────────────────────


def load_plugins():
    """
    Dynamically load the plugin file and return its PLUGINS dict.
    Returns empty dict if file not found or has errors.
    """
    try:
        spec = importlib.util.spec_from_file_location("user_plugins", PLUGIN_FILE_PATH)
        if spec is None:
            return {}
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        plugins = getattr(module, "PLUGINS", {})
        return plugins
    except FileNotFoundError:
        return {}
    except Exception:
        traceback.print_exc()
        return {}


def run_plugin(label, df, app):
    """
    Reload the plugin file and run the named function.
    Always reloads fresh — so edits to the .py file take effect immediately.
    Returns modified df or original df if plugin returns None.
    """
    plugins = load_plugins()
    if label not in plugins:
        app.set_status(f"Plugin not found: {label}")
        return df
    
    try:
        fn = plugins[label]
        result = fn(df, app)
        return result if result is not None else df
    except Exception as e:
        traceback.print_exc()
        app.set_status(f"Plugin error: {e}")
        return df