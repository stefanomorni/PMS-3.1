import os
import sys
from pathlib import Path

try:
    import tomllib  # Python 3.11+
except ImportError:
    try:
        import tomLI as tomllib  # Fallback if installed
    except ImportError:
        tomllib = None


def load_config(root_path=None):
    """
    Loads vba-config.toml from the root of the project.
    If root_path is not provided, it searches upwards from the script location.
    """
    if root_path is None:
        # Start from the current working directory
        curr = Path.cwd()
        # Search up to 3 levels for vba-config.toml
        for _ in range(4):
            candidate = curr / "vba-config.toml"
            if candidate.exists():
                root_path = curr
                break
            if curr.parent == curr:  # Root reached
                break
            curr = curr.parent

    if not root_path:
        return {}

    config_path = Path(root_path) / "vba-config.toml"
    if not config_path.exists():
        return {"root": str(root_path)}

    if tomllib:
        try:
            with open(config_path, "rb") as f:
                data = tomllib.load(f)
                config = data.get("general", {})
                config["root"] = str(root_path)
                return config
        except Exception as e:
            print(f"[CONFIG ERROR] Failed to parse {config_path}: {e}")
            return {"root": str(root_path)}
    else:
        # Simple line-based parser fallback if tomllib is missing
        config = {"root": str(root_path)}
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if "=" in line and not line.startswith("["):
                        k, v = line.split("=", 1)
                        config[k.strip()] = v.strip().strip("'").strip('"')
            return config
        except:
            return config


def get_project_name(config):
    """Infers project name from file path or root folder name."""
    if "file" in config:
        return Path(config["file"]).stem
    return Path(config.get("root", os.getcwd())).name


if __name__ == "__main__":
    c = load_config()
    print(f"Detected Config: {c}")
    print(f"Project Name: {get_project_name(c)}")
