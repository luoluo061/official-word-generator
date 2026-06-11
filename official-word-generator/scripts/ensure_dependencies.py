from __future__ import annotations

import importlib.util
import subprocess
import sys


BASE_DEPENDENCIES = {
    "docx": "python-docx",
    "lxml": "lxml",
}

OPTIONAL_DEPENDENCIES = {
    "win32com.client": "pywin32",
}


def _missing(import_names: list[str]) -> list[str]:
    missing_packages: list[str] = []
    package_map = {**BASE_DEPENDENCIES, **OPTIONAL_DEPENDENCIES}
    for import_name in import_names:
        root_name = import_name.split(".", 1)[0]
        if importlib.util.find_spec(root_name) is None:
            missing_packages.append(package_map[import_name])
    return sorted(set(missing_packages))


def ensure(import_names: list[str] | None = None) -> None:
    requested = import_names or list(BASE_DEPENDENCIES)
    packages = _missing(requested)
    if not packages:
        return
    print(f"Installing missing Word generator dependencies: {', '.join(packages)}")
    subprocess.check_call([sys.executable, "-m", "pip", "install", *packages])


if __name__ == "__main__":
    ensure(list(BASE_DEPENDENCIES) + list(OPTIONAL_DEPENDENCIES))
