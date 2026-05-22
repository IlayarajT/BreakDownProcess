"""
bump_version.py  —  Called by build.bat to auto-update version.py

Usage:
    python bump_version.py patch            # 1.0.0 → 1.0.1
    python bump_version.py minor            # 1.0.1 → 1.1.0
    python bump_version.py major            # 1.1.0 → 2.0.0
    python bump_version.py set 2.5.0        # force a specific version
    python bump_version.py build-only       # keep version, update build date only
"""

import re
import sys
from datetime import datetime
from pathlib import Path

VERSION_FILE = Path(__file__).parent / "version.py"


def read_version_file():
    return VERSION_FILE.read_text(encoding="utf-8")


def write_version_file(content: str):
    VERSION_FILE.write_text(content, encoding="utf-8")


def get_current_version(content: str) -> str:
    m = re.search(r'__version__\s*=\s*"([\d.]+)"', content)
    if not m:
        raise ValueError("Could not find __version__ in version.py")
    return m.group(1)


def bump(version: str, part: str) -> str:
    major, minor, patch = (int(x) for x in version.split("."))
    if part == "major":
        major += 1; minor = 0; patch = 0
    elif part == "minor":
        minor += 1; patch = 0
    elif part == "patch":
        patch += 1
    else:
        raise ValueError(f"Unknown bump part: {part}")
    return f"{major}.{minor}.{patch}"


def update_version_file(new_version: str, build_date: str, content: str) -> str:
    content = re.sub(
        r'(__version__\s*=\s*)"[\d.]+"',
        f'\\1"{new_version}"',
        content
    )
    content = re.sub(
        r'(__build__\s*=\s*)"[\d]+"',
        f'\\1"{build_date}"',
        content
    )
    return content


def main():
    args = sys.argv[1:]
    if not args:
        print("Usage: python bump_version.py [patch|minor|major|set <ver>|build-only]")
        sys.exit(1)

    content    = read_version_file()
    current    = get_current_version(content)
    build_date = datetime.now().strftime("%Y%m%d")

    command = args[0].lower()

    if command == "build-only":
        new_version = current
        print(f"[version] keeping {current}, updating build → {build_date}")

    elif command == "set":
        if len(args) < 2:
            print("Provide version: python bump_version.py set 2.0.0")
            sys.exit(1)
        new_version = args[1]
        print(f"[version] {current} → {new_version}  (build {build_date})")

    elif command in ("patch", "minor", "major"):
        new_version = bump(current, command)
        print(f"[version] {current} → {new_version}  [{command}]  (build {build_date})")

    else:
        print(f"Unknown command: {command}")
        sys.exit(1)

    updated = update_version_file(new_version, build_date, content)
    write_version_file(updated)
    print(f"[version] version.py updated successfully")

    # Write version to a plain text file for the batch script to read back
    (Path(__file__).parent / ".build_version").write_text(new_version, encoding="utf-8")


if __name__ == "__main__":
    main()
