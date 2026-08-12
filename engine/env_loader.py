"""Read a .env file into the process environment.

main.py calls load_env_file() before anything reads os.environ — the database
path, the Google OAuth credentials and the rest — so that a developer can keep
those in a file instead of exporting them by hand.

Two rules matter more than the parsing:

  Variables already in the environment win. Railway sets its own on the
  container, and a .env that ever found its way into an image must not be able
  to overwrite production's database path or credentials with a developer's.

  A missing file is not an error. Production has no .env at all, so the loader
  returning quietly is the normal case, not a degraded one.

Deliberately not python-dotenv: this is a dozen lines, it is called once at
import time, and it keeps the deployment free of a dependency it would only
use here.
"""
from __future__ import annotations

import os


def _unquote(value: str) -> str:
    """Strip one matching pair of surrounding quotes, if present."""
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
        return value[1:-1]
    return value


def load_env_file(path: str | None = None, override: bool = False) -> int:
    """Load KEY=VALUE lines from `path` (default: .env beside the project root)
    into os.environ. Returns how many variables were set.

    Blank lines and lines starting with # are skipped, an optional `export `
    prefix is allowed, and the value is everything after the first `=` so that
    connection strings and base64 secrets survive intact.
    """
    if path is None:
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(root, ".env")

    try:
        # utf-8-sig so a file saved from Notepad with a BOM does not turn the
        # first key into "﻿KEY".
        with open(path, "r", encoding="utf-8-sig") as fh:
            lines = fh.readlines()
    except FileNotFoundError:
        return 0
    except OSError:
        return 0

    loaded = 0
    for raw in lines:
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        if line.startswith("export "):
            line = line[len("export "):].lstrip()
        key, _, value = line.partition("=")
        key = key.strip()
        if not key:
            continue
        if not override and key in os.environ:
            continue                    # the platform's value stands
        os.environ[key] = _unquote(value.strip())
        loaded += 1
    return loaded
