"""Pytest configuration for examples.

Adds the examples directory to sys.path so that docs_examples_base
can be imported by all example modules.
"""

import sys
from pathlib import Path

# Add examples dir to path so `from docs_examples_base import ...` works
sys.path.insert(0, str(Path(__file__).resolve().parent))
