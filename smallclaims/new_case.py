#!/usr/bin/env python3
"""
Shortcut: generate a blank case template.

Usage:
    python new_case.py                          # writes cases/new_case.json
    python new_case.py cases/client_name.json   # writes to specified path
"""
import sys
from fill_forms import generate_template

output = sys.argv[1] if len(sys.argv) > 1 else "cases/new_case.json"
generate_template(output)
