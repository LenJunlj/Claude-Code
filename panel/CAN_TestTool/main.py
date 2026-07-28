#!/usr/bin/env python
"""
CAN Test Tool - FSCM/RSCM
Main entry point.

Usage:
    py main.py              # Launch GUI
    py main.py --virtual    # Launch with virtual CAN
"""

import sys
import os

# Ensure the script directory is in the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from can_test_tool import load_dbc, CanTestTool


def main():
    print('Loading DBC file...')
    db = load_dbc()
    print(f'DBC loaded: {len(db.messages)} messages, {len(db.nodes)} nodes')

    app = CanTestTool(db)

    # Auto-connect in virtual mode if requested
    if '--virtual' in sys.argv:
        app.after(500, lambda: (
            app.channel_var.set('virtual'),
            app._connect_virtual()
        ))

    app.mainloop()


if __name__ == '__main__':
    main()
