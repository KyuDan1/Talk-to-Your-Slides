"""
Talk to Your Slides - Main Entry Point
This is the recommended entry point for the PPT Agent system.

Usage:
    For Web UI (Flask): Use main_flask.py
    For CLI: Use main_cli.py
"""

import sys
import os

def show_usage():
    print("""
╔═══════════════════════════════════════════════════════════════╗
║         Talk to Your Slides - PPT Agent System                ║
╚═══════════════════════════════════════════════════════════════╝

Choose your interface:

1. Web UI (Flask) - Recommended for interactive use
   Run: python pptagent/main_flask.py

2. CLI - For batch processing and automation
   Run: python pptagent/main_cli.py

Setup required:
- Create credentials.yml from credentials.yml.example
- Set API keys in .env file in pptagent/ directory

For more information, see README.md
""")

if __name__ == "__main__":
    show_usage()
