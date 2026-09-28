"""Entry point: python donations.py run --year 2025"""
import sys

from donation_tracker.cli import main

if __name__ == "__main__":
    sys.exit(main())
