# main.py
import os
import sys
import argparse

# Import the main app class
from .ui import PQManagerUI

sys.path.append(os.path.dirname(__file__))


def run():
    """
    This is the main entry point function that will be called
    by the command-line script.
    """

    print(f"🚀 Shan's PQ Magic")
    print(f"Trying to Start UI")

    parser = argparse.ArgumentParser(
        description="Shan's PQ Magic ✨ - A UI for managing Power Query files."
    )

    parser.add_argument(
        "root_path",
        default="NOT_PROVIDED",
        help="The root directory of your Power Query repository. "
             "Defaults to the current directory if not provided."
    )

    parser.add_argument(
        "excel_hwnd",
        nargs="?",  # Makes the argument optional
        default=None,
        help="Excel Instance in which the UI should be launched"
    )

    args = parser.parse_args()

    # Use the resolved path
    if args.root_path == "NOT_PROVIDED":
        print(f"Give a Correct Path to your Power Query Repository")
        sys.exit(1)

    hwnd = int(args.excel_hwnd) if args.excel_hwnd is not None else None
    repo_path = os.path.abspath(args.root_path)

    if not os.path.exists(repo_path):
        print(f"Error: Path not found: {repo_path}")
        sys.exit(1)

    print(f"Started UI in {repo_path}")

    # Run the app
    PQManagerUI(repo_path, hwnd=hwnd)


if __name__ == "__main__":
    run()
