"""Allow ``python -m gsa_fcp_toolkit`` to run the CLI."""

import sys

from gsa_fcp_toolkit.cli import main

if __name__ == "__main__":
    sys.exit(main())
