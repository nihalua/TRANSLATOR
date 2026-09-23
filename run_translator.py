"""Entry script used to build the Windows executable (see .github/workflows/build-exe.yml)."""

from translator.cli import main

raise SystemExit(main())
