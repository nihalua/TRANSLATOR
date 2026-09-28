"""Entry script used to build CHECKER.exe (see .github/workflows/build-exe.yml)."""

from checker.app import main

raise SystemExit(main())
