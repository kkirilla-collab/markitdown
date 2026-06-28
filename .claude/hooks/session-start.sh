#!/bin/bash
# SessionStart hook for markitdown — installs the package and dev tools so that
# tests and linters work in Claude Code on the web sessions.
set -euo pipefail

# Only run in the remote (web) environment; locally the user manages their own env.
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

cd "${CLAUDE_PROJECT_DIR:-$(pwd)}"

# Editable install of the main package with all optional converters, plus tooling.
# Idempotent: safe to re-run; container state is cached after the hook completes.
pip install --quiet -e "packages/markitdown[all]"
# cffi provides the _cffi_backend used by cryptography (pulled in via pdfminer);
# the base image ships a Debian cryptography without it, so install it explicitly.
pip install --quiet pytest black cffi

echo "markitdown dev environment ready."
