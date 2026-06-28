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

# ffmpeg is the system binary that pydub needs for the audio-transcription tests;
# install it if it isn't already present (best-effort, don't fail the hook on it).
if ! command -v ffmpeg >/dev/null 2>&1; then
  apt-get update -qq && apt-get install -y -qq ffmpeg || \
    echo "warning: ffmpeg install failed; audio-transcription tests may be skipped/fail"
fi

echo "markitdown dev environment ready."
