#!/bin/sh
# Launch the editor server without depending on the current working directory.
# Node resolves a relative script path via getcwd(), which fails with EPERM when
# the shell sits in a macOS TCC-protected folder (~/Documents, ~/Downloads).
# We cd to $HOME and pass an ABSOLUTE script path so getcwd is never needed.
DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$HOME" || exit 1
exec node "$DIR/server.js" "$@"
