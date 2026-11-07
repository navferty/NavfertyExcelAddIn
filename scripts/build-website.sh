#!/usr/bin/env bash
set -euo pipefail

if [[ $# -ne 1 || -z "$1" ]]; then
    echo "Usage: $0 <destination-directory>" >&2
    exit 2
fi

destination=$1
hugo --source Website --destination "$destination" --gc --minify
python3 scripts/verify-hugo.py "$destination"
python3 scripts/test-hugo-output.py "$destination"
node scripts/test-hugo-search.js
