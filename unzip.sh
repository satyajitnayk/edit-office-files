#!/bin/bash

# Check if file argument is provided
if [ $# -eq 0 ]; then
    echo "Usage: $0 <file>"
    exit 1
fi

# Ensure xmllint is installed
command -v xmllint >/dev/null 2>&1 || { echo >&2 "xmllint is required but not installed. Aborting."; exit 1; }


unzip "./input/$1" -d "extracted/$1"
find "extracted/$1" -type f -name "*.xml" -exec xmllint --format {} --output {} \;
find "extracted/$1" -type f -name "*.rels" -exec xmllint --format {} --output {} \;
