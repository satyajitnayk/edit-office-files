#!/bin/bash

# Check if file argument is provided
if [ $# -eq 0 ]; then
    echo "Usage: $0 <file>"
    exit 1
fi

cd extracted/$1
zip -r '../../compressed/'$1 *
