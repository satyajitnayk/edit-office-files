#!/bin/bash

# Check if file argument is provided
if [ $# -eq 0 ]; then
    echo "Usage: $0 <file>"
    exit 1
fi

unzip "./input/$1" -d "extracted/$1"
