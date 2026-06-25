#!/usr/bin/env bash
# Preview website locally before launch
cd "$(dirname "$0")/../website"
echo "Preview at http://localhost:8080"
python3 -m http.server 8080