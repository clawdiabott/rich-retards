#!/usr/bin/env bash
# Usage: ./scripts/update-ca.sh <CONTRACT_ADDRESS> <PUMP_FUN_URL>
set -euo pipefail

CA="${1:?Usage: $0 <CA> <pump_url>}"
PUMP="${2:?Usage: $0 <CA> <pump_url>}"

WEBSITE="/home/patrick/rich-retards/website/index.html"

sed -i "s|LAUNCHING TONIGHT — CA POSTED AT DEPLOY|${CA}|g" "$WEBSITE"
sed -i "s|https://pump.fun\"|${PUMP}\"|g" "$WEBSITE"
sed -i "s|// document.getElementById('ca').textContent = 'YOUR_CA_HERE';|document.getElementById('ca').textContent = '${CA}';|g" "$WEBSITE"
sed -i "s|// document.getElementById('buy-btn').href = 'https://pump.fun/YOUR_CA_HERE';|document.getElementById('buy-btn').href = '${PUMP}';|g" "$WEBSITE"

echo "Updated website with CA: $CA"
echo "Buy link: $PUMP"