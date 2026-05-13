#!/bin/bash
set -e

# Regenerate the dashboard from /tmp/
cd /tmp && python3 generate_web.py

# Sync all three files into the repo
cp ~/Desktop/tread_customers.html ~/Code/customer-pulse/index.html
cp /tmp/generate_web.py      ~/Code/customer-pulse/generate_web.py
cp /tmp/create_deck_v3.py    ~/Code/customer-pulse/create_deck_v3.py

# Commit and push if anything changed
cd ~/Code/customer-pulse
git add index.html generate_web.py create_deck_v3.py
if git diff --staged --quiet; then
  echo "No changes to push."
else
  git commit -m "Update Customer Pulse $(date '+%Y-%m-%d %H:%M')"
  git push
  echo "Pushed to https://amytread.github.io/customer-pulse/"
fi
