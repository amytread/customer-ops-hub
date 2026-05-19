#!/bin/bash
set -e

# Regenerate the Customer Pulse dashboard
cd /tmp && python3 generate_web.py

# Sync files into the repo (local folder should be ~/Code/cs-hub)
REPO=~/Code/cs-hub
cp ~/Desktop/tread_customers.html $REPO/customer-pulse/index.html
cp /tmp/generate_web.py           $REPO/generate_web.py
cp /tmp/create_deck_v3.py         $REPO/create_deck_v3.py

# Commit and push if anything changed
cd $REPO
git add customer-pulse/index.html generate_web.py create_deck_v3.py
if git diff --staged --quiet; then
  echo "No changes to push."
else
  git commit -m "Update Customer Pulse $(date '+%Y-%m-%d %H:%M')"
  git push
  echo "Pushed to https://amytread.github.io/cs-hub/customer-pulse/"
fi
