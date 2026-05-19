#!/bin/bash
# Watches industry-training/index.html and auto-pushes on every save.
# Run this in your terminal tab: ./watch-industry-training.sh

REPO=~/Code/customer-pulse
FILE="$REPO/industry-training/index.html"

FAVICON_TAG='  <link rel="icon" type="image/svg+xml" href="data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2060%2060%22%3E%3Crect%20width%3D%2260%22%20height%3D%2260%22%20fill%3D%22%230A1820%22%20rx%3D%2212%22/%3E%3Cg%20transform%3D%22translate%281%2C%2015%29%20scale%280.966%29%22%3E%3Cpath%20d%3D%22M38.4037%2015.2638L28.8771%205.73713L19.3504%2015.2638L16.7622%2012.6756L28.8771%200.561646L40.991%2012.6756L38.4037%2015.2638Z%22%20fill%3D%22%23FFAA13%22/%3E%3Cpath%20d%3D%22M38.4037%2021.8914L28.8771%2012.3647L19.3504%2021.8914L16.7622%2019.3032L28.8771%207.18921L40.991%2019.3032L38.4037%2021.8914Z%22%20fill%3D%22%23FFAA13%22/%3E%3Cpath%20d%3D%22M40.991%2025.9307L40.2975%2026.6241L35.7826%2025.8968L28.8771%2018.9913L21.9715%2025.8968L17.4566%2026.6241L16.7622%2025.9307L28.8771%2013.8158L40.991%2025.9307Z%22%20fill%3D%22%23FFE500%22/%3E%3Cpath%20d%3D%22M34.0509%2025.6177L28.8772%2024.7833L23.7026%2025.6177L28.8772%2020.4431L34.0509%2025.6177Z%22%20fill%3D%22%23FFE500%22/%3E%3Cpath%20d%3D%22M16.0143%2013.423L13.4272%2016.01L15.9936%2018.5763L18.5806%2015.9893L16.0143%2013.423Z%22%20fill%3D%22%23FFAA13%22/%3E%3Cpath%20d%3D%22M12.7003%2016.7372L10.1133%2019.3242L15.9931%2025.2041L18.5802%2022.617L12.7003%2016.7372Z%22%20fill%3D%22%23FFAA13%22/%3E%3Cpath%20d%3D%22M16.1692%2026.8316L11.7128%2027.5498L6.80078%2022.6378L9.38898%2020.0505L16.1692%2026.8316Z%22%20fill%3D%22%23FFE500%22/%3E%3Cpath%20d%3D%22M10.4626%2027.752L0%2029.4382L6.07481%2023.3643L10.4626%2027.752Z%22%20fill%3D%22%23FFE500%22/%3E%3Cpath%20d%3D%22M41.7406%2013.4231L39.1738%2015.989L41.7611%2018.5771L44.3278%2016.0113L41.7406%2013.4231Z%22%20fill%3D%22%23FFAA13%22/%3E%3Cpath%20d%3D%22M45.0522%2016.7337L39.1724%2022.6135L41.76%2025.2012L47.6399%2019.3214L45.0522%2016.7337Z%22%20fill%3D%22%23FFAA13%22/%3E%3Cpath%20d%3D%22M50.9548%2022.6387L46.0428%2027.5498L41.5854%2026.8316L48.3665%2020.0505L50.9548%2022.6387Z%22%20fill%3D%22%23FFE500%22/%3E%3Cpath%20d%3D%22M57.7542%2029.4382L47.2925%2027.752L51.6803%2023.3643L57.7542%2029.4382Z%22%20fill%3D%22%23FFE500%22/%3E%3C/g%3E%3C/svg%3E"/>'

ensure_favicon() {
  if ! grep -q 'rel="icon"' "$FILE"; then
    python3 -c "
f=open('$FILE','r');h=f.read();f.close()
h=h.replace('</title>','</title>\n$FAVICON_TAG',1)
f=open('$FILE','w');f.write(h);f.close()
"
  fi
}

echo "Watching industry-training/index.html — save the file to deploy..."
echo "(Ctrl+C to stop)"
echo ""

fswatch -o "$FILE" | while read; do
  echo "Change detected — pushing..."
  ensure_favicon
  cd "$REPO"
  git add industry-training/index.html
  if git diff --staged --quiet; then
    echo "No changes."
  else
    git commit -m "Update Industry Training $(date '+%H:%M:%S')"
    git push
    echo "✓ Live at https://amytread.github.io/customer-ops-hub/industry-training/"
    echo ""
  fi
done
