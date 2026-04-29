#!/bin/bash
# Instala un job diario que descarga la cartola a las 9:00 AM
# USO: bash programar_descarga_diaria.sh

PLIST_ID="cl.nativaelements.cartola"
PLIST_PATH="$HOME/Library/LaunchAgents/${PLIST_ID}.plist"
PYTHON=$(which python3)
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

cat > "$PLIST_PATH" << EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>${PLIST_ID}</string>
    <key>ProgramArguments</key>
    <array>
        <string>${PYTHON}</string>
        <string>${SCRIPT_DIR}/automatizacion_diaria.py</string>
    </array>
    <key>WorkingDirectory</key>
    <string>${SCRIPT_DIR}</string>
    <key>StartCalendarInterval</key>
    <dict>
        <key>Hour</key>
        <integer>9</integer>
        <key>Minute</key>
        <integer>0</integer>
    </dict>
    <key>StandardOutPath</key>
    <string>${SCRIPT_DIR}/logs/cartola.log</string>
    <key>StandardErrorPath</key>
    <string>${SCRIPT_DIR}/logs/cartola_error.log</string>
    <key>RunAtLoad</key>
    <false/>
</dict>
</plist>
EOF

mkdir -p "${SCRIPT_DIR}/logs"
launchctl unload "$PLIST_PATH" 2>/dev/null
launchctl load "$PLIST_PATH"

echo "✓ Job diario instalado: descarga cartola cada día a las 9:00 AM"
echo "  Log: ${SCRIPT_DIR}/logs/cartola.log"
echo ""
echo "  Para desinstalar: launchctl unload $PLIST_PATH && rm $PLIST_PATH"
echo "  Para ejecutar ahora: launchctl start ${PLIST_ID}"
