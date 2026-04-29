#!/bin/bash
# Instala el sync diario a las 9:00 AM (ventas + inventario Shopify → Excel)
# USO: bash programar_sync_diario.sh

PLIST_ID="cl.nativaelements.sync"
PLIST_PATH="$HOME/Library/LaunchAgents/${PLIST_ID}.plist"
PYTHON=$(which python3)
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

mkdir -p "${SCRIPT_DIR}/logs"

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
        <string>${SCRIPT_DIR}/sync_diario.py</string>
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
    <string>${SCRIPT_DIR}/logs/sync_diario_stdout.log</string>
    <key>StandardErrorPath</key>
    <string>${SCRIPT_DIR}/logs/sync_diario_stderr.log</string>
    <key>EnvironmentVariables</key>
    <dict>
        <key>PATH</key>
        <string>/usr/local/bin:/usr/bin:/bin:/opt/homebrew/bin</string>
    </dict>
    <key>RunAtLoad</key>
    <false/>
</dict>
</plist>
EOF

launchctl unload "$PLIST_PATH" 2>/dev/null
launchctl load  "$PLIST_PATH"

echo ""
echo "✓ Sync diario instalado: ventas + inventario cada día a las 9:00 AM"
echo ""
echo "  Log mensual:  ${SCRIPT_DIR}/logs/sync_diario_YYYYMM.log"
echo "  Stdout:       ${SCRIPT_DIR}/logs/sync_diario_stdout.log"
echo ""
echo "  Ejecutar ahora:    launchctl start ${PLIST_ID}"
echo "  Ver estado:        launchctl list | grep nativa"
echo "  Desinstalar:       launchctl unload $PLIST_PATH && rm $PLIST_PATH"
