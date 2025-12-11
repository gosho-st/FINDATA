#!/bin/bash
# ============================================================
#  FINDATA — Quick Launcher
#  Run this after installation for daily use
# ============================================================

cd "$(dirname "$0")"

# Check if venv exists, if not create it
if [ ! -d ".venv" ]; then
    echo "⏳ First run - setting up environment..."
    python3 -m venv .venv
    source .venv/bin/activate
    pip install --upgrade pip --quiet
    pip install -r requirements.txt --quiet
    echo "✅ Setup complete!"
else
    source .venv/bin/activate
fi

echo ""
echo "════════════════════════════════════════════════════════════"
echo "  🚀 FINDATA is running!"
echo "  📍 URL: http://127.0.0.1:5050"
echo "  🛑 To stop: Close this window"
echo "════════════════════════════════════════════════════════════"
echo ""

# Open browser after short delay
(sleep 2 && open "http://127.0.0.1:5050") &

python3 web_gui.py
