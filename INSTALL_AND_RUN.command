#!/bin/bash
# ============================================================
#  FINDATA — Financial Fundamental Data
#  One-Click Installer & Launcher for macOS
# ============================================================

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# Colors
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
NC='\033[0m'
BOLD='\033[1m'

clear
echo ""
echo -e "${CYAN}╔════════════════════════════════════════════════════════════╗${NC}"
echo -e "${CYAN}║${NC}    ${BOLD}${GREEN}FINDATA${NC} — Financial Fundamental Data                   ${CYAN}║${NC}"
echo -e "${CYAN}║${NC}    ${YELLOW}50,000+ tickers across 50+ global exchanges${NC}            ${CYAN}║${NC}"
echo -e "${CYAN}╚════════════════════════════════════════════════════════════╝${NC}"
echo ""

# Check Python
echo -e "${BLUE}▶${NC} Checking for Python 3..."
if command -v python3 &> /dev/null; then
    echo -e "${GREEN}✓${NC} Found $(python3 --version)"
else
    echo -e "${YELLOW}Installing Python via Homebrew...${NC}"
    if ! command -v brew &> /dev/null; then
        /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    fi
    brew install python3
fi

# Create venv
if [ ! -d ".venv" ]; then
    echo -e "${BLUE}▶${NC} Creating virtual environment..."
    python3 -m venv .venv
    echo -e "${GREEN}✓${NC} Virtual environment created"
fi

source .venv/bin/activate

# Install dependencies
echo -e "${BLUE}▶${NC} Checking dependencies..."
python3 -c "import flask, pandas, selenium, xlsxwriter" 2>/dev/null
if [ $? -ne 0 ]; then
    echo -e "${BLUE}▶${NC} Installing dependencies..."
    pip install --upgrade pip --quiet
    pip install -r requirements.txt --quiet
    echo -e "${GREEN}✓${NC} Dependencies installed"
else
    echo -e "${GREEN}✓${NC} All dependencies ready"
fi

# Create Desktop app
DESKTOP="$HOME/Desktop"
APP_PATH="$DESKTOP/FINDATA.app"

if [ ! -d "$APP_PATH" ]; then
    echo -e "${BLUE}▶${NC} Creating Desktop app..."
    mkdir -p "$APP_PATH/Contents/MacOS"
    
    cat > "$APP_PATH/Contents/MacOS/FINDATA" << EOF
#!/bin/bash
cd "$SCRIPT_DIR"
source .venv/bin/activate
python3 web_gui.py
EOF
    chmod +x "$APP_PATH/Contents/MacOS/FINDATA"
    
    cat > "$APP_PATH/Contents/Info.plist" << 'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key><string>FINDATA</string>
    <key>CFBundleName</key><string>FINDATA</string>
    <key>CFBundleVersion</key><string>1.0</string>
    <key>CFBundlePackageType</key><string>APPL</string>
</dict>
</plist>
EOF
    echo -e "${GREEN}✓${NC} Desktop app created - drag to Dock for quick access!"
fi

echo ""
echo -e "${GREEN}════════════════════════════════════════════════════════════${NC}"
echo -e "${BOLD}${GREEN}🚀 Starting FINDATA...${NC}"
echo -e "   ${CYAN}URL:${NC} http://127.0.0.1:5050"
echo -e "   ${CYAN}Stop:${NC} Close this window or Ctrl+C"
echo -e "${GREEN}════════════════════════════════════════════════════════════${NC}"
echo ""

python3 web_gui.py
