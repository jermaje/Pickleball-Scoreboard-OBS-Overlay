# 🏓 Pickleball Scoreboard & OBS Overlay

Developed by **Jerson Mata Jemenez** for Camsur Pickleball Club.

A lightweight, web-based Pickleball Scoreboard system designed for live broadcasting, OBS (Open Broadcaster Software) overlays, and match logging. It features a responsive Control Panel, an OBS-friendly Display Overlay, and a Match History Log, all synchronized in real-time through a lightweight Python local server with Excel-based persistence.

---

## ✨ Features

- **🎮 Live Control Panel (`index.html`)**:
  - Score adjustments (increment/decrement, tap/click-to-score).
  - Serve tracking (Team 1 vs. Team 2).
  - Switch sides button (`↔️ Swap Sides`) and reset button (`🔄 Reset Score`).
  - Input field for custom Match Category (e.g., *Mixed Doubles*, *Men's Singles*).
  - Match saving (`💾 Save Match`) once a game is won.
  - Keyboard shortcuts toggle for rapid scoring.
  - Live server connection status indicator.
  
- **📺 OBS-Friendly Display Overlay (`display.html`)**:
  - Transparent background, ideal for OBS overlay windows or browser sources.
  - Displays match category, team names, live scores, and serving indicators.
  - Automatically updates every 500ms using server polling.
  - Local state fallback when offline.

- **📜 Match Results Log (`history.html`)**:
  - Displays a chronicle of all played matches grouped by Category.
  - Highlighted winners and trophy tags.
  - Chronological sorting (newest first).
  - Live synchronization with the broadcast backend.

- **💾 Excel-Backed Persistence (`server.py`)**:
  - Saves the active state (`state.xlsx`) automatically every 5 seconds to ensure no scores are lost during a crash or reload.
  - Appends finished games to a match history file (`MATCH.xlsx`).
  - Serves JSON API endpoints for state queries and updates.

- **⚡ Easy Startup (`run_scoreboard.bat`)**:
  - Automatic check and installation of the `openpyxl` python library.
  - Automated cleaning of legacy files.
  - Immediate server launch with a custom port/autosave interval parameter.

---

## 🚀 Getting Started

### Prerequisites

- **Python 3.x** installed on your system and added to your system `PATH`.

### Windows Quickstart

1. Double-click the **`run_scoreboard.bat`** file.
2. The batch script will automatically check for Python, verify/install `openpyxl`, and start the server on port `8000`.
3. Open your browser and navigate to:
   - **Control Panel**: [http://localhost:8000/index.html](http://localhost:8000/index.html) or [http://localhost:8000/](http://localhost:8000/)
   - **Display Overlay**: [http://localhost:8000/display.html](http://localhost:8000/display.html)
   - **Match History**: [http://localhost:8000/history.html](http://localhost:8000/history.html)

*Alternatively, run the batch script from command prompt to supply custom ports and autosave intervals:*
```cmd
run_scoreboard.bat [port] [autosave_interval_seconds]
```
Example:
```cmd
run_scoreboard.bat 8080 3
```

### Manual Command Line Startup (All OS)

Install the Excel support dependency and start the server manually:

```bash
# Install dependency
pip install openpyxl

# Start the server (default: port 8000, autosave every 5 seconds)
python server.py [port] [autosave_interval_seconds]
```

---

## ⌨️ Keyboard Shortcuts

Enable the **Keyboard** checkbox in the Control Panel footer to control the scoreboard without clicking:

| Key | Action |
|---|---|
| **`Q`** | Increment Team 1 Score (only if serving or serve switches) |
| **`A`** | Decrement Team 1 Score |
| **`P`** | Increment Team 2 Score (only if serving or serve switches) |
| **`L`** | Decrement Team 2 Score |
| **`S`** | Switch Serving Team (swaps serve dot) |
| **`R`** | Reset Score |

*Note: In Pickleball rules, points can only be scored by the serving team. The control panel enforces this! Tapping score for a non-serving team will switch the serve first.*

---

## 🎥 OBS Studio Setup

To display the scoreboard on your live stream:

1. In OBS, click **`+`** under the **Sources** dock.
2. Select **Browser**.
3. Name it (e.g., `Pickleball Overlay`).
4. In the URL field, enter: `http://localhost:8000/display.html`
5. Set the **Width** to `1200` and **Height** to `800` (or scale it to match your layout).
6. Set the **CSS** field to empty (or keep default if it doesn't affect the view).
7. Check "Refresh browser when scene becomes active" if desired.
8. Use standard OBS transform filters to position and scale it over your feed. The background of `display.html` is transparent by default!

---

## 📁 Repository Structure

```
├── MATCH.xlsx               # Match log history database (auto-generated)
├── display.html             # The broadcast viewer scoreboard overlay
├── history.html             # Historical record board web page
├── index.html               # Operator's Control Dashboard page
├── run_scoreboard.bat       # Windows batch script launcher
├── server.py                # Python HTTP Server and JSON REST API
├── state.xlsx               # Current live match state cache (auto-generated)
└── state.xlsx.bak           # Backup of the live match state cache
```

---

## 🔌 API Reference

The server exposes endpoints to programmatically control or read the scoreboard status:

### `GET /state`
Returns the active state of the match in JSON format.
**Response Example:**
```json
{
  "team1Score": 4,
  "team2Score": 2,
  "team1Name": "TEAM A",
  "team2Name": "TEAM B",
  "serving": 1,
  "servingNumber": 1,
  "matchConfig": {
    "bestOf": 1,
    "pointsToWin": 11,
    "doubles": false
  },
  "team1Sets": 0,
  "team2Sets": 0,
  "category": "MIXED DOUBLES",
  "lastUpdated": 1729012345678
}
```

### `POST /state`
Updates specific values in the scoreboard. Validates and saves the updated state.
**Request Payload Example:**
```json
{
  "team1Score": 5,
  "serving": 2
}
```

### `GET /history-data`
Returns a list of all finished match results recorded in `MATCH.xlsx`.

### `POST /save_match`
Triggers the server to record the current finished match state into `MATCH.xlsx` if a winner is decided.

---

## 🤝 Brand Details

This application is customized for **Camsur Pickleball Club**. Feel free to customize labels, fonts, and stylesheets in `index.html`, `display.html`, and `history.html` to fit your local club branding.

---

## 👤 Author

Created and developed by **Jerson Mata Jemenez**.
