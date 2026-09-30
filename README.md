# Fantasy Trade Calculator

Fantasy Trade Calculator is a Windows desktop app for exploring your Sleeper fantasy football league. It brings roster and matchup information together with player value estimates and history so you can review trades, compare teams, and look for roster upgrades.

The app uses Sleeper's public read-only API and does not need your Sleeper password or an API token. FantasyCalc and Stats Guy Fantasy provide separate player value estimates; these are useful reference points, not guarantees of a player's future performance or what a league mate will accept.

## For league mates

1. Open `FantasyTradeCalculator.exe`.
2. The first time, enter your Sleeper username, choose your league, and load it. The app saves your username and last selected league on this computer.
3. On later launches, it tries to reload the saved league and refresh its roster data automatically.
4. Use the tabs to explore your league. Choose **Refresh Player Data** if you want to download Sleeper's player directory again.

The executable includes Python, the app, and its required libraries. You do not need to install Python, open a terminal, or install dependencies to use it. Windows may show a security prompt because the app is not code-signed; only run an executable from someone you trust.

## App features

- **Position Needs** compares teams' roster depth with their league lineup requirements. It is a roster-count overview, not a player projection.
- **Roster View** shows a team's players grouped by their Sleeper lineup slot, including starters, bench, and IR/reserve. Select a player to view available value history and roster details.
- **Weekly Help** brings together matchup context, bye-week coverage, available free agents, and weekly projections when Sleeper provides them. FantasyCalc market values are not weekly point projections.
- **Value Trends** compares FantasyCalc and Stats Guy Fantasy values, shows which team rostered each player (or whether the player is a free agent), and graphs historical value movement. Select the history window and players to compare trends where data is available.
- **Trade Review** lets you choose players from two teams, compare totals from both value sources, inspect roster-depth impact, and copy a trade proposal. The player list can be sorted by last name, position priority (QB, RB, WR, TE, K), status, NFL team, or value.
- **Saved Trades** stores proposals, accepted trades, and rejected trades on this computer. Reopen a saved trade to review or edit it; the displayed totals use currently loaded market values.
- **Trade Targets** suggests possible one-for-one trade ideas based on roster fit and market values. Treat suggestions as starting points to adjust, not automatic recommendations.
- **Export to Excel** creates a workbook with one sheet per team, named after the team in Sleeper. Each roster is ordered with starters first in lineup order, then bench, then IR/reserve. The columns are Position, Player, NFL Team, Status, Injury Status, and Slot. After saving, the app asks whether you want to open the workbook.
- **Dark mode** lets you switch the app's appearance.

FantasyCalc data is attributed in the app and links to [FantasyCalc](https://fantasycalc.com/). Historical value data is attributed to [Stats Guy Fantasy](https://statsguyfantasy.com/).

## Saved information and data use

Settings, saved trades, and cached player/value information are stored in `.sleeper_fantasy_exporter` inside your Windows user folder. Your Sleeper username, selected league, and saved trades remain on your computer. The app contacts Sleeper to retrieve public league data, FantasyCalc for market values, and Stats Guy Fantasy for values and historical trends. The app does not request or store a Sleeper password.

League roster data refreshes when the saved league is loaded. The Sleeper player directory is cached for up to 24 hours, FantasyCalc values for up to six hours, and other API data may use shorter caches to avoid unnecessary repeat downloads.

## Build the Windows executable from source

Building the executable is intended for the person preparing a release; league mates can use the executable without setting up Python.

1. Install Python 3.10 or newer with Tcl/Tk support enabled.
2. Open PowerShell in this project folder.
3. Run:

   ```powershell
   .\build_windows.ps1
   ```

The build script creates an isolated `.build-venv`, installs the packages listed in `requirements.txt` and PyInstaller, then writes `build\dist\FantasyTradeCalculator.exe`. It does not change your system Python packages. The generated build files are excluded from Git.

To run the app directly from source, install the packages in `requirements.txt` in a Python environment that includes Tcl/Tk, then run `python sleeper_exporter.py`.
