#This file contains the desktop app, its Sleeper and FantasyCalc API helpers, and the Excel export logic.
import json
import itertools
import hashlib
import os
import csv
import re
import shutil
import statistics
import subprocess
import sys
import tempfile
import threading
import uuid
import zipfile
from collections import Counter
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import requests
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

#This URL is the root used by Sleeper's read-only API.
BASE_URL = "https://api.sleeper.app/v1"
APP_VERSION = "1.1.1"
GITHUB_RELEASE_API = "https://api.github.com/repos/jasonBuras/SleeperFantasyCalculator/releases/latest"
UPDATE_ASSET_NAME = "FantasyTradeCalculator-Windows.zip"

def resource_path(filename):
    """Return an asset path for source runs and PyInstaller one-file builds."""
    bundle_dir = getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)
    return Path(bundle_dir) / filename

#This folder keeps app settings and cached data in the current Windows user's home directory.
CACHE_DIR = Path.home() / ".sleeper_fantasy_exporter"
#Sleeper's large player directory is cached so it is not downloaded every time the app opens.
PLAYER_CACHE = CACHE_DIR / "players_nfl.json"
#The selected username, league, season, format, and appearance preference are stored here.
SETTINGS_FILE = CACHE_DIR / "settings.json"
PLAYER_CACHE_MAX_AGE_HOURS = 24
#FantasyCalc values use a separate cache because they change more often than Sleeper's player IDs.
FANTASYCALC_URL = "https://api.fantasycalc.com/values/current"
FANTASYCALC_CACHE = CACHE_DIR / "fantasycalc_values.json"
FANTASYCALC_CACHE_MAX_AGE_HOURS = 6
#The published NFL schedule is used to derive team bye weeks absent from Sleeper's player map.
NFL_SCHEDULE_URL = "https://raw.githubusercontent.com/nflverse/nfldata/master/data/games.csv"
NFL_SCHEDULE_CACHE = CACHE_DIR / "nfl_games.csv"
NFL_SCHEDULE_CACHE_MAX_AGE_HOURS = 12
#Saved trade ideas are stored locally and grouped by Sleeper league ID.
SAVED_TRADES_FILE = CACHE_DIR / "saved_trades.json"
#Stats Guy Fantasy provides an independent Sleeper-ID-keyed value board and daily value history.
STATSGUY_BASE_URL = "https://api.statsguyfantasy.com/api/v1"
STATSGUY_CACHE = CACHE_DIR / "stats_guy_values.json"
STATSGUY_CACHE_MAX_AGE_HOURS = 12
SLEEPER_PROJECTION_CACHE = CACHE_DIR / "sleeper_projections.json"
SLEEPER_PROJECTION_CACHE_MAX_AGE_HOURS = 6
VALUE_HISTORY_WINDOWS = {
    "7 days": 8,
    "14 days": 15,
    "1 month": 31,
    "3 months": 91,
    "6 months": 181,
    "1 year": 366,
    "All available": None,
}


#This custom error lets API problems be shown as readable app messages.
class SleeperAPIError(Exception):
    pass


#This class wraps the HTTP calls made to Sleeper and FantasyCalc.
class SleeperAPI:
    def __init__(self):
        #Reuse one HTTP session so requests share connection settings and the app identifier.
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "FantasyTradeCalculator/1.0"
        })
        self.stats_guy_history_cache = {}

    def get(self, path):
        #Make a Sleeper API request and convert network or JSON errors to app-friendly messages.
        url = f"{BASE_URL}{path}"
        try:
            response = self.session.get(url, timeout=20)
            response.raise_for_status()
            return response.json()
        except requests.RequestException as exc:
            raise SleeperAPIError(f"Could not reach Sleeper API:\n{exc}") from exc
        except ValueError as exc:
            raise SleeperAPIError("Sleeper returned invalid JSON.") from exc

    def get_user(self, username):
        #Look up a Sleeper account by username and reject empty results.
        data = self.get(f"/user/{username.strip()}")
        if not data:
            raise SleeperAPIError("Sleeper user was not found.")
        return data

    def get_leagues(self, user_id, season):
        #Return the user's NFL leagues for the requested season.
        return self.get(f"/user/{user_id}/leagues/nfl/{season}") or []

    def get_league(self, league_id):
        #Fetch league settings such as scoring and starting roster positions.
        return self.get(f"/league/{league_id}")

    def get_rosters(self, league_id):
        #Fetch every fantasy team's roster and current starter IDs.
        return self.get(f"/league/{league_id}/rosters") or []

    def get_users(self, league_id):
        #Fetch manager names and team-name metadata for the league.
        return self.get(f"/league/{league_id}/users") or []

    def get_nfl_state(self):
        #Get Sleeper's current NFL season, week, and display week.
        return self.get("/state/nfl") or {}

    def get_matchups(self, league_id, week):
        #Get each fantasy team's starters and current score for one league week.
        return self.get(f"/league/{league_id}/matchups/{week}") or []

    def get_projections(self, season, week, ppr):
        #Read Sleeper's weekly projection feed and cache it briefly for this season/week.
        metric = "pts_ppr" if ppr >= 0.75 else "pts_half_ppr" if ppr >= 0.25 else "pts_std"
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        cache_key = f"{season}:{week}:{metric}"
        try:
            age_hours = (datetime.now().timestamp() - SLEEPER_PROJECTION_CACHE.stat().st_mtime) / 3600
            if age_hours < SLEEPER_PROJECTION_CACHE_MAX_AGE_HOURS:
                cached = json.loads(SLEEPER_PROJECTION_CACHE.read_text(encoding="utf-8"))
                if cached.get("key") == cache_key and isinstance(cached.get("values"), dict):
                    return cached["values"], metric
        except (OSError, ValueError, AttributeError):
            pass
        url = f"https://api.sleeper.com/projections/nfl/{season}/{week}"
        try:
            response = self.session.get(
                url, params={"season_type": "regular", "order_by": metric}, timeout=25,
            )
            response.raise_for_status()
            payload = response.json()
            rows = payload.get("players", []) if isinstance(payload, dict) else payload
            if not isinstance(rows, list):
                raise SleeperAPIError("Sleeper returned projections in an unexpected format.")
            values = {}
            for row in rows:
                if not isinstance(row, dict) or row.get("player_id") in (None, ""):
                    continue
                stats = row.get("stats") if isinstance(row.get("stats"), dict) else row
                points = stats.get(metric)
                if points is not None:
                    try:
                        values[str(row["player_id"])] = float(points)
                    except (TypeError, ValueError):
                        continue
            try:
                SLEEPER_PROJECTION_CACHE.write_text(
                    json.dumps({"key": cache_key, "values": values}), encoding="utf-8",
                )
            except OSError:
                pass
            return values, metric
        except requests.RequestException as exc:
            raise SleeperAPIError(f"Could not load Sleeper weekly projections:\n{exc}") from exc

    def get_nfl_schedule(self, season):
        #Read the current schedule from cache or download nflverse's maintained game schedule.
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        schedule_text = None
        try:
            age_hours = (datetime.now().timestamp() - NFL_SCHEDULE_CACHE.stat().st_mtime) / 3600
            if age_hours < NFL_SCHEDULE_CACHE_MAX_AGE_HOURS:
                schedule_text = NFL_SCHEDULE_CACHE.read_text(encoding="utf-8")
        except OSError:
            pass
        if schedule_text is None:
            try:
                response = self.session.get(NFL_SCHEDULE_URL, timeout=30)
                response.raise_for_status()
                schedule_text = response.text
                try:
                    NFL_SCHEDULE_CACHE.write_text(schedule_text, encoding="utf-8")
                except OSError:
                    #A valid download remains usable even if the local cache cannot be written.
                    pass
            except requests.RequestException as exc:
                raise SleeperAPIError(f"Could not load NFL bye-week schedule:\n{exc}") from exc
        try:
            return [
                row for row in csv.DictReader(schedule_text.splitlines())
                if str(row.get("season", "")) == str(season)
                and str(row.get("game_type", "")).upper() == "REG"
            ]
        except (csv.Error, TypeError) as exc:
            raise SleeperAPIError("The NFL schedule data could not be read.") from exc

    def get_stats_guy_values(self):
        #Load all player cards once and reuse them while the API snapshot is still fresh.
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        try:
            age_hours = (datetime.now().timestamp() - STATSGUY_CACHE.stat().st_mtime) / 3600
            if age_hours < STATSGUY_CACHE_MAX_AGE_HOURS:
                cached = json.loads(STATSGUY_CACHE.read_text(encoding="utf-8"))
                if isinstance(cached, dict) and isinstance(cached.get("players"), list):
                    return cached
        except (OSError, ValueError, AttributeError):
            pass
        try:
            response = self.session.get(f"{STATSGUY_BASE_URL}/players", timeout=20)
            response.raise_for_status()
            data = response.json()
            if not isinstance(data, dict) or not isinstance(data.get("players"), list):
                raise SleeperAPIError("Stats Guy Fantasy returned player values in an unexpected format.")
            try:
                STATSGUY_CACHE.write_text(json.dumps(data), encoding="utf-8")
            except OSError:
                #The downloaded values are usable even if the local cache cannot be written.
                pass
            return data
        except requests.RequestException as exc:
            raise SleeperAPIError(f"Could not load Stats Guy Fantasy values:\n{exc}") from exc

    def get_stats_guy_history(self, player_id, value_format, window=30):
        #Fetch one player's historical snapshots and cache them for the rest of this app session.
        cache_key = (str(player_id), str(value_format), int(window) if window is not None else None)
        if cache_key in self.stats_guy_history_cache:
            return self.stats_guy_history_cache[cache_key]
        try:
            params = {"format": value_format}
            if window is not None:
                params["window"] = int(window)
            response = self.session.get(
                f"{STATSGUY_BASE_URL}/players/{player_id}/value-history",
                params=params, timeout=20,
            )
            response.raise_for_status()
            data = response.json()
            if not isinstance(data, dict) or not isinstance(data.get("history"), list):
                raise SleeperAPIError("Stats Guy Fantasy returned player history in an unexpected format.")
            self.stats_guy_history_cache[cache_key] = data
            return data
        except requests.RequestException as exc:
            raise SleeperAPIError(f"Could not load Stats Guy Fantasy value history:\n{exc}") from exc

    def get_players(self, force_refresh=False):
        #Use the local player map while it is fresh unless the user asks to refresh it.
        CACHE_DIR.mkdir(parents=True, exist_ok=True)

        if PLAYER_CACHE.exists() and not force_refresh:
            age_hours = (
                datetime.now().timestamp() - PLAYER_CACHE.stat().st_mtime
            ) / 3600
            if age_hours < PLAYER_CACHE_MAX_AGE_HOURS:
                with PLAYER_CACHE.open("r", encoding="utf-8") as f:
                    return json.load(f)

        #Download the current player map and save it for later roster-ID lookups.
        data = self.get("/players/nfl")
        with PLAYER_CACHE.open("w", encoding="utf-8") as f:
            json.dump(data, f)
        return data

    def get_fantasycalc_values(self, params):
        #Load market values from the format-specific cache or fetch them from FantasyCalc.
        """Fetch current market values, using a short-lived format-specific cache."""
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        cache_key = json.dumps(params, sort_keys=True)
        try:
            if FANTASYCALC_CACHE.exists():
                cache = json.loads(FANTASYCALC_CACHE.read_text(encoding="utf-8"))
                age = (datetime.now().timestamp() - FANTASYCALC_CACHE.stat().st_mtime) / 3600
                if age < FANTASYCALC_CACHE_MAX_AGE_HOURS and cache.get("key") == cache_key:
                    return cache.get("values", [])
        except (OSError, ValueError, AttributeError):
            pass

        try:
            response = self.session.get(FANTASYCALC_URL, params=params, timeout=20)
            response.raise_for_status()
            values = response.json()
            if not isinstance(values, list):
                raise SleeperAPIError("FantasyCalc returned data in an unexpected format.")
            try:
                FANTASYCALC_CACHE.write_text(
                    json.dumps({"key": cache_key, "values": values}), encoding="utf-8"
                )
            except OSError:
                pass
            return values
        except requests.RequestException as exc:
            raise SleeperAPIError(f"Could not reach FantasyCalc:\n{exc}") from exc


def player_name(player):
    #Build a readable name from the fields available in Sleeper's player record.
    full_name = player.get("full_name")
    if full_name:
        return full_name

    first = player.get("first_name") or ""
    last = player.get("last_name") or ""
    name = f"{first} {last}".strip()

    return name or player.get("search_full_name") or player.get("player_id", "")


def build_data(api, league_id, status_callback, force_player_refresh=False):
    #Download league inputs and turn them into tables used by the app and workbook.
    #The callback updates the status line while these requests run in the background.
    status_callback("Loading league information...")
    league = api.get_league(league_id)

    status_callback("Loading league users...")
    users = api.get_users(league_id)

    status_callback("Loading rosters...")
    rosters = api.get_rosters(league_id)

    status_callback("Loading player database (cached when possible)...")
    players = api.get_players(force_refresh=force_player_refresh)

    #Index managers by ID so each roster can be paired with its team name and owner.
    user_by_id = {str(u.get("user_id")): u for u in users}

    team_rows = []
    roster_rows = []
    player_rows = []
    needs_rows = []

    #Build one set of team, roster, and positional-depth rows for every league roster.
    team_rosters = []
    for team_number, roster in enumerate(rosters, start=1):
        roster_id = roster.get("roster_id")
        owner_id = str(roster.get("owner_id") or "")
        user = user_by_id.get(owner_id, {})

        metadata = user.get("metadata") or {}
        team_name = (
            metadata.get("team_name")
            or metadata.get("mention_name")
            or user.get("display_name")
            or user.get("username")
            or f"Roster {roster_id}"
        )
        manager = (
            user.get("display_name")
            or user.get("username")
            or owner_id
        )

        team_rows.append({
            "Roster ID": roster_id,
            "Team Name": team_name,
            "Manager": manager,
            "Username": user.get("username", ""),
            "Owner ID": owner_id,
        })

        team_rosters.append({
            "sheet_name": team_name,
            "roster_id": roster_id,
            "rows": [],
        })

        #Sleeper lists starters in lineup order; pair those IDs with league lineup slots.
        starter_order = [str(x) for x in (roster.get("starters") or [])]
        starters = set(starter_order)
        bench_slots = {"BN", "IR", "TAXI", "RESERVE"}
        starting_slots = [
            str(position) for position in (league.get("roster_positions") or [])
            if str(position).upper() not in bench_slots
        ]
        starter_slot_by_id = {
            player_id: starting_slots[index]
            for index, player_id in enumerate(starter_order)
            if player_id and player_id != "0" and index < len(starting_slots)
        }
        all_players = roster.get("players") or []

        #Count required direct-position slots separately from flexible slots.
        position_slots = Counter(
            str(slot).upper() for slot in (league.get("roster_positions") or [])
            if str(slot).upper() in {"QB", "RB", "WR", "TE", "K", "DEF", "DST"}
        )
        #Count each rostered player's eligible positions to estimate team depth.
        position_counts = Counter()
        starter_counts = Counter()
        for raw_id in all_players:
            p = players.get(str(raw_id), {}) if isinstance(players, dict) else {}
            eligible = p.get("fantasy_positions") or ([p.get("position")] if p.get("position") else [])
            for pos in eligible:
                position_counts[str(pos).upper()] += 1
        #Attribute flex starters to the most underfilled eligible position for a rough depth view.
        for player_id in starter_order:
            starter_slot = starter_slot_by_id.get(player_id, "").upper()
            if starter_slot in {"QB", "RB", "WR", "TE", "K", "DEF", "DST"}:
                starter_counts[starter_slot] += 1
            elif starter_slot in {"FLEX", "WRRB", "WRT", "REC_FLEX", "SUPER_FLEX", "SF"}:
                p = players.get(player_id, {}) if isinstance(players, dict) else {}
                eligible = [str(x).upper() for x in (p.get("fantasy_positions") or [p.get("position")]) if x]
                if starter_slot in {"SUPER_FLEX", "SF"}:
                    eligible = ["QB", "RB", "WR", "TE"]
                positions = [x for x in eligible if x in {"QB", "RB", "WR", "TE", "K", "DEF"}]
                if starter_slot == "WRRB":
                    positions = [x for x in positions if x in {"RB", "WR"}]
                elif starter_slot == "WRT":
                    positions = [x for x in positions if x in {"WR", "TE"}]
                if positions:
                    order = ("QB", "RB", "WR", "TE", "K", "DEF")
                    target = max(positions, key=lambda x: (position_slots.get(x, 0) - starter_counts.get(x, 0), -order.index(x)))
                    starter_counts[target] += 1
        #Compare starters and total roster counts with each required lineup position.
        for pos in ("QB", "RB", "WR", "TE", "K", "DEF"):
            required = position_slots.get(pos, 0) + (position_slots.get("DST", 0) if pos == "DEF" else 0)
            if not required:
                continue
            rostered = position_counts.get(pos, 0) + (position_counts.get("DST", 0) if pos == "DEF" else 0)
            starting = starter_counts.get(pos, 0) + (starter_counts.get("DST", 0) if pos == "DEF" else 0)
            needs_rows.append({
                "Team": team_name, "Roster ID": roster_id, "Position": pos,
                "Required": required, "Starting": starting, "Rostered": rostered,
                "Bench": max(0, rostered - starting),
                "Assessment": "Starter needed" if starting < required else (
                    "Thin depth" if pos not in {"K", "DEF"} and rostered <= required else "Covered"
                ),
            })

        reserve = set(str(x) for x in (roster.get("reserve") or []))

        #Create a detailed player row and preserve the starters/bench/reserve grouping.
        for roster_order, raw_player_id in enumerate(all_players):
            player_id = str(raw_player_id)
            player = players.get(player_id, {}) if isinstance(players, dict) else {}

            slot = starter_slot_by_id.get(player_id, "Starter") if player_id in starters else "Bench"
            if player_id in reserve:
                slot = "Reserve"

            fantasy_positions = player.get("fantasy_positions") or []
            position = ", ".join(fantasy_positions) or player.get("position") or ""

            roster_row = {
                "Roster ID": roster_id,
                "Team": team_name,
                "Manager": manager,
                "Player": player_name(player) if player else player_id,
                "First Name": player.get("first_name") or "",
                "Last Name": player.get("last_name") or "",
                "Player ID": player_id,
                "Roster Order": roster_order,
                "Lineup Order": starter_order.index(player_id) if player_id in starters else None,
                "Position": position,
                "NFL Team": player.get("team") or "",
                "Status": player.get("status") or "",
                "Injury Status": player.get("injury_status") or "",
                "Slot": slot,
                "Number": player.get("number") or "",
                "Bye Week": player.get("bye_week") or "",
            }
            roster_rows.append(roster_row)
            row_order = (
                (0, starter_order.index(player_id))
                if player_id in starters and player_id not in reserve
                else (2, roster_order) if player_id in reserve
                else (1, roster_order)
            )
            team_rosters[-1]["rows"].append({
                "Position": position,
                "Player": roster_row["Player"],
                "NFL Team": roster_row["NFL Team"],
                "Status": roster_row["Status"],
                "Injury Status": roster_row["Injury Status"],
                "Slot": slot,
                "__order": row_order,
            })

    #Include a compact player table for league roster members only.
    used_player_ids = {row["Player ID"] for row in roster_rows}
    free_agent_rows = []
    for raw_id, player in players.items():
        player_id = str(raw_id)
        eligible = player.get("fantasy_positions") or ([player.get("position")] if player.get("position") else [])
        nfl_team = str(player.get("team") or "").strip()
        if player_id in used_player_ids or not eligible or not nfl_team:
            continue
        free_agent_rows.append({
            "Player ID": player_id,
            "Player": player_name(player),
            "Position": ", ".join(str(position) for position in eligible),
            "NFL Team": nfl_team,
            "Status": player.get("status") or "",
            "Injury Status": player.get("injury_status") or "",
            "Bye Week": player.get("bye_week") or "",
        })

    for player_id in sorted(used_player_ids):
        player = players.get(player_id, {}) if isinstance(players, dict) else {}
        fantasy_positions = player.get("fantasy_positions") or []
        player_rows.append({
            "Player ID": player_id,
            "Player": player_name(player) if player else player_id,
            "First Name": player.get("first_name") or "",
            "Last Name": player.get("last_name") or "",
            "Position": ", ".join(fantasy_positions) or player.get("position") or "",
            "NFL Team": player.get("team") or "",
            "Status": player.get("status") or "",
            "Injury Status": player.get("injury_status") or "",
            "Number": player.get("number") or "",
            "College": player.get("college") or "",
            "Age": player.get("age") or "",
            "Bye Week": player.get("bye_week") or "",
        })

    #Store league metadata and app/cache notes for the export data model.
    league_info = [{
        "League Name": league.get("name", ""),
        "League ID": league.get("league_id", league_id),
        "Season": league.get("season", ""),
        "Sport": league.get("sport", ""),
        "Season Type": league.get("season_type", ""),
        "Status": league.get("status", ""),
        "Teams": league.get("total_rosters", ""),
        "Exported": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Roster Positions": ", ".join(league.get("roster_positions") or []),
    }]

    settings = [{
        "Setting": "Program",
        "Value": "Fantasy Trade Calculator 1.0",
    }, {
        "Setting": "Player Cache",
        "Value": str(PLAYER_CACHE),
    }, {
        "Setting": "Player Cache Refresh",
        "Value": f"Every {PLAYER_CACHE_MAX_AGE_HOURS} hours unless manually refreshed",
    }]

    #Keep workbook-only tables in the result; team tabs are added separately below.
    result = {
        "League Info": pd.DataFrame(league_info),
        "Teams": pd.DataFrame(team_rows),
        "Rosters": pd.DataFrame(roster_rows),
        "Position Needs": pd.DataFrame(needs_rows),
        "Players": pd.DataFrame(player_rows),
        "Free Agents": pd.DataFrame(free_agent_rows),
        "Settings": pd.DataFrame(settings),
    }
    #Make each team sheet name safe and unique under Excel's naming limits.
    used_sheet_names = set()
    for team_number, team in enumerate(team_rosters, start=1):
        # Excel sheet names cannot contain these characters and are limited to 31 chars.
        base_name = "".join(
            "_" if character in "[]:*?/\\" else character
            for character in str(team["sheet_name"])
        ).strip("'")[:31] or f"Team {team_number}"
        sheet_name = base_name
        suffix = 2
        while sheet_name.casefold() in used_sheet_names:
            ending = f" ({suffix})"
            sheet_name = f"{base_name[:31 - len(ending)]}{ending}"
            suffix += 1
        used_sheet_names.add(sheet_name.casefold())
        result[sheet_name] = pd.DataFrame(
            sorted(team["rows"], key=lambda row: row["__order"]),
            columns=["Position", "Player", "NFL Team", "Status", "Injury Status", "Slot"],
        )
    return result


def export_excel(dataframes, filename):
    #Write the team roster tabs, then apply readable headers, filters, and column widths.
    internal_sheets = {"League Info", "Teams", "Rosters", "Position Needs", "Players", "Free Agents", "Settings"}
    team_sheets = {
        name: df for name, df in dataframes.items()
        if name not in internal_sheets
    }
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        for sheet_name, df in team_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    workbook = load_workbook(filename)

    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        for cell in worksheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4F81BD")
            cell.alignment = Alignment(horizontal="center")

        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column_cells[0].column)

            for cell in column_cells:
                value = "" if cell.value is None else str(cell.value)
                max_length = max(max_length, len(value))

            worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 45)

    workbook.save(filename)


#This class owns the window, app state, tabs, and event handlers.
class App(tk.Tk):
    def __init__(self):
        #Create the window and initialize data before building controls that use it.
        super().__init__()

        self.title("Fantasy Trade Calculator")
        try:
            self.iconbitmap(default=str(resource_path("app_icon.ico")))
        except tk.TclError:
            #Keep startup working if the optional window icon cannot be loaded.
            pass

        self.api = SleeperAPI()
        self.user = None
        self.leagues = []
        self.dataframes = None
        self.loaded_username = ""
        self.fantasycalc_values = {}
        self.fantasycalc_trends = {}
        self.fantasycalc_overall_ranks = {}
        self.fantasycalc_position_ranks = {}
        self.stats_guy_players = {}
        self.stats_guy_values_asof = {}
        self.stats_guy_error = ""
        self.value_trend_rows = []
        self.value_trend_sort_reverse = {}
        self.value_history_cache = {}
        self.value_history_window_var = tk.StringVar(value="3 months")
        self.selected_history_player_id = None
        self.trade_history_series = []
        self.trade_history_signature = None
        self.trade_history_request_id = 0
        self.trade_history_metric_var = tk.StringVar(value="Percent Change")
        self.trade_history_plot_points = {}
        self.trade_history_window = None
        self.trade_history_popup_canvas = None
        self.player_history_window = None
        self.player_history_popup_canvas = None
        self.player_history_popup_source = None
        self.player_history_popup_points_name = None
        self.player_history_popup_empty_text = None
        self.player_history_popup_scale = None
        self.team_names = []
        self.team_options = []
        self._trade_populating = False
        self.trade_sort_state = {}
        self.trade_sort_info = {}
        self._background_task_count = 0
        self.weekly_state = {}
        self.weekly_matchups = []
        self.weekly_bye_weeks = {}
        self.weekly_context_error = ""
        self.weekly_schedule_error = ""
        self.weekly_sort_reverse = {}
        self.weekly_priority_positions = set()
        self.weekly_waiver_rows = []
        self.weekly_team_roster_id = None
        self.roster_projection_values = {}
        self.roster_projection_metric = ""
        self.roster_projection_error = ""
        self.roster_history_player_id = None
        self.roster_history_points = []
        self.roster_search_items = []
        self.weekly_history_player_id = None
        self.weekly_history_points = []
        #Keep trades and user preferences between app sessions.
        self.saved_trades = self._load_saved_trades()
        self.active_saved_trade_id = None
        self.active_counteroffer_parent_id = None
        self.saved_settings = self._load_settings()
        self.startup_autoload_pending = bool(
            self.saved_settings.get("username") and self.saved_settings.get("league_id")
        )

        self.username_var = tk.StringVar(value=self.saved_settings.get("username", ""))
        self.season_var = tk.StringVar(
            value=self.saved_settings.get("season", str(datetime.now().year))
        )
        self.league_var = tk.StringVar()
        self.window_context_var = tk.StringVar(value="Select a league in Main Menu")
        self.status_var = tk.StringVar(value="Enter a Sleeper username and load leagues.")
        self.update_status_var = tk.StringVar(value=f"Installed version {APP_VERSION}. Check for updates when you're ready.")
        self.update_notes_var = tk.StringVar(value="Release notes will appear here after checking for updates.")
        self.available_update = None
        self.copy_trade_preview_var = tk.StringVar(
            value="Select at least one player from each team to preview the copied text."
        )
        self.dark_mode_var = tk.BooleanVar(value=bool(self.saved_settings.get("dark_mode", False)))
        self.attribution_labels = []
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass
        self._apply_theme()

        #Apply the user's saved appearance, then build the visible app and settings hooks.
        self._build_ui()
        self._compact_tree_columns()
        self._size_window_to_screen()
        self.username_var.trace_add("write", self._save_settings)
        self.season_var.trace_add("write", self._save_settings)
        if self.startup_autoload_pending:
            self.after(100, self.load_leagues)

    def _load_settings(self):
        #Read the local preferences file; missing or damaged settings should not block startup.
        try:
            with SETTINGS_FILE.open("r", encoding="utf-8") as settings_file:
                settings = json.load(settings_file)
                return settings if isinstance(settings, dict) else {}
        except (OSError, ValueError, AttributeError):
            return {}

    def _save_settings(self, *_):
        #Write current choices without interrupting the user if the folder is unavailable.
        self.saved_settings["username"] = self.username_var.get().strip()
        self.saved_settings["season"] = self.season_var.get().strip()
        if hasattr(self, "format_var"):
            self.saved_settings["format"] = self.format_var.get()
        if hasattr(self, "dark_mode_var"):
            self.saved_settings["dark_mode"] = self.dark_mode_var.get()
        try:
            CACHE_DIR.mkdir(parents=True, exist_ok=True)
            with SETTINGS_FILE.open("w", encoding="utf-8") as settings_file:
                json.dump(self.saved_settings, settings_file)
        except OSError:
            # Keep the UI usable if the settings folder is unavailable.
            pass

    def _update_window_context(self):
        #Keep the active username and league visible while users move between tabs.
        username = self.loaded_username or self.username_var.get().strip()
        league_name = ""
        if self.dataframes is not None:
            try:
                league_name = str(self.dataframes["League Info"].iloc[0]["League Name"])
            except (KeyError, IndexError, AttributeError):
                league_name = ""
        else:
            league = self.selected_league() if hasattr(self, "league_combo") else None
            if league and str(league.get("season", self.season_var.get())) == self.season_var.get().strip():
                league_name = str(league.get("name") or "Unnamed League")
        if username and league_name:
            text = f"{username}  |  {league_name}"
        elif username:
            text = f"{username}  |  Choose a league in Main Menu"
        else:
            text = "Choose a Sleeper league in Main Menu"
        self.window_context_var.set(text)

    def _apply_theme(self):
        #Set colors for all ttk controls so the toggle changes the full app consistently.
        dark = bool(self.dark_mode_var.get())
        colors = ({
            "background": "#1f2329",
            "surface": "#292f36",
            "field": "#343b44",
            "foreground": "#edf1f5",
            "muted": "#b5bec8",
            "accent": "#3978b8",
            "hover": "#3a4653",
            "selected": "#285b89",
            "border": "#49535e",
        } if dark else {
            "background": "#f1f3f5",
            "surface": "#ffffff",
            "field": "#ffffff",
            "foreground": "#20252b",
            "muted": "#5d6670",
            "accent": "#3978b8",
            "hover": "#e4eaf0",
            "selected": "#cfe3f6",
            "border": "#c5cbd1",
        })
        bg, surface, field = colors["background"], colors["surface"], colors["field"]
        fg, muted, accent = colors["foreground"], colors["muted"], colors["accent"]
        self.theme_colors = colors
        self.configure(background=bg)
        self.style.configure(".", background=bg, foreground=fg)
        self.style.configure("TFrame", background=bg)
        self.style.configure("TLabel", background=bg, foreground=fg)
        self.style.configure("TLabelframe", background=bg, foreground=fg, bordercolor=colors["border"])
        self.style.configure("TLabelframe.Label", background=bg, foreground=fg)
        self.style.configure("TButton", background=surface, foreground=fg, bordercolor=colors["border"], padding=(9, 5))
        self.style.map("TButton", background=[("active", colors["hover"]), ("pressed", accent)],
                       foreground=[("disabled", muted)])
        self.style.configure("TCheckbutton", background=bg, foreground=fg)
        self.style.map("TCheckbutton", background=[("active", bg)], foreground=[("disabled", muted)])
        self.style.configure("TEntry", fieldbackground=field, foreground=fg, insertcolor=fg,
                             bordercolor=colors["border"])
        self.style.configure("TCombobox", fieldbackground=field, background=surface,
                             foreground=fg, arrowcolor=fg, bordercolor=colors["border"])
        self.style.map("TCombobox", fieldbackground=[("readonly", field), ("disabled", bg)],
                       foreground=[("readonly", fg), ("disabled", muted)],
                       selectbackground=[("readonly", field)], selectforeground=[("readonly", fg)])
        self.style.configure("TNotebook", background=bg, bordercolor=colors["border"])
        self.style.configure("TNotebook.Tab", background=surface, foreground=fg, padding=(12, 6))
        self.style.map("TNotebook.Tab", background=[("selected", accent), ("active", colors["hover"])],
                       foreground=[("selected", "#ffffff"), ("active", fg)])
        self.style.configure("Treeview", background=field, fieldbackground=field, foreground=fg,
                             bordercolor=colors["border"], rowheight=22)
        self.style.map("Treeview", background=[("selected", colors["selected"])],
                       foreground=[("selected", fg)])
        self.style.configure("Treeview.Heading", background=surface, foreground=fg,
                             bordercolor=colors["border"], relief="flat")
        self.style.map("Treeview.Heading", background=[("active", colors["hover"])])
        self.style.configure("TScrollbar", background=surface, troughcolor=bg,
                             arrowcolor=fg, bordercolor=bg)
        self.style.configure("TProgressbar", background=accent, troughcolor=surface)
        self.style.configure("TSeparator", background=colors["border"])
        self.option_add("*TCombobox*Listbox*Background", field)
        self.option_add("*TCombobox*Listbox*Foreground", fg)
        self.option_add("*TCombobox*Listbox*SelectBackground", colors["selected"])
        self.option_add("*TCombobox*Listbox*SelectForeground", fg)
        link_color = "#8fc7ff" if dark else "#1f5a92"
        for label in self.attribution_labels:
            label.configure(foreground=link_color)
        for canvas_name, points_name, empty_text in (
            ("value_history_canvas", "value_history_points", "Select a player to load a 90-day trend chart."),
            ("roster_history_canvas", "roster_history_points", "Select a player to load a 90-day trend."),
            ("weekly_history_canvas", "weekly_history_points", "Select a player to load a 90-day trend."),
        ):
            canvas = getattr(self, canvas_name, None)
            if canvas:
                canvas.configure(bg=surface, highlightbackground=colors["border"])
                self._draw_history_chart(canvas, getattr(self, points_name, []), empty_text)
        if hasattr(self, "trade_history_canvas"):
            self._draw_trade_history()
        if hasattr(self, "trade_tab_canvas"):
            self.trade_tab_canvas.configure(bg=bg)

    def _toggle_dark_mode(self):
        #Apply the chosen palette and save it for the next launch.
        self._apply_theme()
        self._save_settings()

    @staticmethod
    def _app_version_key(version):
        #Compare the numeric parts of a version tag while accepting an optional leading v.
        match = re.fullmatch(r"v?(\d+)\.(\d+)(?:\.(\d+))?", str(version).strip())
        if not match:
            raise ValueError(f"Unsupported version number: {version}")
        return tuple(int(part or 0) for part in match.groups())

    @staticmethod
    def _plain_release_notes(body):
        #Trim common Markdown formatting so release notes read cleanly in the compact menu.
        text = str(body or "").strip()
        text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
        text = re.sub(r"^\s{0,3}#{1,6}\s*", "", text, flags=re.MULTILINE)
        text = re.sub(r"[*`_]", "", text)
        return text[:2400] or "No release notes were provided."

    def _fetch_latest_app_release(self):
        #Read the latest stable GitHub Release and find its Windows update package.
        response = requests.get(
            GITHUB_RELEASE_API,
            headers={"Accept": "application/vnd.github+json", "User-Agent": "FantasyTradeCalculator"},
            timeout=(8, 20),
        )
        if response.status_code == 404:
            return {"not_published": True}
        response.raise_for_status()
        release = response.json()
        version = str(release.get("tag_name", "")).strip()
        version_key = self._app_version_key(version)
        asset = next((item for item in release.get("assets", [])
                      if item.get("name") == UPDATE_ASSET_NAME), None)
        return {
            "version": version,
            "version_key": version_key,
            "url": str(asset.get("browser_download_url", "")) if asset else "",
            "digest": str(asset.get("digest", "")) if asset else "",
            "size": int(asset.get("size", 0)) if asset else 0,
            "notes": self._plain_release_notes(release.get("body", "")),
            "release_url": str(release.get("html_url", "")),
        }

    def _check_for_app_updates(self):
        #Check GitHub in the background so a slow connection never freezes the app.
        self.update_check_button.configure(state="disabled")
        self.update_install_button.configure(state="disabled")
        self.available_update = None
        self.update_status_var.set("Checking GitHub for the latest release…")
        self.update_notes_var.set("")
        self.run_background(self._fetch_latest_app_release_safely, self._show_app_update_result)

    def _fetch_latest_app_release_safely(self):
        #Return network errors to the update panel without leaving its controls disabled.
        try:
            return self._fetch_latest_app_release()
        except Exception as exc:
            return {"error": str(exc)}

    def _show_app_update_result(self, result):
        #Explain the release status and enable installation only for a verified Windows package.
        self.update_check_button.configure(state="normal")
        if result.get("error"):
            self.update_status_var.set("Could not check for updates.")
            self.update_notes_var.set(result["error"])
            return
        if result.get("not_published"):
            self.update_status_var.set(f"Installed version {APP_VERSION} · No release has been published yet.")
            self.update_notes_var.set("When a new version is released, its changes will appear here.")
            return
        self.update_notes_var.set(result["notes"])
        try:
            current_version = self._app_version_key(APP_VERSION)
        except ValueError as exc:
            self.update_status_var.set(str(exc))
            return
        if result["version_key"] <= current_version:
            self.update_status_var.set(f"You're up to date · version {APP_VERSION}.")
            return
        self.available_update = result
        if not result["url"] or not result["digest"].startswith("sha256:"):
            self.update_status_var.set(f"Version {result['version']} is available, but its verified Windows package is missing.")
            return
        if not getattr(sys, "frozen", False) or os.name != "nt":
            self.update_status_var.set(
                f"Version {result['version']} is available. Install updates from the packaged Windows app."
            )
            return
        self.update_status_var.set(f"Version {result['version']} is ready to install.")
        self.update_install_button.configure(state="normal")

    def _download_app_update(self):
        #Download and verify the release package before asking to close the app.
        release = self.available_update
        if not release:
            return
        self.update_install_button.configure(state="disabled")
        self.update_check_button.configure(state="disabled")
        self.update_status_var.set(f"Downloading version {release['version']}…")
        self.run_background(lambda: self._download_verified_update(release), self._app_update_downloaded)

    @staticmethod
    def _download_verified_update(release):
        #Accept only the project's HTTPS release asset and verify its published SHA-256 digest.
        url = release["url"]
        prefix = "https://github.com/jasonBuras/SleeperFantasyCalculator/releases/download/"
        if not url.startswith(prefix):
            return {"error": "The release package link is not from the project's GitHub Releases page."}
        if release["size"] <= 0 or release["size"] > 300 * 1024 * 1024:
            return {"error": "The release package has an unexpected file size."}
        update_dir = Path(tempfile.mkdtemp(prefix="fantasy-trade-update-"))
        archive_path = update_dir / UPDATE_ASSET_NAME
        digest = hashlib.sha256()
        total = 0
        try:
            with requests.get(url, stream=True, timeout=(10, 60), headers={"User-Agent": "FantasyTradeCalculator"}) as response:
                response.raise_for_status()
                with archive_path.open("wb") as archive_file:
                    for chunk in response.iter_content(chunk_size=1024 * 1024):
                        if not chunk:
                            continue
                        total += len(chunk)
                        if total > 300 * 1024 * 1024:
                            raise ValueError("The downloaded update is larger than expected.")
                        digest.update(chunk)
                        archive_file.write(chunk)
            if total != release["size"]:
                raise ValueError("The downloaded update size does not match the release details.")
            expected_digest = release["digest"].split(":", 1)[1].casefold()
            if not re.fullmatch(r"[0-9a-f]{64}", expected_digest) or digest.hexdigest() != expected_digest:
                raise ValueError("The update failed its SHA-256 integrity check. Nothing was installed.")
            with zipfile.ZipFile(archive_path) as package:
                files = [item for item in package.infolist() if not item.is_dir()]
                if len(files) != 1 or files[0].filename != "FantasyTradeCalculator.exe":
                    raise ValueError("The update package does not contain the expected app file.")
                if files[0].file_size < 1_000_000 or package.testzip() is not None:
                    raise ValueError("The downloaded update package is incomplete or damaged.")
            return {"archive": str(archive_path), "helper_dir": str(update_dir)}
        except Exception as exc:
            try:
                archive_path.unlink(missing_ok=True)
                update_dir.rmdir()
            except OSError:
                pass
            return {"error": str(exc)}

    def _app_update_downloaded(self, result):
        #Ask before closing, then hand the verified archive to the post-exit updater.
        if result.get("error"):
            self.update_check_button.configure(state="normal")
            self.update_install_button.configure(state="normal" if self.available_update else "disabled")
            self.update_status_var.set("The update could not be downloaded.")
            messagebox.showerror("Update Download Failed", result["error"])
            return
        release = self.available_update
        if not messagebox.askyesno(
            "Install Update",
            f"Version {release['version']} is downloaded and verified.\n\n"
            "The app will close, replace its program file, and reopen. Your saved leagues and trades are kept separately. "
            "The app folder must allow file changes. Continue?",
        ):
            shutil.rmtree(result["helper_dir"], ignore_errors=True)
            self.update_check_button.configure(state="normal")
            self.update_install_button.configure(state="normal")
            self.update_status_var.set(f"Version {release['version']} is ready when you are.")
            return
        executable = Path(sys.executable).resolve()
        if executable.name.casefold() != "fantasytradecalculator.exe":
            shutil.rmtree(result["helper_dir"], ignore_errors=True)
            self.update_check_button.configure(state="normal")
            self.update_install_button.configure(state="normal")
            messagebox.showerror("Cannot Install Update", "This app file has been renamed. Restore the name FantasyTradeCalculator.exe and try again.")
            return
        try:
            config_path = Path(result["helper_dir"]) / "update.json"
            script_path = Path(result["helper_dir"]) / "install_update.ps1"
            script_path.write_text(resource_path("install_update.ps1").read_text(encoding="utf-8"), encoding="utf-8")
            config_path.write_text(json.dumps({
                "process_id": os.getpid(),
                "executable": str(executable),
                "archive": result["archive"],
                "helper_script": str(script_path),
            }), encoding="utf-8")
            subprocess.Popen(
                ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden",
                 "-File", str(script_path), "-ConfigPath", str(config_path)],
                cwd=str(result["helper_dir"]), creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception as exc:
            shutil.rmtree(result["helper_dir"], ignore_errors=True)
            self.update_check_button.configure(state="normal")
            self.update_install_button.configure(state="normal")
            messagebox.showerror("Cannot Start Update", str(exc))
            return
        self.update_status_var.set("Closing the app to install the update…")
        self.destroy()

    def _build_ui(self):
        #Create the app header, league controls, work tabs, and bottom status indicator.
        main = ttk.Frame(self, padding=10)
        main.pack(fill="both", expand=True)

        header = ttk.Frame(main)
        header.pack(fill="x", pady=(0, 8))
        ttk.Label(
            header,
            text="Fantasy Trade Calculator",
            font=("TkDefaultFont", 18, "bold")
        ).pack(side="left", anchor="w")
        ttk.Label(
            header, textvariable=self.window_context_var,
            font=("TkDefaultFont", 10),
        ).pack(side="left", anchor="w", padx=(14, 0), fill="x", expand=True)
        ttk.Checkbutton(
            header, text="Dark mode", variable=self.dark_mode_var,
            command=self._toggle_dark_mode
        ).pack(side="right", padx=5)

        #The Main Menu contains league selection and actions; the active context stays in the header.
        self.workspace = ttk.Notebook(main)
        self.workspace.pack(fill="both", expand=True)
        main_menu_tab = ttk.Frame(self.workspace, padding=12)
        self.workspace.add(main_menu_tab, text="Main Menu")
        ttk.Label(
            main_menu_tab,
            text="Choose the Sleeper account, season, and league to open. Your active league is shown at the top of every page.",
            wraplength=1050,
        ).pack(anchor="w", pady=(0, 10))

        form = ttk.LabelFrame(main_menu_tab, text="League Selection", padding=9)
        form.pack(fill="x")

        ttk.Label(form, text="Sleeper Username:").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.username_var, width=35).grid(
            row=0, column=1, sticky="ew", padx=10, pady=6
        )

        ttk.Label(form, text="NFL Season:").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Combobox(
            form,
            textvariable=self.season_var,
            values=[str(y) for y in range(datetime.now().year + 1, 2017, -1)],
            width=10,
            state="readonly",
        ).grid(row=1, column=1, sticky="w", padx=10, pady=6)

        ttk.Button(
            form,
            text="Load Leagues",
            command=self.load_leagues
        ).grid(row=0, column=2, rowspan=2, padx=10)

        ttk.Label(form, text="League:").grid(row=2, column=0, sticky="w", pady=6)
        self.league_combo = ttk.Combobox(
            form,
            textvariable=self.league_var,
            width=50,
            state="readonly"
        )
        self.league_combo.grid(row=2, column=1, columnspan=2, sticky="ew", padx=10, pady=6)
        self.league_combo.bind("<<ComboboxSelected>>", self._save_selected_league)

        form.columnconfigure(1, weight=1)

        #Keep loading and export actions together above the analysis tabs.
        actions = ttk.Frame(main_menu_tab)
        actions.pack(fill="x", pady=(12, 0))

        ttk.Button(
            actions,
            text="Load Selected League",
            command=self.load_selected_league
        ).pack(side="left")

        ttk.Button(
            actions,
            text="Refresh Player Data",
            command=self.refresh_players
        ).pack(side="left", padx=10)

        self.export_button = ttk.Button(
            actions,
            text="Export to Excel",
            command=self.export,
            state="disabled"
        )
        self.export_button.pack(side="right")

        updates = ttk.LabelFrame(main_menu_tab, text="Software Updates", padding=9)
        updates.pack(fill="x", pady=(12, 0))
        update_actions = ttk.Frame(updates)
        update_actions.pack(fill="x")
        self.update_check_button = ttk.Button(
            update_actions, text="Check for Updates", command=self._check_for_app_updates,
        )
        self.update_check_button.pack(side="left")
        self.update_install_button = ttk.Button(
            update_actions, text="Download and Install", command=self._download_app_update,
            state="disabled",
        )
        self.update_install_button.pack(side="left", padx=(8, 10))
        ttk.Label(update_actions, textvariable=self.update_status_var).pack(side="left", fill="x", expand=True)
        ttk.Label(
            updates, textvariable=self.update_notes_var, wraplength=1050, justify="left",
        ).pack(anchor="w", fill="x", pady=(7, 0))

        #Position Needs compares one team's roster counts with the league lineup slots.
        needs_tab = ttk.Frame(self.workspace, padding=12)
        self.position_needs_tab = needs_tab
        self.workspace.add(needs_tab, text="Position Needs")
        ttk.Label(needs_tab, text="Review roster depth against your league's starting lineup.").pack(anchor="w", pady=(0, 8))
        needs_controls = ttk.Frame(needs_tab)
        needs_controls.pack(fill="x", pady=(0, 8))
        ttk.Label(needs_controls, text="Team:").pack(side="left")
        self.needs_team_combo = ttk.Combobox(needs_controls, state="readonly", width=36)
        self.needs_team_combo.pack(side="left", padx=8)
        self.needs_team_combo.bind("<<ComboboxSelected>>", self._show_position_needs)
        self.needs_tree = ttk.Treeview(needs_tab, columns=("required", "starting", "rostered", "bench", "assessment"), show="tree headings", height=8)
        self.needs_tree.heading("#0", text="Position")
        for key, label in (("required", "Starter Spots"), ("starting", "Starting"), ("rostered", "Rostered"), ("bench", "Bench Depth"), ("assessment", "Overview")):
            self.needs_tree.heading(key, text=label)
        self.needs_tree.column("#0", width=125)
        for key in ("required", "starting", "rostered", "bench"):
            self.needs_tree.column(key, width=95, anchor="center")
        self.needs_tree.column("assessment", width=190)
        self.needs_tree.pack(fill="both", expand=True)
        self.needs_summary_var = tk.StringVar(value="Load a league to see position depth.")
        ttk.Label(needs_tab, textvariable=self.needs_summary_var, wraplength=850).pack(anchor="w", pady=(8, 0))

        #Roster View groups a selected team's players by their Sleeper lineup assignment.
        roster_tab = ttk.Frame(self.workspace, padding=12)
        self.workspace.add(roster_tab, text="Roster View")
        ttk.Label(roster_tab, text="See each player's lineup slot, availability, bye week, and weekly projection.",
                  font=("TkDefaultFont", 11, "bold")).pack(anchor="w", pady=(0, 7))
        roster_controls = ttk.Frame(roster_tab)
        roster_controls.pack(fill="x", pady=(0, 6))
        ttk.Label(roster_controls, text="Team:").pack(side="left")
        self.roster_team_combo = ttk.Combobox(roster_controls, state="readonly", width=38)
        self.roster_team_combo.pack(side="left", padx=8)
        self.roster_team_combo.bind("<<ComboboxSelected>>", self._refresh_roster_view)
        self.roster_projection_note = tk.StringVar(value="Weekly projections appear when Sleeper publishes them.")
        ttk.Label(roster_controls, textvariable=self.roster_projection_note).pack(side="left", padx=(8, 0))
        roster_search_row = ttk.Frame(roster_tab)
        roster_search_row.pack(fill="x", pady=(0, 5))
        ttk.Label(roster_search_row, text="Search Player:").pack(side="left")
        self.roster_search_var = tk.StringVar()
        ttk.Entry(roster_search_row, textvariable=self.roster_search_var, width=32).pack(
            side="left", padx=(8, 0)
        )
        self.roster_search_status = tk.StringVar(value="Search by player, position, team, or status.")
        ttk.Label(roster_search_row, textvariable=self.roster_search_status).pack(side="left", padx=10)
        self.roster_search_var.trace_add("write", lambda *_: self._apply_roster_search())
        roster_list = ttk.Frame(roster_tab)
        roster_list.pack(fill="both", expand=True)
        self.roster_tree = ttk.Treeview(
            roster_list, columns=("slot", "position", "nfl_team", "status", "bye", "projection"),
            show="tree headings", selectmode="browse", height=8,
        )
        self.roster_tree.heading("#0", text="Name")
        for key, label in (("slot", "Slot"), ("position", "Position"), ("nfl_team", "NFL Team"),
                           ("status", "Status"), ("bye", "Bye"), ("projection", "Proj Points")):
            self.roster_tree.heading(key, text=label)
        self.roster_tree.column("#0", width=240)
        self.roster_tree.column("slot", width=115, anchor="center")
        self.roster_tree.column("position", width=100, anchor="center")
        self.roster_tree.column("nfl_team", width=90, anchor="center")
        self.roster_tree.column("status", width=130, anchor="center")
        self.roster_tree.column("bye", width=70, anchor="center")
        self.roster_tree.column("projection", width=110, anchor="e")
        roster_scroll = ttk.Scrollbar(roster_list, orient="vertical", command=self.roster_tree.yview)
        self.roster_tree.configure(yscrollcommand=roster_scroll.set)
        self.roster_tree.pack(side="left", fill="both", expand=True)
        roster_scroll.pack(side="right", fill="y")
        self.roster_tree.bind("<<TreeviewSelect>>", self._select_roster_history)
        self.roster_history_detail = tk.StringVar(value="Select a player to see recent value movement and volatility.")
        roster_history_credit = ttk.Label(
            roster_tab, text="Value history provided by Stats Guy Fantasy · statsguyfantasy.com",
            foreground="#1f5a92", cursor="hand2", font=("TkDefaultFont", 9, "bold"),
        )
        self.attribution_labels.append(roster_history_credit)
        roster_history_credit.pack(anchor="w", pady=(5, 0))
        roster_history_credit.bind("<Button-1>", lambda _event: __import__("webbrowser").open("https://statsguyfantasy.com/"))
        roster_history_panel = ttk.LabelFrame(roster_tab, text="Player Value History", padding=7)
        roster_history_panel.pack(fill="x", pady=(6, 0))
        self._build_history_window_control(roster_history_panel, action=lambda: self._open_player_history_graph("roster"))
        ttk.Label(roster_history_panel, textvariable=self.roster_history_detail, wraplength=1050,
                  justify="left").pack(anchor="w")
        self.roster_history_canvas = tk.Canvas(
            roster_history_panel, height=72, highlightthickness=0,
            bg=self.theme_colors["surface"], highlightbackground=self.theme_colors["border"],
        )
        self.roster_history_canvas.pack(fill="x", pady=(4, 0))
        self.roster_history_canvas.bind("<Configure>", lambda _event: self._draw_history_chart(
            self.roster_history_canvas, self.roster_history_points,
            "Select a player to load a 90-day value trend.",
        ))

        #Weekly Help brings matchup context, bye coverage, and waiver options together.
        weekly_tab = ttk.Frame(self.workspace, padding=12)
        self.workspace.add(weekly_tab, text="Weekly Help")
        ttk.Label(weekly_tab, text="Plan this week around your matchup, byes, and available players.",
                  font=("TkDefaultFont", 11, "bold")).pack(anchor="w", pady=(0, 7))
        weekly_controls = ttk.Frame(weekly_tab)
        weekly_controls.pack(fill="x", pady=(0, 8))
        ttk.Label(weekly_controls, text="Your team:").pack(side="left")
        self.weekly_team_combo = ttk.Combobox(weekly_controls, state="readonly", width=36)
        self.weekly_team_combo.pack(side="left", padx=(8, 14))
        self.weekly_team_combo.bind("<<ComboboxSelected>>", self._refresh_weekly_help)
        ttk.Label(weekly_controls, text="Waiver focus:").pack(side="left")
        self.weekly_filter_combo = ttk.Combobox(
            weekly_controls, state="readonly", width=20,
            values=("Bye Coverage", "All Positions", "QB", "RB", "WR", "TE", "K", "DEF"),
        )
        self.weekly_filter_combo.current(0)
        self.weekly_filter_combo.pack(side="left", padx=8)
        self.weekly_filter_combo.bind("<<ComboboxSelected>>", self._refresh_weekly_waivers)
        ttk.Button(weekly_controls, text="Refresh Weekly Data", command=self._load_weekly_context).pack(side="right")
        weekly_summary = ttk.LabelFrame(weekly_tab, text="Matchup and Bye Coverage", padding=9)
        weekly_summary.pack(fill="x", pady=(0, 8))
        self.weekly_summary_var = tk.StringVar(value="Load a league to see this week's matchup and bye coverage.")
        ttk.Label(weekly_summary, textvariable=self.weekly_summary_var, wraplength=1050,
                  justify="left").pack(anchor="w")
        weekly_credit = ttk.Label(
            weekly_tab, text="FantasyCalc market values · fantasycalc.com (not weekly point projections)",
            foreground="#1f5a92", cursor="hand2", font=("TkDefaultFont", 9, "bold"),
        )
        self.attribution_labels.append(weekly_credit)
        weekly_credit.pack(anchor="w", pady=(0, 6))
        weekly_credit.bind("<Button-1>", lambda _event: __import__("webbrowser").open("https://fantasycalc.com/"))
        schedule_credit = ttk.Label(
            weekly_tab, text="NFL schedule and bye weeks: nflverse / Lee Sharpe · CC BY 4.0",
            foreground="#1f5a92", cursor="hand2",
        )
        self.attribution_labels.append(schedule_credit)
        schedule_credit.pack(anchor="w", pady=(0, 6))
        schedule_credit.bind("<Button-1>", lambda _event: __import__("webbrowser").open("https://nflreadr.nflverse.com/reference/load_schedules.html"))
        ttk.Label(weekly_tab, text="Available players", font=("TkDefaultFont", 10, "bold")).pack(anchor="w")
        weekly_list = ttk.Frame(weekly_tab)
        weekly_list.pack(fill="both", expand=True, pady=(5, 0))
        self.weekly_tree = ttk.Treeview(
            weekly_list, columns=("position", "nfl_team", "status", "bye", "value"),
            show="tree headings", selectmode="browse", height=8,
        )
        self.weekly_tree.heading("#0", text="Player", command=lambda: self._sort_weekly_waivers("name"))
        self.weekly_tree.heading("position", text="Position", command=lambda: self._sort_weekly_waivers("position"))
        self.weekly_tree.heading("nfl_team", text="NFL Team", command=lambda: self._sort_weekly_waivers("nfl_team"))
        self.weekly_tree.heading("status", text="Status", command=lambda: self._sort_weekly_waivers("status"))
        self.weekly_tree.heading("bye", text="Bye Week", command=lambda: self._sort_weekly_waivers("bye"))
        self.weekly_tree.heading("value", text="FC Value", command=lambda: self._sort_weekly_waivers("value"))
        self.weekly_tree.column("#0", width=240)
        self.weekly_tree.column("position", width=110, anchor="center")
        self.weekly_tree.column("nfl_team", width=100, anchor="center")
        self.weekly_tree.column("status", width=130, anchor="center")
        self.weekly_tree.column("bye", width=90, anchor="center")
        self.weekly_tree.column("value", width=100, anchor="e")
        weekly_scroll = ttk.Scrollbar(weekly_list, orient="vertical", command=self.weekly_tree.yview)
        self.weekly_tree.configure(yscrollcommand=weekly_scroll.set)
        self.weekly_tree.pack(side="left", fill="both", expand=True)
        weekly_scroll.pack(side="right", fill="y")
        self.weekly_waiver_count = tk.StringVar(value="Load a league to view free agents.")
        ttk.Label(weekly_tab, textvariable=self.weekly_waiver_count).pack(anchor="w", pady=(5, 0))
        self.weekly_history_detail = tk.StringVar(value="Select a waiver option to see its recent value movement.")
        weekly_history_credit = ttk.Label(
            weekly_tab, text="Player value history provided by Stats Guy Fantasy · statsguyfantasy.com",
            foreground="#1f5a92", cursor="hand2", font=("TkDefaultFont", 9, "bold"),
        )
        self.attribution_labels.append(weekly_history_credit)
        weekly_history_credit.pack(anchor="w", pady=(5, 0))
        weekly_history_credit.bind("<Button-1>", lambda _event: __import__("webbrowser").open("https://statsguyfantasy.com/"))
        weekly_history_panel = ttk.LabelFrame(weekly_tab, text="Selected Player Value History", padding=7)
        weekly_history_panel.pack(fill="x", pady=(5, 0))
        self._build_history_window_control(weekly_history_panel, action=lambda: self._open_player_history_graph("weekly"))
        ttk.Label(weekly_history_panel, textvariable=self.weekly_history_detail, wraplength=1050,
                  justify="left").pack(anchor="w")
        self.weekly_history_canvas = tk.Canvas(
            weekly_history_panel, height=68, highlightthickness=0,
            bg=self.theme_colors["surface"], highlightbackground=self.theme_colors["border"],
        )
        self.weekly_history_canvas.pack(fill="x", pady=(4, 0))
        self.weekly_history_canvas.bind("<Configure>", lambda _event: self._draw_history_chart(
            self.weekly_history_canvas, self.weekly_history_points,
            "Select a player to load a 90-day value trend.",
        ))
        self.weekly_tree.bind("<<TreeviewSelect>>", self._select_weekly_history)

        #Value Trends compares independent trade-value estimates and displays historical movement.
        trends_tab = ttk.Frame(self.workspace, padding=12)
        self.workspace.add(trends_tab, text="Value Trends")
        ttk.Label(trends_tab, text="Compare market estimates and review recent value movement.",
                  font=("TkDefaultFont", 11, "bold")).pack(anchor="w", pady=(0, 7))
        trends_controls = ttk.Frame(trends_tab)
        trends_controls.pack(fill="x", pady=(0, 6))
        ttk.Label(trends_controls, text="Players:").pack(side="left")
        self.value_trend_team_combo = ttk.Combobox(trends_controls, state="readonly", width=38)
        self.value_trend_team_combo.pack(side="left", padx=(8, 14))
        self.value_trend_team_combo.bind("<<ComboboxSelected>>", self._refresh_value_trends)
        self.value_trend_source_status = tk.StringVar(value="Load a league to compare value sources.")
        ttk.Label(trends_controls, textvariable=self.value_trend_source_status).pack(side="left", fill="x", expand=True)
        stats_guy_credit = ttk.Label(
            trends_tab, text="Stats Guy Fantasy values and value history · statsguyfantasy.com",
            foreground="#1f5a92", cursor="hand2", font=("TkDefaultFont", 9, "bold"),
        )
        self.attribution_labels.append(stats_guy_credit)
        stats_guy_credit.pack(anchor="w", pady=(0, 3))
        stats_guy_credit.bind("<Button-1>", lambda _event: __import__("webbrowser").open("https://statsguyfantasy.com/"))
        trends_fantasycalc_credit = ttk.Label(
            trends_tab, text="FantasyCalc values · fantasycalc.com",
            foreground="#1f5a92", cursor="hand2", font=("TkDefaultFont", 9, "bold"),
        )
        self.attribution_labels.append(trends_fantasycalc_credit)
        trends_fantasycalc_credit.pack(anchor="w", pady=(0, 5))
        trends_fantasycalc_credit.bind("<Button-1>", lambda _event: __import__("webbrowser").open("https://fantasycalc.com/"))
        self.value_trend_format_note = tk.StringVar(
            value="Stats Guy uses its own trade-value model; PPR and TE-premium adjustments are not available."
        )
        ttk.Label(trends_tab, textvariable=self.value_trend_format_note, wraplength=1050).pack(anchor="w", pady=(0, 6))
        trend_list = ttk.Frame(trends_tab)
        trend_list.pack(fill="both", expand=True)
        self.value_trend_tree = ttk.Treeview(
            trend_list, columns=("position", "rostered_by", "fc", "fc30", "stats_guy", "spread"),
            show="tree headings", selectmode="browse", height=8,
        )
        self.value_trend_tree.heading("#0", text="Player", command=lambda: self._sort_value_trends("name"))
        self.value_trend_tree.heading("position", text="Position", command=lambda: self._sort_value_trends("position"))
        self.value_trend_tree.heading("rostered_by", text="Rostered By", command=lambda: self._sort_value_trends("rostered_by"))
        self.value_trend_tree.heading("fc", text="FantasyCalc", command=lambda: self._sort_value_trends("fc"))
        self.value_trend_tree.heading("fc30", text="FC 30d Δ", command=lambda: self._sort_value_trends("fc30"))
        self.value_trend_tree.heading("stats_guy", text="Stats Guy", command=lambda: self._sort_value_trends("stats_guy"))
        self.value_trend_tree.heading("spread", text="FC − SG", command=lambda: self._sort_value_trends("spread"))
        self.value_trend_tree.column("#0", width=230)
        self.value_trend_tree.column("position", width=85, anchor="center")
        self.value_trend_tree.column("rostered_by", width=220)
        for column in ("fc", "fc30", "stats_guy", "spread"):
            self.value_trend_tree.column(column, width=100, anchor="e")
        trend_scroll = ttk.Scrollbar(trend_list, orient="vertical", command=self.value_trend_tree.yview)
        self.value_trend_tree.configure(yscrollcommand=trend_scroll.set)
        self.value_trend_tree.pack(side="left", fill="both", expand=True)
        trend_scroll.pack(side="right", fill="y")
        self.value_trend_tree.bind("<<TreeviewSelect>>", self._select_value_trend_player)
        self.value_trend_detail = tk.StringVar(value="Select a player to see volatility and the recent value history.")
        history_panel = ttk.LabelFrame(trends_tab, text="Selected Player: History and Volatility", padding=8)
        history_panel.pack(fill="x", pady=(7, 0))
        self._build_history_window_control(history_panel, action=lambda: self._open_player_history_graph("trends"))
        ttk.Label(history_panel, textvariable=self.value_trend_detail, wraplength=1050,
                  justify="left").pack(anchor="w")
        self.value_history_canvas = tk.Canvas(
            history_panel, height=82, highlightthickness=0,
            bg=self.theme_colors["surface"], highlightbackground=self.theme_colors["border"],
        )
        self.value_history_canvas.pack(fill="x", pady=(5, 0))
        self.value_history_canvas.bind("<Configure>", lambda _event: self._draw_value_history())
        self.value_history_points = []

        #Trade Review scrolls as one page when the window cannot show every section at once.
        self.trade_tab = ttk.Frame(self.workspace)
        self.workspace.add(self.trade_tab, text="Trade Review")
        self.trade_tab_canvas = tk.Canvas(
            self.trade_tab, highlightthickness=0, borderwidth=0,
            bg=self.theme_colors["background"],
        )
        trade_tab_scroll = ttk.Scrollbar(
            self.trade_tab, orient="vertical", command=self.trade_tab_canvas.yview
        )
        self.trade_tab_canvas.configure(yscrollcommand=trade_tab_scroll.set)
        self.trade_tab_canvas.pack(side="left", fill="both", expand=True)
        trade_tab_scroll.pack(side="right", fill="y")
        trade_tab = ttk.Frame(self.trade_tab_canvas, padding=8)
        trade_tab_window = self.trade_tab_canvas.create_window(
            (0, 0), window=trade_tab, anchor="nw"
        )
        trade_tab.bind(
            "<Configure>",
            lambda _event: self.trade_tab_canvas.configure(
                scrollregion=self.trade_tab_canvas.bbox("all")
            ),
        )
        self.trade_tab_canvas.bind(
            "<Configure>",
            lambda event: self.trade_tab_canvas.itemconfigure(trade_tab_window, width=event.width),
        )
        self.bind_all("<MouseWheel>", self._scroll_trade_tab, add="+")
        ttk.Label(trade_tab, text="Select the players on each side to compare market values.", font=("TkDefaultFont", 11, "bold")).pack(anchor="w")
        #ttk.Label(trade_tab, text="Click a column title to sort. Names sort by last name; positions follow QB, RB, WR, TE, K.", wraplength=850).pack(anchor="w", pady=(2, 3))
        #ttk.Label(trade_tab, text="Market estimate from FantasyCalc. Roster fit and projections are not included.", wraplength=850).pack(anchor="w", pady=(3, 5))
        credit = ttk.Label(trade_tab, text="Trade values provided by FantasyCalc · fantasycalc.com", foreground="#1f5a92", cursor="hand2")
        self.attribution_labels.append(credit)
        credit.pack(anchor="w", pady=(0, 5))
        credit.bind("<Button-1>", lambda _event: __import__("webbrowser").open("https://fantasycalc.com/"))
        stats_trade_credit = ttk.Label(
            trade_tab, text="Stats Guy Fantasy · statsguyfantasy.com (no PPR, TE-premium, or team-count adjustment)",
            foreground="#1f5a92", cursor="hand2",
        )
        self.attribution_labels.append(stats_trade_credit)
        stats_trade_credit.pack(anchor="w", pady=(0, 5))
        stats_trade_credit.bind("<Button-1>", lambda _event: __import__("webbrowser").open("https://statsguyfantasy.com/"))
        #ttk.Label(trade_tab, text="FantasyCalc values are a market signal, not a guarantee. A 10% value gap is treated as fairly close. This tool does not submit trades.", wraplength=850).pack(anchor="w", pady=(0, 8))

        format_row = ttk.Frame(trade_tab)
        format_row.pack(fill="x", pady=(0, 8))
        ttk.Label(format_row, text="League format:").pack(side="left")
        self.format_var = tk.StringVar(value=self.saved_settings.get("format", "Dynasty"))
        self.format_combo = ttk.Combobox(format_row, textvariable=self.format_var, values=("Dynasty", "Redraft"), state="readonly", width=12)
        self.format_combo.pack(side="left", padx=(8, 14))
        self.format_combo.bind("<<ComboboxSelected>>", self._format_changed)
        self.market_settings_var = tk.StringVar(value="Load a league to set market-value settings.")
        ttk.Label(format_row, textvariable=self.market_settings_var).pack(side="left")

        sides = ttk.Frame(trade_tab)
        sides.pack(fill="both", expand=True)
        sides.columnconfigure(0, weight=1)
        sides.columnconfigure(1, weight=1)
        left = ttk.LabelFrame(sides, text="You give", padding=8)
        right = ttk.LabelFrame(sides, text="You receive", padding=8)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        left_team_row = ttk.Frame(left)
        right_team_row = ttk.Frame(right)
        left_team_row.pack(fill="x", pady=(0, 7))
        right_team_row.pack(fill="x", pady=(0, 7))
        self.trade_left_team = ttk.Combobox(left_team_row, state="readonly")
        self.trade_right_team = ttk.Combobox(right_team_row, state="readonly")
        self.trade_left_team.pack(side="left", fill="x", expand=True)
        self.trade_right_team.pack(side="left", fill="x", expand=True)
        self.trade_left_team.bind("<<ComboboxSelected>>", self._populate_trade_rosters)
        self.trade_right_team.bind("<<ComboboxSelected>>", self._populate_trade_rosters)
        ttk.Button(left_team_row, text="Clear Selection", command=lambda: self._clear_trade_selection(self.trade_left_tree)).pack(side="left", padx=(8, 0))
        ttk.Button(right_team_row, text="Clear Selection", command=lambda: self._clear_trade_selection(self.trade_right_tree)).pack(side="left", padx=(8, 0))
        left_roster_list = ttk.Frame(left)
        right_roster_list = ttk.Frame(right)
        left_roster_list.pack(fill="both", expand=True)
        right_roster_list.pack(fill="both", expand=True)
        self.trade_left_tree = self._make_trade_tree(left_roster_list)
        self.trade_right_tree = self._make_trade_tree(right_roster_list)
        left_scroll = ttk.Scrollbar(left_roster_list, orient="vertical", command=self.trade_left_tree.yview)
        right_scroll = ttk.Scrollbar(right_roster_list, orient="vertical", command=self.trade_right_tree.yview)
        self.trade_left_tree.configure(yscrollcommand=left_scroll.set)
        self.trade_right_tree.configure(yscrollcommand=right_scroll.set)
        self.trade_left_tree.pack(side="left", fill="both", expand=True)
        left_scroll.pack(side="right", fill="y")
        self.trade_right_tree.pack(side="left", fill="both", expand=True)
        right_scroll.pack(side="right", fill="y")
        ttk.Label(left, text="Hold Ctrl to select multiple players.").pack(anchor="w", pady=(5, 0))
        ttk.Label(right, text="Hold Ctrl to select multiple players.").pack(anchor="w", pady=(5, 0))
        self.trade_left_tree.bind("<<TreeviewSelect>>", self._update_trade_review)
        self.trade_right_tree.bind("<<TreeviewSelect>>", self._update_trade_review)
        copy_preview = ttk.LabelFrame(trade_tab, text="Trade Overview", padding=5)
        copy_preview.pack(fill="x", pady=(5, 0))
        ttk.Label(copy_preview, textvariable=self.copy_trade_preview_var,
                  wraplength=1300, justify="left").pack(anchor="w")
        trade_actions = ttk.Frame(trade_tab)
        trade_actions.pack(fill="x", pady=(5, 0))
        self.save_trade_button = ttk.Button(trade_actions, text="Save Trade", command=self._save_current_trade)
        self.save_trade_button.pack(side="left")
        self.save_as_new_trade_button = ttk.Button(
            trade_actions, text="Save as New Trade",
            command=lambda: self._save_current_trade(save_as_new=True),
        )
        self.new_trade_button = ttk.Button(trade_actions, text="New Trade", command=self._new_trade)
        self.new_trade_button.pack(side="left", padx=8)
        self.copy_trade_button = ttk.Button(
            trade_actions, text="Copy Trade Overview Text", command=self._copy_trade_text, state="disabled"
        )
        self.copy_trade_button.pack(side="left")
        note_row = ttk.Frame(trade_tab)
        note_row.pack(fill="x", pady=(4, 0))
        ttk.Label(note_row, text="Note:").pack(side="left")
        self.trade_note_var = tk.StringVar()
        ttk.Entry(note_row, textvariable=self.trade_note_var).pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.trade_result_var = tk.StringVar(value="Load a league to review a trade.")
        ttk.Label(trade_tab, textvariable=self.trade_result_var, font=("TkDefaultFont", 12, "bold"), wraplength=850).pack(anchor="w", pady=(6, 0))
        balance_row = ttk.Frame(trade_tab)
        balance_row.pack(fill="x", pady=(3, 0))
        ttk.Label(balance_row, text="Trade balance:").pack(side="left")
        self.trade_balance_canvas = tk.Canvas(
            balance_row, width=260, height=18, highlightthickness=0,
            bg=self.theme_colors["background"],
        )
        self.trade_balance_canvas.pack(side="left", padx=(8, 10))
        self.trade_balance_summary = tk.StringVar(value="Select players on both sides.")
        ttk.Label(balance_row, textvariable=self.trade_balance_summary).pack(side="left")
        self.trade_stats_guy_result_var = tk.StringVar(value="Stats Guy trade values will appear here after loading.")
        ttk.Label(trade_tab, textvariable=self.trade_stats_guy_result_var, font=("TkDefaultFont", 10, "bold"),
                  wraplength=850).pack(anchor="w", pady=(3, 0))
        self.trade_impact_var = tk.StringVar(value="Roster and lineup impact details will appear here.")
        impact_row = ttk.Frame(trade_tab)
        impact_row.pack(fill="x", pady=(5, 0))
        self.trade_impact_summary = tk.StringVar(value="Roster impact appears when you select players.")
        ttk.Label(impact_row, textvariable=self.trade_impact_summary, wraplength=760).pack(side="left", anchor="w", fill="x", expand=True)
        ttk.Button(impact_row, text="View Lineup Impact", command=self._show_lineup_impact).pack(side="right", padx=(8, 0))
        trade_history_panel = ttk.LabelFrame(trade_tab, text="Selected Players' Value Trends", padding=4)
        trade_history_panel.pack(fill="x", pady=(4, 0))
        trade_history_toolbar = ttk.Frame(trade_history_panel)
        trade_history_toolbar.pack(fill="x")
        self._build_history_window_control(trade_history_toolbar, metric=True, action=self._open_trade_history_window)
        ttk.Label(trade_history_panel, text="Choose Percent Change or Value; hover over a line to see the date, value, and change.",
                  wraplength=1000).pack(anchor="w")
        self.trade_history_status = tk.StringVar(value="Select one or more players to compare their value trends.")
        ttk.Label(trade_history_panel, textvariable=self.trade_history_status, wraplength=1000).pack(anchor="w")
        self.trade_history_canvas = tk.Canvas(
            trade_history_panel, height=170 if self.winfo_screenheight() > 1200 else 96, highlightthickness=0,
            bg=self.theme_colors["surface"], highlightbackground=self.theme_colors["border"],
        )
        self.trade_history_canvas.pack(fill="x", pady=(3, 0))
        self.trade_history_canvas.bind("<Configure>", lambda _event: self._draw_trade_history())
        self.trade_history_canvas.bind(
            "<Motion>", lambda event: self._show_trade_history_hover(event, self.trade_history_canvas)
        )
        self.trade_history_canvas.bind("<Leave>", lambda _event: self._clear_trade_history_hover(self.trade_history_canvas))

        #Saved Trades keeps proposals available for reopening and editing later.
        self.saved_trade_tab = ttk.Frame(self.workspace, padding=12)
        self.workspace.add(self.saved_trade_tab, text="Saved Trades")
        ttk.Label(self.saved_trade_tab, text="Saved trades refresh their totals whenever current FantasyCalc values load.", wraplength=850).pack(anchor="w", pady=(0, 4))
        saved_credit = ttk.Label(self.saved_trade_tab, text="Trade values provided by FantasyCalc · fantasycalc.com", foreground="#1f5a92", cursor="hand2")
        self.attribution_labels.append(saved_credit)
        saved_credit.pack(anchor="w", pady=(0, 8))
        saved_credit.bind("<Button-1>", lambda _event: __import__("webbrowser").open("https://fantasycalc.com/"))
        self.saved_trade_tree = ttk.Treeview(
            self.saved_trade_tab,
            columns=("outcome", "give", "receive", "lean", "note", "updated"),
            show="tree headings", selectmode="browse", height=7,
        )
        self.saved_trade_tree.heading("#0", text="Trade")
        for key, label in (("outcome", "Outcome"), ("give", "You Give"), ("receive", "You Receive"),
                           ("lean", "Market Lean"), ("note", "Note"), ("updated", "Last Saved")):
            self.saved_trade_tree.heading(key, text=label)
        self.saved_trade_tree.column("#0", width=240)
        self.saved_trade_tree.column("outcome", width=105, anchor="center")
        self.saved_trade_tree.column("give", width=115, anchor="e")
        self.saved_trade_tree.column("receive", width=115, anchor="e")
        self.saved_trade_tree.column("lean", width=180)
        self.saved_trade_tree.column("note", width=180)
        self.saved_trade_tree.column("updated", width=150)
        self.saved_trade_tree.pack(fill="both", expand=True)
        self.saved_trade_tree.bind("<<TreeviewSelect>>", self._show_saved_trade_details)
        self.saved_trade_details = tk.StringVar(value="Select a saved trade to see its players.")
        ttk.Label(self.saved_trade_tab, textvariable=self.saved_trade_details, wraplength=850, justify="left").pack(anchor="w", pady=(8, 8))
        saved_actions = ttk.Frame(self.saved_trade_tab)
        saved_actions.pack(fill="x")
        ttk.Button(saved_actions, text="Reopen Trade", command=self._reopen_saved_trade).pack(side="left")
        ttk.Button(saved_actions, text="Add Counteroffer", command=lambda: self._reopen_saved_trade(as_counteroffer=True)).pack(side="left", padx=(8, 0))
        ttk.Button(saved_actions, text="Mark Accepted", command=lambda: self._mark_saved_trade_outcome("Accepted")).pack(side="left", padx=(8, 0))
        ttk.Button(saved_actions, text="Mark Rejected", command=lambda: self._mark_saved_trade_outcome("Rejected")).pack(side="left", padx=8)
        ttk.Button(saved_actions, text="Mark Pending", command=lambda: self._mark_saved_trade_outcome("Proposed")).pack(side="left")
        self.copy_outcome_button = ttk.Button(
            saved_actions, text="Copy Proposed Text", command=self._copy_saved_trade_outcome_text, state="disabled"
        )
        self.copy_outcome_button.pack(side="left", padx=8)
        ttk.Button(saved_actions, text="Delete Saved Trade", command=self._delete_saved_trade).pack(side="right")
        ttk.Button(saved_actions, text="Import Trades", command=self._import_saved_trades).pack(side="right", padx=(0, 8))
        ttk.Button(saved_actions, text="Export Trades", command=self._export_saved_trades).pack(side="right", padx=(0, 8))

        #Trade Targets lists bench-player one-for-one ideas based on reciprocal position needs.
        ideas_tab = ttk.Frame(self.workspace, padding=12)
        self.workspace.add(ideas_tab, text="Trade Targets")
        ttk.Label(ideas_tab, text="Potential trade partners based on reciprocal roster needs.", font=("TkDefaultFont", 11, "bold")).pack(anchor="w")
        ttk.Label(ideas_tab, text="Suggestions focus on bench depth and market values. They are starting points, not predictions that another manager will accept.", wraplength=900).pack(anchor="w", pady=(3, 8))
        ideas_credit = ttk.Label(ideas_tab, text="Trade values provided by FantasyCalc · fantasycalc.com", foreground="#1f5a92", cursor="hand2")
        self.attribution_labels.append(ideas_credit)
        ideas_credit.pack(anchor="w", pady=(0, 6))
        ideas_credit.bind("<Button-1>", lambda _event: __import__("webbrowser").open("https://fantasycalc.com/"))
        ideas_controls = ttk.Frame(ideas_tab)
        ideas_controls.pack(fill="x", pady=(0, 8))
        ttk.Label(ideas_controls, text="Your team:").pack(side="left")
        self.suggestion_team_combo = ttk.Combobox(ideas_controls, state="readonly", width=38)
        self.suggestion_team_combo.pack(side="left", padx=8)
        self.suggestion_team_combo.bind("<<ComboboxSelected>>", self._refresh_trade_ideas)
        ttk.Button(ideas_controls, text="Review Selected Idea", command=self._open_selected_trade_idea).pack(side="right")
        ideas_list = ttk.Frame(ideas_tab)
        ideas_list.pack(fill="both", expand=True)
        self.trade_ideas_tree = ttk.Treeview(
            ideas_list,
            columns=("match", "offer", "offer_value", "target", "target_value", "fit"),
            show="tree headings", selectmode="browse", height=8,
        )
        self.trade_ideas_tree.heading("#0", text="Trade Partner")
        for key, label in (("match", "Position Fit"), ("offer", "You Could Offer"),
                           ("offer_value", "Value"), ("target", "Ask About"),
                           ("target_value", "Value"), ("fit", "Market Fit")):
            self.trade_ideas_tree.heading(key, text=label)
        self.trade_ideas_tree.column("#0", width=165)
        self.trade_ideas_tree.column("match", width=190)
        self.trade_ideas_tree.column("offer", width=190)
        self.trade_ideas_tree.column("offer_value", width=85, anchor="e")
        self.trade_ideas_tree.column("target", width=190)
        self.trade_ideas_tree.column("target_value", width=85, anchor="e")
        self.trade_ideas_tree.column("fit", width=110)
        ideas_scroll = ttk.Scrollbar(ideas_list, orient="vertical", command=self.trade_ideas_tree.yview)
        ideas_xscroll = ttk.Scrollbar(ideas_list, orient="horizontal", command=self.trade_ideas_tree.xview)
        self.trade_ideas_tree.configure(yscrollcommand=ideas_scroll.set, xscrollcommand=ideas_xscroll.set)
        self.trade_ideas_tree.pack(side="left", fill="both", expand=True)
        ideas_scroll.pack(side="right", fill="y")
        ideas_xscroll.pack(side="bottom", fill="x")
        self.trade_ideas_tree.bind("<Double-1>", lambda _event: self._open_selected_trade_idea())
        self.trade_ideas_summary = tk.StringVar(value="Load a league to find possible trade fits.")
        ttk.Label(ideas_tab, textvariable=self.trade_ideas_summary, wraplength=900).pack(anchor="w", pady=(8, 0))
        self.trade_ideas_by_id = {}

        #Trade Builder finds near-even player packages across every opposing roster.
        builder_tab = ttk.Frame(self.workspace, padding=12)
        self.workspace.add(builder_tab, text="Trade Builder")
        ttk.Label(
            builder_tab, text="Build a market-balanced offer from players on your roster.",
            font=("TkDefaultFont", 11, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            builder_tab,
            text="Choose the players you would send. Suggestions compare current FantasyCalc values and are starting points, not acceptance predictions.",
            wraplength=1000,
        ).pack(anchor="w", pady=(3, 7))
        builder_controls = ttk.Frame(builder_tab)
        builder_controls.pack(fill="x", pady=(0, 7))
        ttk.Label(builder_controls, text="Your team:").pack(side="left")
        self.trade_builder_team_combo = ttk.Combobox(builder_controls, state="readonly", width=36)
        self.trade_builder_team_combo.pack(side="left", padx=(8, 16))
        self.trade_builder_team_combo.bind("<<ComboboxSelected>>", self._trade_builder_team_changed)
        ttk.Label(builder_controls, text="Max value gap:").pack(side="left")
        self.trade_builder_tolerance_var = tk.StringVar(value="15%")
        self.trade_builder_tolerance_combo = ttk.Combobox(
            builder_controls, textvariable=self.trade_builder_tolerance_var,
            values=("5%", "10%", "15%", "20%", "25%", "35%"),
            state="readonly", width=7,
        )
        self.trade_builder_tolerance_combo.pack(side="left", padx=(8, 14))
        self.trade_builder_tolerance_combo.bind("<<ComboboxSelected>>", self._trade_builder_options_changed)
        ttk.Button(builder_controls, text="Find Matches", command=self._build_trade_suggestions).pack(side="left")
        builder_content = ttk.Frame(builder_tab)
        builder_content.pack(fill="both", expand=True)
        builder_content.columnconfigure(0, weight=1)
        builder_content.columnconfigure(1, weight=2)
        builder_content.rowconfigure(0, weight=1)
        give_box = ttk.LabelFrame(builder_content, text="Players You Would Send", padding=6)
        return_box = ttk.LabelFrame(builder_content, text="Near-Value Returns", padding=6)
        give_box.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        return_box.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        builder_search = ttk.Frame(give_box)
        builder_search.pack(fill="x", pady=(0, 5))
        ttk.Label(builder_search, text="Search:").pack(side="left")
        self.trade_builder_search_var = tk.StringVar()
        ttk.Entry(builder_search, textvariable=self.trade_builder_search_var).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )
        self.trade_builder_give_tree = ttk.Treeview(
            give_box, columns=("position", "value", "slot"),
            show="tree headings", selectmode="extended", height=10,
        )
        self.trade_builder_give_tree.heading("#0", text="Player", command=lambda: self._sort_trade_builder_tree(self.trade_builder_give_tree, "#0"))
        self.trade_builder_give_tree.heading("position", text="Position", command=lambda: self._sort_trade_builder_tree(self.trade_builder_give_tree, "position"))
        self.trade_builder_give_tree.heading("value", text="FC Value", command=lambda: self._sort_trade_builder_tree(self.trade_builder_give_tree, "value"))
        self.trade_builder_give_tree.heading("slot", text="Slot", command=lambda: self._sort_trade_builder_tree(self.trade_builder_give_tree, "slot"))
        self.trade_builder_give_tree.column("#0", width=170)
        self.trade_builder_give_tree.column("position", width=75, anchor="center")
        self.trade_builder_give_tree.column("value", width=75, anchor="e")
        self.trade_builder_give_tree.column("slot", width=75, anchor="center")
        give_scroll = ttk.Scrollbar(give_box, orient="vertical", command=self.trade_builder_give_tree.yview)
        self.trade_builder_give_tree.configure(yscrollcommand=give_scroll.set)
        self.trade_builder_give_tree.pack(side="left", fill="both", expand=True)
        give_scroll.pack(side="right", fill="y")
        self.trade_builder_give_tree.bind("<<TreeviewSelect>>", self._update_trade_builder_selection)
        self.trade_builder_search_var.trace_add("write", lambda *_: self._filter_trade_builder_roster())
        self.trade_builder_give_items = []
        self.trade_builder_selection_var = tk.StringVar(value="Select one or more players to send.")
        ttk.Label(give_box, textvariable=self.trade_builder_selection_var, wraplength=360).pack(anchor="w", pady=(5, 0))
        return_filters = ttk.Frame(return_box)
        return_filters.pack(fill="x", pady=(0, 5))
        ttk.Label(return_filters, text="Position:").pack(side="left")
        self.trade_builder_position_filter = ttk.Combobox(
            return_filters, state="readonly", width=8, values=("Any",),
        )
        self.trade_builder_position_filter.set("Any")
        self.trade_builder_position_filter.pack(side="left", padx=(4, 10))
        self.trade_builder_position_filter.bind("<<ComboboxSelected>>", self._filter_trade_builder_matches)
        ttk.Label(return_filters, text="Players received:").pack(side="left")
        self.trade_builder_package_size_filter = ttk.Combobox(
            return_filters, state="readonly", width=9, values=("Any", "1", "2", "3"),
        )
        self.trade_builder_package_size_filter.set("Any")
        self.trade_builder_package_size_filter.pack(side="left", padx=(4, 0))
        self.trade_builder_package_size_filter.bind("<<ComboboxSelected>>", self._filter_trade_builder_matches)
        ttk.Button(
            return_filters, text="Reset Filters", command=self._reset_trade_builder_filters,
        ).pack(side="left", padx=(10, 0))
        self.trade_builder_results_tree = ttk.Treeview(
            return_box, columns=("players", "positions", "value", "difference", "fit"),
            show="tree headings", selectmode="browse", height=10,
        )
        self.trade_builder_results_tree.heading("#0", text="Trade Partner", command=lambda: self._sort_trade_builder_tree(self.trade_builder_results_tree, "#0"))
        for key, label in (("players", "You Receive"), ("positions", "Positions"), ("value", "Total Value"),
                           ("difference", "Value Gap"), ("fit", "Roster Fit")):
            self.trade_builder_results_tree.heading(key, text=label, command=lambda column=key: self._sort_trade_builder_tree(self.trade_builder_results_tree, column))
        self.trade_builder_results_tree.column("#0", width=150)
        self.trade_builder_results_tree.column("players", width=300)
        self.trade_builder_results_tree.column("positions", width=100, anchor="center")
        self.trade_builder_results_tree.column("value", width=95, anchor="e")
        self.trade_builder_results_tree.column("difference", width=100, anchor="e")
        self.trade_builder_results_tree.column("fit", width=120)
        results_scroll = ttk.Scrollbar(return_box, orient="vertical", command=self.trade_builder_results_tree.yview)
        self.trade_builder_results_tree.configure(yscrollcommand=results_scroll.set)
        self.trade_builder_results_tree.pack(side="left", fill="both", expand=True)
        results_scroll.pack(side="right", fill="y")
        self.trade_builder_results_tree.bind("<Double-1>", lambda _event: self._open_trade_builder_match())
        self.trade_builder_summary = tk.StringVar(value="Select players to find near-value returns from other teams.")
        builder_actions = ttk.Frame(builder_tab)
        builder_actions.pack(fill="x", pady=(6, 0))
        ttk.Label(builder_actions, textvariable=self.trade_builder_summary, wraplength=1100).pack(side="left", fill="x", expand=True)
        ttk.Button(builder_actions, text="Review Selected Match", command=self._open_trade_builder_match).pack(side="right")
        self.trade_builder_matches = {}
        self.trade_builder_all_matches = []
        self.trade_builder_sort_state = {}

        #Show the current operation message and its activity indicator.
        info = ttk.Frame(main)
        info.pack(fill="x", pady=(10, 0))
        ttk.Label(info, textvariable=self.status_var, wraplength=820).pack(anchor="w", pady=(0, 6))
        self.progress = ttk.Progressbar(info, mode="indeterminate")

    def _refresh_roster_view(self, _event=None):
        #Show the selected roster, or all league players while a global search is active.
        if not hasattr(self, "roster_tree"):
            return
        if _event is not None:
            self.roster_history_player_id = None
            self.roster_history_points = []
            self.roster_history_detail.set("Select a player to see recent value movement and volatility.")
            self._draw_history_chart(self.roster_history_canvas, [], "Select a player to load a 90-day value trend.")
        for _group_id, player_id, _owner in self.roster_search_items:
            if self.roster_tree.exists(player_id):
                self.roster_tree.delete(player_id)
        for item in self.roster_tree.get_children():
            self.roster_tree.delete(item)
        self.roster_search_items = []
        if self.dataframes is None:
            return
        global_search = bool(self.roster_search_var.get().strip())
        self.roster_search_global_mode = global_search
        roster_id = self._selected_roster_id(self.roster_team_combo)
        if not global_search and roster_id is None:
            return
        roster_rows = self.dataframes["Rosters"]
        if not global_search:
            roster_rows = roster_rows[roster_rows["Roster ID"].astype(str) == str(roster_id)]
        groups = {} if global_search else {
            "starting": self.roster_tree.insert("", "end", iid="roster-group-start", text="Starting Lineup", open=True),
            "bench": self.roster_tree.insert("", "end", iid="roster-group-bench", text="Bench", open=True),
            "reserve": self.roster_tree.insert("", "end", iid="roster-group-reserve", text="IR / Reserve", open=True),
        }
        owner_groups = {}
        position_order = {position: index for index, position in enumerate(("QB", "RB", "WR", "TE", "K", "DEF", "DST"))}
        ordered_rows = list(roster_rows.iterrows())
        ordered_rows.sort(key=lambda item: (
            2 if str(item[1].get("Slot", "")).casefold() in {"reserve", "ir"} else
            1 if str(item[1].get("Slot", "")).casefold() == "bench" else 0,
            item[1].get("Lineup Order") if pd.notna(item[1].get("Lineup Order")) else
            item[1].get("Roster Order", 999),
            position_order.get(str(item[1].get("Position", "")).split(",")[0].strip().upper(), 99),
        ))
        for _index, row in ordered_rows:
            slot = str(row.get("Slot", "")).strip()
            owner = str(row.get("Team", "Unknown team"))
            if global_search:
                group = owner_groups.get(owner)
                if group is None:
                    group = f"roster-owner-{len(owner_groups)}"
                    owner_groups[owner] = group
                    self.roster_tree.insert("", "end", iid=group, text=owner, open=True)
            else:
                group = "reserve" if slot.casefold() in {"reserve", "ir"} else (
                    "bench" if slot.casefold() == "bench" else "starting"
                )
                group = groups[group]
            player_id = str(row.get("Player ID", ""))
            team_code = self._normalize_nfl_team_code(row.get("NFL Team", ""))
            bye = self.weekly_bye_weeks.get(team_code) or (
                row.get("Bye Week") if pd.notna(row.get("Bye Week")) and row.get("Bye Week") else "—"
            )
            projection = self.roster_projection_values.get(player_id)
            projection_text = f"{projection:.1f}" if projection is not None else "—"
            status_row = {"Status": row.get("Status"), "Injury Status": row.get("Injury Status"), "Slot": slot}
            self.roster_tree.insert(
                group, "end", iid=player_id, text=str(row.get("Player", player_id)),
                values=(slot, row.get("Position", ""), row.get("NFL Team", ""),
                        self._player_list_status(status_row), bye, projection_text),
            )
            self.roster_search_items.append((group, player_id, owner))
        self._apply_roster_search()

    def _apply_roster_search(self):
        #Filter the loaded roster scope and rebuild when switching between local and global search.
        if not hasattr(self, "roster_tree"):
            return
        query = self.roster_search_var.get().strip().casefold()
        global_search = bool(query)
        if global_search != getattr(self, "roster_search_global_mode", False):
            self._refresh_roster_view()
            return
        shown = 0
        for group_id, player_id, owner in self.roster_search_items:
            if not self.roster_tree.exists(player_id):
                continue
            searchable = " ".join((
                str(self.roster_tree.item(player_id, "text")),
                owner,
                *(str(value) for value in self.roster_tree.item(player_id, "values")),
            )).casefold()
            if not query or query in searchable:
                self.roster_tree.move(player_id, group_id, "end")
                shown += 1
            else:
                self.roster_tree.detach(player_id)
        total = len(self.roster_search_items)
        if not total:
            self.roster_search_status.set("Load a league to search its roster.")
        elif query and shown == 0:
            self.roster_search_status.set("No players match across the league's rosters.")
        elif query:
            self.roster_search_status.set(f"Searching all league rosters · showing {shown} of {total} players.")
        else:
            self.roster_search_status.set(f"Showing {shown} of {total} players on the selected roster.")

    def _load_roster_projections(self):
        #Request current-week point projections when Sleeper has published them.
        league = self.selected_league() or {}
        league_id = str(league.get("league_id", ""))
        season = league.get("season")
        state_season = self.weekly_state.get("league_season") or self.weekly_state.get("season")
        week = self.weekly_state.get("display_week") or self.weekly_state.get("week") or self.weekly_state.get("leg")
        if not season or not week or (state_season and str(state_season) != str(season)):
            self.roster_projection_values = {}
            self.roster_projection_error = ""
            self.roster_projection_note.set("No current-week projections for this league season.")
            self._refresh_roster_view()
            return
        scoring = league.get("scoring_settings") or {}
        reception_value = float(scoring.get("rec", 0) or 0)
        ppr = min((0, 0.5, 1), key=lambda value: abs(value - reception_value))
        self.roster_projection_note.set(f"Loading Week {week} Sleeper projections…")

        def task():
            try:
                values, metric = self.api.get_projections(season, week, ppr)
                return values, metric, None
            except Exception as exc:
                return {}, "", str(exc)

        def loaded(result):
            values, metric, error = result
            current_league = self.selected_league() or {}
            if (str(current_league.get("league_id", "")) != league_id or
                    str(current_league.get("season", "")) != str(season)):
                return
            self.roster_projection_values = values
            self.roster_projection_metric = metric
            self.roster_projection_error = error or ""
            if error:
                self.roster_projection_note.set("Sleeper projections aren't available right now.")
            elif not values:
                self.roster_projection_note.set("Sleeper hasn't published projections for this week yet.")
            else:
                ppr_text = {"pts_std": "Standard", "pts_half_ppr": "Half-PPR", "pts_ppr": "PPR"}.get(metric, metric)
                scoring_note = "" if abs(ppr - reception_value) < 0.001 else " · closest scoring match"
                self.roster_projection_note.set(f"Week {week} Sleeper projections · {ppr_text} baseline{scoring_note}")
            self._refresh_roster_view()

        self.run_background(task, loaded)


    def _size_window_to_screen(self):
        #Open maximized so the full workspace is visible without manual resizing.
        self.update_idletasks()
        try:
            self.state("zoomed")
        except tk.TclError:
            #Use a large window on systems without a maximize state.
            width = min(1200, self.winfo_screenwidth())
            height = min(950, self.winfo_screenheight())
            self.geometry(f"{width}x{height}+0+0")

    def _scroll_trade_tab(self, event):
        #Use the mouse wheel to reach lower Trade Review sections on short displays.
        if self.workspace.select() != str(self.trade_tab):
            return
        if event.widget.winfo_toplevel() is not self:
            return
        if isinstance(event.widget, (ttk.Treeview, ttk.Combobox)):
            return
        units = -1 * int(event.delta / 120) if event.delta else 0
        if units:
            self.trade_tab_canvas.yview_scroll(units, "units")
            return "break"

    def _compact_tree_columns(self):
        #Keep list columns close to their configured widths instead of stretching across wide monitors.
        pending = [self]
        while pending:
            parent = pending.pop()
            for child in parent.winfo_children():
                pending.append(child)
                if isinstance(child, ttk.Treeview):
                    for column in ("#0", *child["columns"]):
                        child.column(column, stretch=False)

    def _show_lineup_impact(self):
        #Show detailed lineup notes on demand to keep the trade screen uncluttered.
        messagebox.showinfo("Roster and Lineup Impact", self.trade_impact_var.get())

    def _bye_week_for_player(self, row):
        #Prefer the schedule-derived bye because Sleeper's player map often omits it.
        team_code = self._normalize_nfl_team_code(row.get("NFL Team", ""))
        bye_week = self.weekly_bye_weeks.get(team_code)
        if bye_week:
            return str(bye_week)
        sleeper_bye = row.get("Bye Week")
        if pd.notna(sleeper_bye) and str(sleeper_bye).strip():
            return str(sleeper_bye)
        return "—"

    def _refresh_trade_bye_weeks(self):
        #Update the bye column in place so async schedule loading keeps current selections.
        if self.dataframes is None:
            return
        rosters = self.dataframes["Rosters"]
        rows_by_player = {
            str(row["Player ID"]): row for _, row in rosters.iterrows()
        }
        for tree in (self.trade_left_tree, self.trade_right_tree):
            for iid in tree.get_children():
                tags = tree.item(iid, "tags")
                row = rows_by_player.get(str(tags[0])) if tags else None
                if row is None:
                    continue
                values = list(tree.item(iid, "values"))
                if len(values) > 3:
                    values[3] = self._bye_week_for_player(row)
                    tree.item(iid, values=values)

    def _make_trade_tree(self, parent):
        #Create a sortable, multi-select roster list for either side of a trade.
        visible_rows = 8 if self.winfo_screenheight() > 1200 else (4 if self.winfo_screenheight() <= 1100 else 6)
        tree = ttk.Treeview(parent, columns=("position", "status", "team", "bye", "position_rank", "overall_rank", "value"), show="tree headings", selectmode="extended", height=visible_rows)
        tree.heading("#0", text="Player", command=lambda t=tree: self._sort_trade_tree(t, "name"))
        tree.heading("position", text="Position", command=lambda t=tree: self._sort_trade_tree(t, "position"))
        tree.heading("status", text="Status", command=lambda t=tree: self._sort_trade_tree(t, "status"))
        tree.heading("team", text="Team", command=lambda t=tree: self._sort_trade_tree(t, "team"))
        tree.heading("bye", text="Bye")
        tree.heading("position_rank", text="FC Pos Rank")
        tree.heading("overall_rank", text="FC Overall")
        tree.heading("value", text="FC Value", command=lambda t=tree: self._sort_trade_tree(t, "value"))
        tree.column("#0", width=145)
        tree.column("position", width=58, anchor="center")
        tree.column("status", width=85, anchor="center")
        tree.column("team", width=48, anchor="center")
        tree.column("bye", width=42, anchor="center")
        tree.column("position_rank", width=74, anchor="center")
        tree.column("overall_rank", width=66, anchor="center")
        tree.column("value", width=74, anchor="e")
        return tree

    def write_log(self, text):
        #Retained as a hook for older status calls; the current UI uses one status line.
        pass

    def status(self, text):
        #Update the visible status text and let Tk redraw before continuing.
        self.status_var.set(text)
        self.write_log(text)
        self.update_idletasks()

    def _selected_roster_id(self, combo):
        #Translate the team label shown in a combo box back to Sleeper's roster ID.
        label = combo.get()
        return next((roster_id for display, roster_id in self.team_options if display == label), None)

    def _show_position_needs(self, _event=None):
        #Fill the depth table for the selected team and summarize the likely thin spots.
        if self.dataframes is None:
            return
        roster_id = self._selected_roster_id(self.needs_team_combo)
        needs = self.dataframes.get("Position Needs", pd.DataFrame())
        for item in self.needs_tree.get_children():
            self.needs_tree.delete(item)
        if roster_id is None or needs.empty:
            self.needs_summary_var.set("No position requirements were found for this league.")
            return
        selected = needs[needs["Roster ID"].astype(str) == str(roster_id)]
        gaps = []
        for _, row in selected.iterrows():
            self.needs_tree.insert("", "end", text=row["Position"], values=(
                row["Required"], row["Starting"], row["Rostered"], row["Bench"], row["Assessment"]
            ))
            if row["Starting"] < row["Required"]:
                gaps.append(f'{row["Position"]} ({int(row["Required"] - row["Starting"])} starter gap)')
            elif row["Position"] not in {"K", "DEF"} and row["Rostered"] <= row["Required"]:
                gaps.append(f'{row["Position"]} depth')
        self.needs_summary_var.set(
            "Potential needs: " + (", ".join(gaps) if gaps else "No obvious positional gaps by roster counts.")
        )

    def _populate_trade_rosters(self, _event=None):
        #Refresh only the list whose team changed so the other side's selection stays selected.
        if self.dataframes is None:
            return
        rosters = self.dataframes["Rosters"]
        self._trade_populating = True
        roster_lists = ((self.trade_left_team, self.trade_left_tree),
                        (self.trade_right_team, self.trade_right_tree))
        if _event is not None:
            roster_lists = tuple(pair for pair in roster_lists if pair[0] is _event.widget)
        try:
            for combo, tree in roster_lists:
                for item in tree.get_children():
                    tree.delete(item)
                roster_id = self._selected_roster_id(combo)
                if roster_id is None:
                    continue
                team = rosters[rosters["Roster ID"].astype(str) == str(roster_id)]
                self.trade_sort_state[id(tree)] = {}
                for _, row in team.iterrows():
                    player_id = str(row["Player ID"])
                    market = self.fantasycalc_values.get(player_id)
                    shown_value = f'{market:,.0f}' if market is not None else "N/A"
                    bye = self._bye_week_for_player(row)
                    position_ranks = self.fantasycalc_position_ranks.get(player_id, {})
                    position_rank = " / ".join(
                        f"{position} {rank}" for position, rank in position_ranks.items()
                    ) or "—"
                    overall_rank = self.fantasycalc_overall_ranks.get(player_id, "—")
                    name = str(row["Player"])
                    last_name = str(row.get("Last Name") or name.rsplit(" ", 1)[-1])
                    first_name = str(row.get("First Name") or "")
                    self.trade_sort_info[player_id] = {
                        "name": (last_name.casefold(), first_name.casefold(), name.casefold()),
                        "position": self._position_sort_key(row["Position"], last_name, first_name),
                        "status": self._player_list_status(row).casefold(),
                        "team": str(row["NFL Team"]).casefold(),
                        "value": market,
                    }
                    tree.insert("", "end", iid=f"{id(tree)}-{player_id}",
                                text=row["Player"], values=(row["Position"], self._player_list_status(row),
                                                            row["NFL Team"] or "—", bye, position_rank, overall_rank, shown_value),
                                tags=(player_id,))
                self._sort_trade_tree(tree, "position", toggle=False)
        finally:
            self._trade_populating = False
        self._update_trade_review()

    def _clear_trade_selection(self, tree):
        #Clear one side's selected players while leaving the opposite side untouched.
        selection = tree.selection()
        if selection:
            tree.selection_remove(*selection)

    @staticmethod
    def _position_sort_key(position, last_name="", first_name=""):
        #Use the requested QB/RB/WR/TE/K priority and name as a tie-breaker.
        priority = {"QB": 0, "RB": 1, "WR": 2, "TE": 3, "K": 4}
        positions = [part.strip().upper() for part in str(position).split(",")]
        rank = min((priority.get(part, 5) for part in positions), default=5)
        return rank, str(last_name).casefold(), str(first_name).casefold()

    def _sort_trade_tree(self, tree, column, toggle=True):
        #Sort names by surname, positions by priority, or values as numbers; keep unknown values last.
        tree_id = id(tree)
        state = self.trade_sort_state.setdefault(tree_id, {})
        if toggle:
            reverse = not state[column] if column in state else column == "value"
        else:
            reverse = column == "value"
        state[column] = reverse
        items = list(tree.get_children(""))

        def player_id_for(item):
            tags = tree.item(item, "tags")
            return tags[0] if tags else ""

        if column == "value":
            known = [item for item in items if self.trade_sort_info.get(player_id_for(item), {}).get("value") is not None]
            missing = [item for item in items if item not in known]
            known.sort(key=lambda item: self.trade_sort_info[player_id_for(item)]["value"], reverse=reverse)
            ordered = known + missing
        else:
            ordered = sorted(
                items,
                key=lambda item: self.trade_sort_info.get(player_id_for(item), {}).get(column, ()),
                reverse=reverse,
            )
        for index, item in enumerate(ordered):
            tree.move(item, "", index)

    def _sort_trade_builder_tree(self, tree, column):
        #Sort builder rows by the selected field, toggling direction on repeated clicks.
        tree_key = "give" if tree is self.trade_builder_give_tree else "results"
        state = self.trade_builder_sort_state.setdefault(tree_key, {})
        reverse = not state[column] if column in state else column == "value"
        state[column] = reverse
        items = list(tree.get_children(""))

        if tree_key == "give":
            def sort_key(item):
                if column == "#0":
                    return str(tree.item(item, "text")).casefold()
                values = tree.item(item, "values")
                index = {"position": 0, "slot": 2}[column]
                return str(values[index] if len(values) > index else "").casefold()

            if column == "value":
                def order_by_value(player_ids):
                    known = [item for item in player_ids if self.fantasycalc_values.get(str(item)) is not None]
                    missing = [item for item in player_ids if self.fantasycalc_values.get(str(item)) is None]
                    known.sort(key=lambda item: self.fantasycalc_values[str(item)], reverse=reverse)
                    return known + missing
                items = order_by_value(items)
                self.trade_builder_give_items = order_by_value(self.trade_builder_give_items)
            else:
                items.sort(key=sort_key, reverse=reverse)
                self.trade_builder_give_items.sort(key=sort_key, reverse=reverse)
            self._filter_trade_builder_roster()
        else:
            def sort_key(item):
                match = self.trade_builder_matches.get(item, {})
                if column == "#0":
                    return str(match.get("partner", "")).casefold()
                if column == "players":
                    return " + ".join(asset["name"] for asset in match.get("receive_assets", [])).casefold()
                if column == "positions":
                    return " + ".join(asset["position"] for asset in match.get("receive_assets", [])).casefold()
                if column == "value":
                    return match.get("receive_value", 0)
                if column == "difference":
                    return match.get("gap", 0)
                if column == "fit":
                    return match.get("fit_order", 99)
                return ""

            if column == "#0":
                items.sort(key=lambda item: str(tree.item(item, "text")).casefold(), reverse=reverse)
                for index, item in enumerate(items):
                    tree.move(item, "", index)
            else:
                for group in items:
                    children = list(tree.get_children(group))
                    children.sort(key=sort_key, reverse=reverse)
                    for index, item in enumerate(children):
                        tree.move(item, group, index)

    @staticmethod
    def _player_list_status(row):
        #Prefer a current injury designation, then show IR/reserve or Sleeper's general status.
        injury = str(row.get("Injury Status") or "").strip()
        player_status = str(row.get("Status") or "").strip()
        slot = str(row.get("Slot") or "").strip().casefold()
        if injury and injury.casefold() not in {"healthy", "none", "null"}:
            return injury
        if slot in {"reserve", "ir"}:
            return "IR"
        return player_status or "—"

    def _update_trade_balance(self, give_value=None, receive_value=None):
        #Show the relative value split, centered when both sides are even.
        canvas = self.trade_balance_canvas
        canvas.delete("all")
        canvas.configure(bg=self.theme_colors["background"])
        width, height = 260, 18
        left, right, top, bottom = 2, width - 2, 3, height - 3
        center = (left + right) / 2
        colors = self.theme_colors
        canvas.create_rectangle(left, top, center, bottom, fill="#9b5c5c", outline="")
        canvas.create_rectangle(center, top, right, bottom, fill="#4d9568", outline="")
        canvas.create_line(center, top - 1, center, bottom + 1, fill=colors["foreground"], width=1)
        if give_value is None or receive_value is None:
            self.trade_balance_summary.set("Select players on both sides.")
            marker = center
        else:
            denominator = max(give_value, receive_value)
            if denominator <= 0:
                difference = 0
                gap = 0
            else:
                difference = receive_value - give_value
                gap = abs(difference) / denominator
            marker = center + (difference / denominator if denominator > 0 else 0) * (right - left) / 2
            if gap <= 0.10:
                self.trade_balance_summary.set(f"Close value · {gap:.0%} gap")
            elif difference > 0:
                self.trade_balance_summary.set(f"You receive more · {gap:.0%} gap")
            else:
                self.trade_balance_summary.set(f"You give more · {gap:.0%} gap")
        marker = min(max(marker, left + 3), right - 3)
        canvas.create_line(marker, top - 2, marker, bottom + 2,
                           fill=colors["foreground"], width=3)

    def _update_trade_review(self, _event=None):
        #Recalculate market totals and update the positional roster impact as selections change.
        if self._trade_populating:
            return
        self._update_position_impact()
        give_selection = self.trade_left_tree.selection()
        get_selection = self.trade_right_tree.selection()
        self._update_trade_text_preview()
        self.copy_trade_button.configure(
            state="normal" if give_selection and get_selection else "disabled"
        )
        give_ids = [self.trade_left_tree.item(iid, "tags")[0] for iid in give_selection]
        get_ids = [self.trade_right_tree.item(iid, "tags")[0] for iid in get_selection]
        self._update_stats_guy_trade_review(give_ids, get_ids)
        self._refresh_trade_history(give_selection, get_selection)
        if not give_ids and not get_ids:
            self._update_trade_balance()
            self.trade_result_var.set("Select players on either side to compare their FantasyCalc values.")
            return
        unknown = [pid for pid in give_ids + get_ids if pid not in self.fantasycalc_values]
        if unknown:
            self._update_trade_balance()
            self.trade_result_var.set("Values are unavailable for one or more selected players. Try refreshing later.")
            return
        give_value = sum(self.fantasycalc_values[pid] for pid in give_ids)
        get_value = sum(self.fantasycalc_values[pid] for pid in get_ids)
        self._update_trade_balance(
            give_value if give_ids else None, get_value if get_ids else None
        )
        difference = get_value - give_value
        prefix = "You give: none" if not give_ids else f"You give: {give_value:,.0f}"
        receive = "You receive: none" if not get_ids else f"You receive: {get_value:,.0f}"
        if not give_ids or not get_ids:
            verdict = "Select players on both sides to review the offer."
        elif abs(difference) <= max(give_value, get_value) * 0.10:
            verdict = "Market values are fairly close."
        elif difference > 0:
            verdict = f"Market leans toward you receiving more by {difference:,.0f}."
        else:
            verdict = f"Market leans toward you giving more by {abs(difference):,.0f}."
        self.trade_result_var.set(f"{prefix}  |  {receive}\n{verdict}")

    def _update_stats_guy_trade_review(self, give_ids, get_ids):
        #Show the same proposed trade through Stats Guy's independent value model.
        if not give_ids and not get_ids:
            self.trade_stats_guy_result_var.set("Stats Guy totals will appear when you select players.")
            return
        if not self.stats_guy_players:
            message = "Stats Guy values are still loading."
            if self.stats_guy_error:
                message = f"Stats Guy values are unavailable: {self.stats_guy_error}"
            self.trade_stats_guy_result_var.set(message)
            return

        value_format = self._stats_guy_format()

        def get_value(player_id):
            card = self.stats_guy_players.get(str(player_id), {})
            values = card.get("value") or {}
            return values.get(value_format) if isinstance(values, dict) else None

        selected_ids = give_ids + get_ids
        missing = [player_id for player_id in selected_ids if get_value(player_id) is None]
        if missing:
            self.trade_stats_guy_result_var.set(
                f"Stats Guy does not have a value for {len(missing)} selected player(s). "
                "Its API covers QB, RB, WR, and TE."
            )
            return
        give_value = sum(float(get_value(player_id)) for player_id in give_ids)
        get_value_total = sum(float(get_value(player_id)) for player_id in get_ids)
        text = f"Stats Guy: You give {give_value:,.0f}  |  You receive {get_value_total:,.0f}."
        if not give_ids or not get_ids:
            text += " Select players on both sides to compare the offer."
        else:
            difference = get_value_total - give_value
            if abs(difference) <= max(give_value, get_value_total) * 0.10:
                text += " Stats Guy values are fairly close."
            elif difference > 0:
                text += f" Stats Guy leans toward you receiving {difference:,.0f} more value."
            else:
                text += f" Stats Guy leans toward you giving {abs(difference):,.0f} more value."
        self.trade_stats_guy_result_var.set(text)

    def _selected_trade_copy_text(self):
        #Build the same user-perspective text for the live preview and clipboard action.
        give_players = [self.trade_left_tree.item(iid, "text") for iid in self.trade_left_tree.selection()]
        get_players = [self.trade_right_tree.item(iid, "text") for iid in self.trade_right_tree.selection()]
        if not give_players or not get_players:
            return None
        return f"I get: {', '.join(get_players)}\nI send: {', '.join(give_players)}"

    def _update_trade_text_preview(self):
        #Keep the visible preview exactly aligned with what the copy button will produce.
        self.copy_trade_preview_var.set(
            self._selected_trade_copy_text()
            or "Select at least one player from each team to preview the copied text."
        )

    def _copy_trade_text(self):
        #Copy the same proposal shown below the player lists.
        trade_text = self._selected_trade_copy_text()
        if not trade_text:
            return
        self.clipboard_clear()
        self.clipboard_append(trade_text)
        self.update()
        self.status("Trade text copied to the clipboard.")

    def _current_trade_outcome(self):
        #Use a saved outcome only while the live selection still matches that saved trade.
        trade = self._find_saved_trade(self.active_saved_trade_id) if self.active_saved_trade_id else None
        if not trade:
            return "Proposed"
        current_give = {asset["id"] for asset in self._selected_trade_assets(self.trade_left_tree)}
        current_receive = {asset["id"] for asset in self._selected_trade_assets(self.trade_right_tree)}
        saved_give = {str(asset.get("id")) for asset in trade.get("give_players", [])}
        saved_receive = {str(asset.get("id")) for asset in trade.get("get_players", [])}
        if current_give == saved_give and current_receive == saved_receive:
            return trade.get("outcome", "Proposed")
        return "Proposed"

    def _copy_trade_outcome_text(self):
        #Copy a proposal or finalized trade using language that matches its saved outcome.
        give_players = [self.trade_left_tree.item(iid, "text") for iid in self.trade_left_tree.selection()]
        receive_players = [self.trade_right_tree.item(iid, "text") for iid in self.trade_right_tree.selection()]
        if not give_players or not receive_players:
            return
        outcome = self._current_trade_outcome()
        if outcome == "Accepted":
            text = (
                "Accepted trade\n"
                f"I sent: {', '.join(give_players)}\n"
                f"I received: {', '.join(receive_players)}"
            )
        elif outcome == "Rejected":
            text = (
                "Rejected trade\n"
                f"I attempted to send: {', '.join(give_players)}\n"
                f"I wanted to receive: {', '.join(receive_players)}"
            )
        else:
            text = f"I get: {', '.join(receive_players)}\nI send: {', '.join(give_players)}"
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        self.status(f"{outcome} trade text copied to the clipboard.")

    def _copy_saved_trade_outcome_text(self):
        #Copy text for the selected saved trade using its current outcome.
        selection = self.saved_trade_tree.selection()
        if not selection:
            messagebox.showinfo("Choose a Trade", "Select a saved trade before copying its text.")
            return
        trade = self._find_saved_trade(selection[0])
        if not trade:
            return
        give_players = [asset.get("name", "Unknown player") for asset in trade.get("give_players", [])]
        receive_players = [asset.get("name", "Unknown player") for asset in trade.get("get_players", [])]
        outcome = trade.get("outcome", "Proposed")
        if outcome == "Accepted":
            text = (
                "Accepted trade\n"
                f"I sent: {', '.join(give_players)}\n"
                f"I received: {', '.join(receive_players)}"
            )
        elif outcome == "Rejected":
            text = (
                "Rejected trade\n"
                f"I attempted to send: {', '.join(give_players)}\n"
                f"I wanted to receive: {', '.join(receive_players)}"
            )
        else:
            text = f"I get: {', '.join(receive_players)}\nI send: {', '.join(give_players)}"
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        self.status(f"{outcome} trade text copied to the clipboard.")

    @staticmethod
    def _change_player_positions(counter, position_text, amount):
        #Add or remove one player's eligible positions from a roster-count summary.
        for position in str(position_text).split(","):
            position = position.strip().upper()
            if position:
                counter[position] += amount

    def _team_position_counts(self, roster_id, excluded_ids=()):
        #Count eligible players at each position, optionally leaving out traded-away players.
        counts = Counter()
        rows = self.dataframes["Rosters"]
        team = rows[rows["Roster ID"].astype(str) == str(roster_id)]
        excluded = {str(pid) for pid in excluded_ids}
        for _, row in team.iterrows():
            if str(row["Player ID"]) not in excluded:
                self._change_player_positions(counts, row["Position"], 1)
        return counts

    def _team_lineup_impact(self, roster_id, outgoing, incoming):
        #Describe starters leaving, eligible bench options, and incoming players for one team.
        roster_rows = self.dataframes["Rosters"]
        team_rows = roster_rows[roster_rows["Roster ID"].astype(str) == str(roster_id)]
        outgoing_ids = {str(asset["id"]) for asset in outgoing}
        outgoing_rows = team_rows[team_rows["Player ID"].astype(str).isin(outgoing_ids)]
        starters = outgoing_rows[
            ~outgoing_rows["Slot"].astype(str).str.casefold().isin({"bench", "reserve", "ir"})
        ]
        incoming_names = [asset["name"] for asset in incoming]

        #Only describe bench replacements for fixed starter slots; FLEX eligibility can be ambiguous.
        starter_notes = []
        for _, starter in starters.iterrows():
            slot = str(starter["Slot"]).strip().upper()
            positions = {part.strip().upper() for part in str(starter["Position"]).split(",")}
            target_positions = {slot} if slot in {"QB", "RB", "WR", "TE", "K", "DEF", "DST"} else positions
            bench_rows = team_rows[
                team_rows["Slot"].astype(str).str.casefold().eq("bench") &
                ~team_rows["Player ID"].astype(str).isin(outgoing_ids)
            ]
            candidates = []
            for _, player in bench_rows.iterrows():
                eligible = {part.strip().upper() for part in str(player["Position"]).split(",")}
                if eligible & target_positions:
                    candidates.append(str(player["Player"]))
            candidates = sorted(set(candidates), key=str.casefold)
            replacement_text = ", ".join(candidates[:3]) if candidates else "no eligible bench option listed"
            starter_notes.append(f"{starter['Player']} ({slot}); bench options: {replacement_text}")

        moved_text = "; ".join(starter_notes) if starter_notes else "no current starter moved"
        incoming_text = ", ".join(incoming_names) if incoming_names else "none"
        return f"Starters moved: {moved_text}. Incoming: {incoming_text}."

    def _update_position_impact(self):
        #Model roster counts and current lineup changes after the proposed exchange.
        if self.dataframes is None or not self.team_options:
            self.trade_impact_var.set("Load a league to see roster depth changes.")
            self.trade_impact_summary.set("Load a league to see roster impact.")
            return
        left_id = self._selected_roster_id(self.trade_left_team)
        right_id = self._selected_roster_id(self.trade_right_team)
        if left_id is None or right_id is None:
            self.trade_impact_var.set("Select two teams to see roster depth changes.")
            self.trade_impact_summary.set("Select two teams to see roster impact.")
            return
        giving = self._selected_trade_assets(self.trade_left_tree)
        receiving = self._selected_trade_assets(self.trade_right_tree)
        give_ids = {asset["id"] for asset in giving}
        receive_ids = {asset["id"] for asset in receiving}
        left_before = self._team_position_counts(left_id)
        right_before = self._team_position_counts(right_id)
        left_after = self._team_position_counts(left_id, give_ids)
        right_after = self._team_position_counts(right_id, receive_ids)
        for asset in receiving:
            self._change_player_positions(left_after, asset["position"], 1)
        for asset in giving:
            self._change_player_positions(right_after, asset["position"], 1)

        def describe(team_name, before, after):
            relevant = ("QB", "RB", "WR", "TE", "K", "DEF")
            changes = [f"{pos} {before[pos]} → {after[pos]}" for pos in relevant if before[pos] != after[pos]]
            return f"{team_name}: " + ("; ".join(changes) if changes else "no positional count change")

        left_lineup = self._team_lineup_impact(left_id, giving, receiving)
        right_lineup = self._team_lineup_impact(right_id, receiving, giving)
        self.trade_impact_summary.set(
            "Roster changes: " + describe(self.trade_left_team.get(), left_before, left_after) +
            "  |  " + describe(self.trade_right_team.get(), right_before, right_after)
        )
        self.trade_impact_var.set(
            "Eligible bench options are based on current lineup assignments; no projections are included.\n\n" +
            self.trade_left_team.get() + "\n" + left_lineup + "\n\n" +
            self.trade_right_team.get() + "\n" + right_lineup
        )

    def _clear_trade_builder_matches(self, message="Select players to find near-value returns from other teams."):
        #Remove suggestions as soon as their offer inputs change.
        for item in self.trade_builder_results_tree.get_children():
            self.trade_builder_results_tree.delete(item)
        self.trade_builder_matches.clear()
        self.trade_builder_all_matches = []
        self.trade_builder_summary.set(message)

    def _trade_builder_options_changed(self, _event=None):
        #Require a fresh search after changing the selected team or value tolerance.
        self._clear_trade_builder_matches("Trade settings changed. Find Matches to refresh suggestions.")

    def _trade_builder_team_changed(self, _event=None):
        #Reload the offer roster and discard stale results when the builder team changes.
        self._clear_trade_builder_matches()
        self._refresh_trade_builder_roster()

    def _refresh_trade_builder_roster(self):
        #List every player on the selected team with their current FantasyCalc value.
        tree = self.trade_builder_give_tree
        selected_before_refresh = set(tree.selection())
        for item in self.trade_builder_give_items:
            if tree.exists(item):
                tree.delete(item)
        for item in tree.get_children():
            tree.delete(item)
        self.trade_builder_give_items = []
        if self.dataframes is None:
            return
        roster_id = self._selected_roster_id(self.trade_builder_team_combo)
        if roster_id is None:
            return
        rows = self.dataframes["Rosters"]
        rows = rows[rows["Roster ID"].astype(str) == str(roster_id)]
        position_order = {position: index for index, position in enumerate(("QB", "RB", "WR", "TE", "K", "DEF", "DST"))}
        ordered_rows = list(rows.iterrows())
        ordered_rows.sort(key=lambda item: (
            2 if str(item[1].get("Slot", "")).casefold() in {"reserve", "ir"} else
            1 if str(item[1].get("Slot", "")).casefold() == "bench" else 0,
            item[1].get("Lineup Order") if pd.notna(item[1].get("Lineup Order")) else
            item[1].get("Roster Order", 999),
            position_order.get(str(item[1].get("Position", "")).split(",")[0].strip().upper(), 99),
        ))
        for _, row in ordered_rows:
            player_id = str(row["Player ID"])
            value = self.fantasycalc_values.get(player_id)
            shown_value = f"{value:,.0f}" if value is not None else "N/A"
            self.trade_builder_give_tree.insert(
                "", "end", iid=player_id, text=str(row["Player"]), tags=(player_id,),
                values=(row.get("Position", ""), shown_value, row.get("Slot", "")),
            )
            self.trade_builder_give_items.append(player_id)
            if player_id in selected_before_refresh:
                self.trade_builder_give_tree.selection_add(player_id)
        self._filter_trade_builder_roster()
        self._update_trade_builder_selection()

    def _filter_trade_builder_roster(self):
        #Search player name, position, and lineup slot while retaining selected assets.
        if not hasattr(self, "trade_builder_give_tree"):
            return
        query = self.trade_builder_search_var.get().strip().casefold()
        for player_id in self.trade_builder_give_items:
            if not self.trade_builder_give_tree.exists(player_id):
                continue
            searchable = " ".join((
                str(self.trade_builder_give_tree.item(player_id, "text")),
                *(str(value) for value in self.trade_builder_give_tree.item(player_id, "values")),
            )).casefold()
            if not query or query in searchable:
                self.trade_builder_give_tree.move(player_id, "", "end")
            else:
                self.trade_builder_give_tree.detach(player_id)

    def _refresh_trade_builder_values(self):
        #Refresh displayed market values without clearing the builder's player selection.
        if not hasattr(self, "trade_builder_give_tree"):
            return
        for player_id in self.trade_builder_give_items:
            if not self.trade_builder_give_tree.exists(player_id):
                continue
            values = list(self.trade_builder_give_tree.item(player_id, "values"))
            value = self.fantasycalc_values.get(str(player_id))
            if len(values) > 1:
                values[1] = f"{value:,.0f}" if value is not None else "N/A"
                self.trade_builder_give_tree.item(player_id, values=values)
        self._update_trade_builder_selection()

    def _update_trade_builder_selection(self, _event=None):
        #Summarize the offer and invalidate results tied to the previous selection.
        if hasattr(self, "trade_builder_results_tree"):
            self._clear_trade_builder_matches("Offer changed. Find Matches to refresh suggestions.")
        selected = self.trade_builder_give_tree.selection()
        if not selected:
            self.trade_builder_selection_var.set("Select one or more players to send.")
            return
        missing = [iid for iid in selected if str(iid) not in self.fantasycalc_values]
        if missing:
            self.trade_builder_selection_var.set(
                f"{len(selected)} selected · value unavailable for {len(missing)} player(s)."
            )
            return
        total = sum(self.fantasycalc_values[str(iid)] for iid in selected)
        self.trade_builder_selection_var.set(f"{len(selected)} selected · total FC value {total:,.0f}.")

    def _build_trade_suggestions(self):
        #Search one-to-three-player returns across opponent rosters by market-value distance.
        tree = self.trade_builder_results_tree
        for item in tree.get_children():
            tree.delete(item)
        self.trade_builder_matches.clear()
        if self.dataframes is None:
            self.trade_builder_summary.set("Load a league before building a trade.")
            return
        selected = self.trade_builder_give_tree.selection()
        if not selected:
            self.trade_builder_summary.set("Select at least one player you would send.")
            return
        missing = [iid for iid in selected if str(iid) not in self.fantasycalc_values]
        if missing:
            self.trade_builder_summary.set("Wait for FantasyCalc values to load before searching.")
            return
        own_roster_id = self._selected_roster_id(self.trade_builder_team_combo)
        if own_roster_id is None:
            self.trade_builder_summary.set("Choose your team first.")
            return
        give_assets = []
        for player_id in selected:
            row = self.dataframes["Rosters"]
            row = row[(row["Roster ID"].astype(str) == str(own_roster_id)) &
                      (row["Player ID"].astype(str) == str(player_id))]
            if row.empty:
                continue
            player = row.iloc[0]
            give_assets.append({
                "id": str(player_id), "name": str(player["Player"]),
                "position": str(player["Position"]),
                "value": self.fantasycalc_values[str(player_id)],
            })
        if not give_assets:
            self.trade_builder_summary.set("The selected players are no longer on this roster.")
            return
        give_value = sum(asset["value"] for asset in give_assets)
        if give_value <= 0:
            self.trade_builder_summary.set("Selected players have no positive FantasyCalc value.")
            return
        try:
            tolerance = float(self.trade_builder_tolerance_var.get().rstrip("%")) / 100
        except ValueError:
            tolerance = 0.15
        needs_data = self.dataframes.get("Position Needs", pd.DataFrame())

        def needs_for(roster_id):
            if needs_data.empty or "Roster ID" not in needs_data:
                return set()
            team_needs = needs_data[needs_data["Roster ID"].astype(str) == str(roster_id)]
            return {str(row["Position"]).upper() for _, row in team_needs.iterrows()
                    if str(row.get("Assessment", "")) != "Covered"}

        own_needs = needs_for(own_roster_id)
        give_positions = {part.strip().upper() for asset in give_assets
                          for part in asset["position"].split(",") if part.strip()}
        roster_rows = self.dataframes["Rosters"]
        matches = []
        for partner, partner_roster_id in self.team_options:
            if str(partner_roster_id) == str(own_roster_id):
                continue
            team_rows = roster_rows[roster_rows["Roster ID"].astype(str) == str(partner_roster_id)]
            candidates = []
            for _, player in team_rows.iterrows():
                player_id = str(player["Player ID"])
                value = self.fantasycalc_values.get(player_id)
                if value is None or value <= 0:
                    continue
                candidates.append({
                    "id": player_id, "name": str(player["Player"]),
                    "position": str(player["Position"]), "value": value,
                })
            candidates.sort(key=lambda asset: (-asset["value"], asset["name"].casefold()))
            partner_needs = needs_for(partner_roster_id)
            for count in range(1, min(3, len(candidates)) + 1):
                for package in itertools.combinations(candidates, count):
                    receive_value = sum(asset["value"] for asset in package)
                    gap = abs(receive_value - give_value) / max(receive_value, give_value)
                    if gap > tolerance:
                        continue
                    receive_positions = {part.strip().upper() for asset in package
                                         for part in asset["position"].split(",") if part.strip()}
                    offer_fit = bool(give_positions & partner_needs)
                    receive_fit = bool(receive_positions & own_needs)
                    if offer_fit and receive_fit:
                        fit, fit_order = "Both sides", 0
                    elif offer_fit or receive_fit:
                        fit, fit_order = "One side", 1
                    else:
                        fit, fit_order = "Value only", 2
                    matches.append({
                        "id": uuid.uuid4().hex,
                        "partner": partner, "partner_roster_id": str(partner_roster_id),
                        "own_roster_id": str(own_roster_id), "give_assets": give_assets,
                        "receive_assets": list(package), "give_value": give_value,
                        "receive_value": receive_value, "gap": gap,
                        "fit": fit, "fit_order": fit_order,
                    })
        matches.sort(key=lambda match: (
            match["gap"], match["fit_order"], len(match["receive_assets"]),
            match["partner"].casefold(),
        ))
        self.trade_builder_all_matches = matches
        self.trade_builder_search_context = {"tolerance": tolerance, "give_value": give_value}
        assets = [asset for match in matches for asset in match["receive_assets"]]
        self._set_trade_builder_filter_values(
            self.trade_builder_position_filter,
            sorted({position for asset in assets for position in asset["position"].split(",") if position.strip()}),
        )
        self._filter_trade_builder_matches()

    @staticmethod
    def _set_trade_builder_filter_values(combo, options):
        #Refresh a builder filter while retaining its selection when still available.
        current = combo.get() or "Any"
        values = ["Any", *options]
        combo.configure(values=values)
        combo.set(current if current in values else "Any")

    def _filter_trade_builder_matches(self, _event=None):
        #Filter packages by a return player's position and number of players received.
        tree = self.trade_builder_results_tree
        for item in tree.get_children():
            tree.delete(item)
        self.trade_builder_matches.clear()
        position = self.trade_builder_position_filter.get()
        package_size = self.trade_builder_package_size_filter.get()
        filtered = []
        for match in self.trade_builder_all_matches:
            matches_size = package_size == "Any" or len(match["receive_assets"]) == int(package_size)
            matches_position = position == "Any" or any(
                position.upper() in {part.strip().upper() for part in asset["position"].split(",")}
                for asset in match["receive_assets"]
            )
            if matches_size and matches_position:
                filtered.append(match)

        partner_groups = {}
        for match in filtered[:200]:
            group_id = partner_groups.get(match["partner"])
            if group_id is None:
                group_id = f"builder-partner-{len(partner_groups)}"
                partner_groups[match["partner"]] = group_id
                tree.insert("", "end", iid=group_id, text=match["partner"], open=True)
            iid = match["id"]
            players = " + ".join(asset["name"] for asset in match["receive_assets"])
            positions = " + ".join(asset["position"] for asset in match["receive_assets"])
            difference = match["receive_value"] - match["give_value"]
            tree.insert(
                group_id, "end", iid=iid,
                text=f"{len(match['receive_assets'])}-player return",
                values=(players, positions, f"{match['receive_value']:,.0f}",
                        f"{difference:+,.0f} ({match['gap']:.0%})", match["fit"]),
            )
            self.trade_builder_matches[iid] = match
        shown = min(len(filtered), 200)
        context = getattr(self, "trade_builder_search_context", {})
        tolerance = context.get("tolerance", 0.15)
        give_value = context.get("give_value", 0)
        if filtered:
            self.trade_builder_summary.set(
                f"Showing {shown} of {len(filtered)} matching package(s) within {tolerance:.0%} of "
                f"{give_value:,.0f}. Expand a team to see its return packages."
            )
        elif self.trade_builder_all_matches:
            self.trade_builder_summary.set("No return packages match these filters. Choose Any or adjust the filters.")
        elif context:
            self.trade_builder_summary.set(
                f"No return packages within {tolerance:.0%} of {give_value:,.0f}. Try a wider value gap."
            )

    def _reset_trade_builder_filters(self):
        #Restore both builder filters and display every available match.
        self.trade_builder_position_filter.set("Any")
        self.trade_builder_package_size_filter.set("Any")
        self._filter_trade_builder_matches()

    def _open_trade_builder_match(self):
        #Load the selected package into Trade Review so it can be edited and assessed.
        selection = self.trade_builder_results_tree.selection()
        if not selection:
            messagebox.showinfo("Choose a Match", "Select a suggested return package first.")
            return
        match = self.trade_builder_matches.get(selection[0])
        if not match:
            return
        left_index = next((index for index, (_label, roster_id) in enumerate(self.team_options)
                           if str(roster_id) == match["own_roster_id"]), None)
        right_index = next((index for index, (_label, roster_id) in enumerate(self.team_options)
                            if str(roster_id) == match["partner_roster_id"]), None)
        if left_index is None or right_index is None:
            return
        self._new_trade()
        self.trade_left_team.current(left_index)
        self.trade_right_team.current(right_index)
        self._populate_trade_rosters()
        wanted = (
            (self.trade_left_tree, {asset["id"] for asset in match["give_assets"]}),
            (self.trade_right_tree, {asset["id"] for asset in match["receive_assets"]}),
        )
        for trade_tree, player_ids in wanted:
            for iid in trade_tree.get_children():
                tags = trade_tree.item(iid, "tags")
                if tags and str(tags[0]) in player_ids:
                    trade_tree.selection_add(iid)
        self.trade_note_var.set(f"Trade Builder suggestion · {match['partner']} · {match['gap']:.0%} value gap")
        self.workspace.select(self.trade_tab)
        self._update_trade_review()

    def _refresh_trade_ideas(self, _event=None):
        #Find reciprocal depth fits and rank candidate one-for-one deals by value gap.
        if self.dataframes is None or not hasattr(self, "trade_ideas_tree"):
            return
        for item in self.trade_ideas_tree.get_children():
            self.trade_ideas_tree.delete(item)
        self.trade_ideas_by_id.clear()
        own_roster_id = self._selected_roster_id(self.suggestion_team_combo)
        if own_roster_id is None:
            self.trade_ideas_summary.set("Choose your team to find potential trade fits.")
            return

        needs_data = self.dataframes.get("Position Needs", pd.DataFrame())
        roster_data = self.dataframes["Rosters"]
        relevant_positions = {"QB", "RB", "WR", "TE"}
        if needs_data.empty or "Roster ID" not in needs_data.columns:
            self.trade_ideas_summary.set("No lineup position requirements were found for this league.")
            return

        #A need is a position currently flagged as a starter gap or thin roster depth.
        def team_needs(roster_id):
            frame = needs_data[needs_data["Roster ID"].astype(str) == str(roster_id)]
            return {str(row["Position"]) for _, row in frame.iterrows()
                    if row["Position"] in relevant_positions and row["Assessment"] != "Covered"}

        #A surplus is an extra rostered player beyond the required starter count.
        def team_surplus(roster_id):
            frame = needs_data[needs_data["Roster ID"].astype(str) == str(roster_id)]
            return {str(row["Position"]) for _, row in frame.iterrows()
                    if row["Position"] in relevant_positions and row["Rostered"] > row["Required"]}

        #Only suggest bench assets so the app does not immediately recommend trading starters.
        def bench_assets(roster_id, position):
            frame = roster_data[
                (roster_data["Roster ID"].astype(str) == str(roster_id)) &
                (roster_data["Slot"].astype(str).str.casefold() == "bench")
            ]
            assets = []
            for _, player in frame.iterrows():
                eligible = {part.strip().upper() for part in str(player["Position"]).split(",")}
                if position in eligible:
                    player_id = str(player["Player ID"])
                    assets.append({
                        "id": player_id,
                        "name": str(player["Player"]),
                        "position": str(player["Position"]),
                        "value": self.fantasycalc_values.get(player_id),
                    })
            return sorted(assets, key=lambda asset: (asset["value"] is None, -(asset["value"] or 0), asset["name"].casefold()))[:5]

        our_needs = team_needs(own_roster_id)
        our_surplus = team_surplus(own_roster_id)
        ideas = []
        for partner_label, partner_roster_id in self.team_options:
            if str(partner_roster_id) == str(own_roster_id):
                continue
            their_needs = team_needs(partner_roster_id)
            their_surplus = team_surplus(partner_roster_id)
            offer_positions = sorted(our_surplus & their_needs)
            target_positions = sorted(our_needs & their_surplus)
            for offer_position in offer_positions:
                offers = bench_assets(own_roster_id, offer_position)
                for target_position in target_positions:
                    targets = bench_assets(partner_roster_id, target_position)
                    for offer in offers:
                        for target in targets:
                            offer_value = offer["value"]
                            target_value = target["value"]
                            gap = None
                            if offer_value is not None and target_value is not None:
                                denominator = max(offer_value, target_value)
                                gap = abs(offer_value - target_value) / denominator if denominator else 0
                            if gap is None:
                                market_fit = "Values unavailable"
                            elif gap <= 0.10:
                                market_fit = "Close"
                            elif gap <= 0.25:
                                market_fit = "Some gap"
                            else:
                                market_fit = "Large gap"
                            ideas.append({
                                "partner": partner_label,
                                "partner_roster_id": partner_roster_id,
                                "offer_roster_id": own_roster_id,
                                "offer": offer,
                                "target": target,
                                "match": f"They need {offer_position} · You need {target_position}",
                                "gap": gap,
                                "market_fit": market_fit,
                            })

        ideas.sort(key=lambda idea: (idea["gap"] is None,
                                    idea["gap"] if idea["gap"] is not None else 1,
                                    idea["partner"].casefold(),
                                    idea["offer"]["name"].casefold()))
        partner_groups = {}
        for idea in ideas[:100]:
            item_id = uuid.uuid4().hex
            offer_value = idea["offer"]["value"]
            target_value = idea["target"]["value"]
            group_id = partner_groups.get(idea["partner"])
            if group_id is None:
                group_id = f"trade-partner-{len(partner_groups)}"
                partner_groups[idea["partner"]] = group_id
                self.trade_ideas_tree.insert(
                    "", "end", iid=group_id, text=idea["partner"], open=True,
                    values=("", "", "", "", "", ""), tags=("partner",),
                )
            self.trade_ideas_tree.insert(
                group_id, "end", iid=item_id,
                text=f"{idea['offer']['name']} for {idea['target']['name']}",
                values=(idea["match"], idea["offer"]["name"],
                        f"{offer_value:,.0f}" if offer_value is not None else "N/A",
                        idea["target"]["name"],
                        f"{target_value:,.0f}" if target_value is not None else "N/A",
                        idea["market_fit"]),
            )
            self.trade_ideas_by_id[item_id] = idea
        self.trade_ideas_summary.set(
            f"Showing {min(len(ideas), 100)} potential one-for-one fit(s) across {len(partner_groups)} team(s). "
            "Open an idea to adjust it and review the full trade."
            if ideas else
            "No reciprocal bench-depth fits found. Check position needs or try a different team."
        )

    def _open_selected_trade_idea(self):
        #Load a selected idea into Trade Review so the user can adjust it before saving.
        selection = self.trade_ideas_tree.selection()
        if not selection:
            messagebox.showinfo("Choose an Idea", "Select a trade idea to open in Trade Review.")
            return
        idea = self.trade_ideas_by_id.get(selection[0])
        if not idea:
            return
        left_index = next((i for i, (_label, roster_id) in enumerate(self.team_options)
                           if str(roster_id) == str(idea["offer_roster_id"])), None)
        right_index = next((i for i, (_label, roster_id) in enumerate(self.team_options)
                            if str(roster_id) == str(idea["partner_roster_id"])), None)
        if left_index is None or right_index is None:
            return
        self._new_trade()
        self.trade_left_team.current(left_index)
        self.trade_right_team.current(right_index)
        self._populate_trade_rosters()
        for tree, player_id in ((self.trade_left_tree, idea["offer"]["id"]),
                                (self.trade_right_tree, idea["target"]["id"])):
            for iid in tree.get_children():
                tags = tree.item(iid, "tags")
                if tags and str(tags[0]) == str(player_id):
                    tree.selection_add(iid)
                    tree.see(iid)
                    break
        self.trade_note_var.set(f"Trade idea: {idea['match']} with {idea['partner']}")
        self.workspace.select(self.trade_tab)
        self._update_trade_review()

    def _load_saved_trades(self):
        #Read proposals saved for each league; absent or invalid data starts as an empty list.
        try:
            with SAVED_TRADES_FILE.open("r", encoding="utf-8") as source:
                saved = json.load(source)
                return saved if isinstance(saved, dict) else {}
        except (OSError, ValueError, AttributeError):
            return {}

    def _persist_saved_trades(self):
        #Write all saved proposals to the local settings folder.
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        with SAVED_TRADES_FILE.open("w", encoding="utf-8") as destination:
            json.dump(self.saved_trades, destination, indent=2)

    def _current_league_id(self):
        #Use the selected league ID as the key for that league's saved proposals.
        league = self.selected_league()
        if league:
            return str(league.get("league_id", ""))
        if self.dataframes is not None:
            return str(self.dataframes["League Info"].iloc[0].get("League ID", ""))
        return ""

    def _export_saved_trades(self):
        #Save the current league's complete saved-trade tree as a portable JSON file.
        league_id = self._current_league_id()
        if not league_id or self.dataframes is None:
            messagebox.showwarning("No League Loaded", "Load a league before exporting saved trades.")
            return
        records = self._saved_trade_records()
        if not records:
            messagebox.showinfo("No Saved Trades", "There are no saved trades to export for this league.")
            return
        league = self.selected_league() or {}
        league_info = self.dataframes["League Info"].iloc[0]
        payload = {
            "format": "sleeper-fantasy-calculator-saved-trades",
            "version": 1,
            "exported_at": datetime.now().astimezone().isoformat(timespec="minutes"),
            "league": {
                "id": league_id,
                "name": str(league.get("name") or league_info.get("League Name", "")),
                "season": str(league.get("season", league_info.get("Season", ""))),
            },
            "trades": records,
        }
        path = filedialog.asksaveasfilename(
            title="Export Saved Trades", defaultextension=".json",
            initialfile=f"saved-trades-{league_id}.json",
            filetypes=(("JSON files", "*.json"), ("All files", "*.*")),
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as destination:
                json.dump(payload, destination, ensure_ascii=False, indent=2)
        except OSError as exc:
            messagebox.showerror("Export Failed", f"Could not write the saved-trades file.\n\n{exc}")
            return
        self.status(f"Exported {len(records)} saved trade(s).")

    def _import_saved_trades(self):
        #Merge a saved-trades file into the active league without replacing local records.
        league_id = self._current_league_id()
        if not league_id or self.dataframes is None:
            messagebox.showwarning("No League Loaded", "Load the destination league before importing trades.")
            return
        path = filedialog.askopenfilename(
            title="Import Saved Trades",
            filetypes=(("Saved trade JSON", "*.json"), ("All files", "*.*")),
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as source:
                payload = json.load(source)
        except (OSError, ValueError, UnicodeError) as exc:
            messagebox.showerror("Import Failed", f"Could not read that JSON file.\n\n{exc}")
            return
        if (not isinstance(payload, dict) or
                payload.get("format") != "sleeper-fantasy-calculator-saved-trades" or
                payload.get("version") != 1 or not isinstance(payload.get("trades"), list) or
                not isinstance(payload.get("league"), dict)):
            messagebox.showerror("Unsupported File", "Choose a saved-trades JSON export from this app.")
            return
        source_league = payload.get("league") or {}
        source_id = str(source_league.get("id", ""))
        if source_id and source_id != league_id:
            label = source_league.get("name") or source_id
            if not messagebox.askyesno(
                "League Does Not Match",
                f"This file was exported from {label}. Import those trades into the selected league?",
            ):
                return

        def valid_trade(trade):
            if not isinstance(trade, dict):
                return False
            for side in ("give_players", "get_players"):
                assets = trade.get(side)
                if not isinstance(assets, list) or not assets or any(
                    not isinstance(asset, dict) or not asset.get("id") or not asset.get("name")
                    for asset in assets
                ):
                    return False
            children = trade.get("counter_offers", [])
            return isinstance(children, list) and all(valid_trade(child) for child in children)

        imported = payload["trades"]
        if not all(valid_trade(trade) for trade in imported):
            messagebox.showerror("Invalid Trade Data", "The file contains a trade with invalid player data.")
            return
        current = self.saved_trades.setdefault(league_id, [])
        known_ids = set()

        def collect_ids(records):
            for record in records:
                if record.get("id"):
                    known_ids.add(str(record["id"]))
                collect_ids(record.get("counter_offers", []))

        collect_ids(current)
        used_ids = set(known_ids)
        imported_count = 0
        skipped_count = 0

        def clone_trade(record, parent_id=None):
            nonlocal imported_count, skipped_count
            record_id = str(record.get("id") or uuid.uuid4().hex)
            if record_id in used_ids:
                skipped_count += 1
                return None
            used_ids.add(record_id)
            cloned = dict(record)
            cloned["id"] = record_id
            cloned["league_id"] = league_id
            if parent_id:
                cloned["parent_trade_id"] = parent_id
            else:
                cloned.pop("parent_trade_id", None)
            cloned["counter_offers"] = []
            imported_count += 1
            for child in record.get("counter_offers", []):
                imported_child = clone_trade(child, record_id)
                if imported_child is not None:
                    cloned["counter_offers"].append(imported_child)
            return cloned

        merged = []
        for record in imported:
            cloned = clone_trade(record)
            if cloned is not None:
                merged.append(cloned)
        if not merged:
            messagebox.showinfo("Nothing to Import", "Every trade in this file is already present.")
            return
        current.extend(merged)
        try:
            self._persist_saved_trades()
        except OSError as exc:
            del current[-len(merged):]
            messagebox.showerror("Import Failed", f"Could not save the imported trades.\n\n{exc}")
            return
        self._refresh_saved_trades_view()
        duplicate_note = f" Skipped {skipped_count} duplicate trade(s)." if skipped_count else ""
        self.status(f"Imported {imported_count} trade(s).{duplicate_note}")

    def _saved_trade_records(self):
        #Return only the saved proposals for the currently selected league.
        return self.saved_trades.get(self._current_league_id(), [])

    def _saved_trade_side_value(self, assets):
        #Sum the latest available market values and report whether any asset lacks a value.
        total = 0.0
        complete = True
        for asset in assets:
            value = self.fantasycalc_values.get(str(asset.get("id", "")))
            if value is None:
                complete = False
            else:
                total += value
        return total, complete

    @staticmethod
    def _market_lean(give_value, receive_value, complete=True):
        #Describe the value comparison with a small tolerance instead of a yes/no verdict.
        if not complete:
            return "Values unavailable"
        difference = receive_value - give_value
        if abs(difference) <= max(give_value, receive_value) * 0.10:
            return "Values are close"
        if difference > 0:
            return f"Leans to you +{difference:,.0f}"
        return f"Leans away {abs(difference):,.0f}"

    def _refresh_saved_trades_view(self, select_id=None):
        #Rebuild the saved-trade list using the newest cached or downloaded market values.
        if not hasattr(self, "saved_trade_tree"):
            return
        if select_id is None:
            previous = self.saved_trade_tree.selection()
            if previous:
                select_id = previous[0]
        for item in self.saved_trade_tree.get_children():
            self.saved_trade_tree.delete(item)
        records = self._saved_trade_records()
        opponent_groups = {}
        for trade in records:
            opponent = trade.get("get_team", "Other team")
            if opponent not in opponent_groups:
                group_id = f"opponent-group-{len(opponent_groups)}"
                opponent_groups[opponent] = group_id
                self.saved_trade_tree.insert(
                    "", "end", iid=group_id, text=f"{opponent}", open=True
                )
            self._insert_saved_trade_tree_item(trade, opponent_groups[opponent])
        if select_id and self.saved_trade_tree.exists(str(select_id)):
            self.saved_trade_tree.selection_set(str(select_id))
            self.saved_trade_tree.see(str(select_id))
            self._show_saved_trade_details()
        elif not records:
            self.saved_trade_details.set("No saved trades for this league yet. Build a trade in Trade Review and choose Save Trade.")
        else:
            self.saved_trade_details.set("Select a saved trade to see its players and note.")

    def _insert_saved_trade_tree_item(self, trade, parent_iid):
        #Show counteroffers as child rows under the original trade, grouped by opposing team above.
        give_value, give_complete = self._saved_trade_side_value(trade.get("give_players", []))
        receive_value, receive_complete = self._saved_trade_side_value(trade.get("get_players", []))
        complete = give_complete and receive_complete
        format_value = lambda amount, ready: f"{amount:,.0f}" if ready else "N/A"
        give_team = trade.get("give_team", "Your team")
        get_team = trade.get("get_team", "Other team")
        updated = trade.get("updated_at") or trade.get("saved_at", "")
        if "T" in updated:
            updated = updated.replace("T", " ")
        row_label = f"{give_team}  ⇄  {get_team}"
        if trade.get("parent_trade_id"):
            row_label = f"Counteroffer: {row_label}"
        self.saved_trade_tree.insert(
            parent_iid, "end", iid=str(trade["id"]),
            text=row_label,
            values=(trade.get("outcome", "Proposed"),
                    format_value(give_value, give_complete),
                    format_value(receive_value, receive_complete),
                    self._market_lean(give_value, receive_value, complete),
                    trade.get("note", ""), updated),
            open=True,
        )
        for counteroffer in trade.get("counter_offers", []):
            self._insert_saved_trade_tree_item(counteroffer, str(trade["id"]))

    def _find_saved_trade(self, trade_id):
        #Find a top-level trade or any attached counteroffer by its stable local ID.
        def find(records):
            for trade in records:
                if str(trade.get("id")) == str(trade_id):
                    return trade
                nested = find(trade.get("counter_offers", []))
                if nested:
                    return nested
            return None
        return find(self._saved_trade_records())

    def _locate_saved_trade(self, trade_id, records=None, parent=None):
        #Return the list that directly owns a saved trade so top-level and nested entries can be edited.
        records = self._saved_trade_records() if records is None else records
        for index, trade in enumerate(records):
            if str(trade.get("id")) == str(trade_id):
                return records, index, trade, parent
            located = self._locate_saved_trade(
                trade_id, trade.get("counter_offers", []), parent=trade
            )
            if located:
                return located
        return None

    def _mark_saved_trade_outcome(self, outcome):
        #Update the selected saved proposal's outcome without changing its players or note.
        selection = self.saved_trade_tree.selection()
        if not selection:
            messagebox.showinfo("Choose a Trade", f"Select a saved trade to mark {outcome.lower()}.")
            return
        trade = self._find_saved_trade(selection[0])
        if not trade:
            return
        trade["outcome"] = outcome
        trade["updated_at"] = datetime.now().astimezone().isoformat(timespec="minutes")
        try:
            self._persist_saved_trades()
        except OSError as exc:
            messagebox.showerror("Could Not Update Trade", f"The saved trade could not be updated.\n\n{exc}")
            return
        self._refresh_saved_trades_view(select_id=trade["id"])
        self.status(f"Saved trade marked {outcome.lower()}.")

    def _show_saved_trade_details(self, _event=None):
        #Show each saved player, current value, side totals, and the user's note.
        selection = self.saved_trade_tree.selection()
        if not selection:
            self.copy_outcome_button.configure(state="disabled", text="Copy Proposed Text")
            return
        trade = self._find_saved_trade(selection[0])
        if not trade:
            self.copy_outcome_button.configure(state="disabled", text="Copy Proposed Text")
            self.saved_trade_details.set("Select a trade under an opposing team to see its players and note.")
            return
        outcome = trade.get("outcome", "Proposed")
        self.copy_outcome_button.configure(state="normal", text=f"Copy {outcome} Text")
        def describe(side):
            lines = []
            total, complete = self._saved_trade_side_value(trade.get(side, []))
            for asset in trade.get(side, []):
                value = self.fantasycalc_values.get(str(asset.get("id", "")))
                shown = f"{value:,.0f}" if value is not None else "N/A"
                lines.append(f"{asset.get('name', 'Unknown player')} ({asset.get('position', '?')}) — {shown}")
            total_text = f"{total:,.0f}" if complete else "N/A"
            return "\n".join(lines) or "No players selected", total_text
        give_lines, give_total = describe("give_players")
        get_lines, get_total = describe("get_players")
        self.saved_trade_details.set(
            f"Outcome: {outcome}\n"
            f"Attached counteroffers: {len(trade.get('counter_offers', []))}\n"
            f"Note: {trade.get('note') or '—'}\n\n"
            f"You give ({trade.get('give_team', 'Your team')}) — {give_total}\n{give_lines}\n\n"
            f"You receive ({trade.get('get_team', 'Other team')}) — {get_total}\n{get_lines}"
        )

    def _selected_trade_assets(self, tree):
        #Capture the selected player IDs and display details for saving or simulating a trade.
        assets = []
        for iid in tree.selection():
            tags = tree.item(iid, "tags")
            if not tags:
                continue
            position = tree.item(iid, "values")[0]
            player_id = str(tags[0])
            assets.append({"id": player_id, "name": tree.item(iid, "text"), "position": position})
        return assets

    def _sync_save_as_new_trade_button(self):
        #Only offer a duplicate action when Save Trade is updating an existing record.
        if self.active_saved_trade_id:
            if not self.save_as_new_trade_button.winfo_manager():
                self.save_as_new_trade_button.pack(side="left", padx=8, before=self.new_trade_button)
        else:
            self.save_as_new_trade_button.pack_forget()

    def _save_current_trade(self, outcome=None, save_as_new=False):
        #Save or update a trade while preserving its outcome unless the user changes it.
        if self.dataframes is None:
            messagebox.showwarning("No League Loaded", "Load a league before saving a trade.")
            return
        give_players = self._selected_trade_assets(self.trade_left_tree)
        get_players = self._selected_trade_assets(self.trade_right_tree)
        if not give_players or not get_players:
            messagebox.showwarning("Incomplete Trade", "Select at least one player on both sides before saving.")
            return
        now = datetime.now().astimezone().isoformat(timespec="minutes")
        counter_parent = (
            self._find_saved_trade(self.active_counteroffer_parent_id)
            if self.active_counteroffer_parent_id and not save_as_new else None
        )
        if self.active_counteroffer_parent_id and not save_as_new and not counter_parent:
            messagebox.showerror("Counteroffer Unavailable", "The saved trade this counteroffer belongs to could not be found.")
            return
        is_counteroffer = counter_parent is not None
        record_id = uuid.uuid4().hex if save_as_new else (self.active_saved_trade_id or uuid.uuid4().hex)
        located = self._locate_saved_trade(record_id) if not save_as_new else None
        existing_trade = located[2] if located else None
        if outcome:
            saved_outcome = outcome
        elif existing_trade and str(record_id) == str(self.active_saved_trade_id) and not save_as_new:
            saved_outcome = self._current_trade_outcome()
        else:
            saved_outcome = "Proposed"
        record = {
            "id": record_id,
            "league_id": self._current_league_id(),
            "outcome": saved_outcome,
            "give_roster_id": str(self._selected_roster_id(self.trade_left_team)),
            "get_roster_id": str(self._selected_roster_id(self.trade_right_team)),
            "give_team": self.trade_left_team.get(),
            "get_team": self.trade_right_team.get(),
            "give_players": give_players,
            "get_players": get_players,
            "note": self.trade_note_var.get().strip(),
            "saved_at": now,
            "updated_at": now,
            "counter_offers": existing_trade.get("counter_offers", []) if existing_trade else [],
        }
        if existing_trade:
            record["saved_at"] = existing_trade.get("saved_at", now)
        if is_counteroffer:
            record["parent_trade_id"] = str(counter_parent["id"])
            records = counter_parent.setdefault("counter_offers", [])
        else:
            records = self.saved_trades.setdefault(record["league_id"], [])
        existing = next((index for index, item in enumerate(records)
                         if str(item.get("id")) == str(record_id)), None)
        if existing is None:
            records.append(record)
        else:
            records[existing] = record
        try:
            self._persist_saved_trades()
        except OSError as exc:
            messagebox.showerror("Could Not Save Trade", f"The trade could not be saved on this computer.\n\n{exc}")
            return
        self.active_saved_trade_id = record_id
        self.active_counteroffer_parent_id = str(counter_parent["id"]) if is_counteroffer else None
        self.save_trade_button.configure(text="Update Counteroffer" if is_counteroffer else "Update Saved Trade")
        self._sync_save_as_new_trade_button()
        self._update_trade_review()
        self._refresh_saved_trades_view(select_id=record_id)
        description = "counteroffer" if is_counteroffer else "trade"
        self.status(f"{saved_outcome} {description} saved on this computer. Its values update when market values load.")

    def _new_trade(self):
        #Clear selections and notes so the next proposal starts independently.
        self.active_saved_trade_id = None
        self.active_counteroffer_parent_id = None
        self.save_trade_button.configure(text="Save Trade")
        self._sync_save_as_new_trade_button()
        self.trade_note_var.set("")
        for tree in (self.trade_left_tree, self.trade_right_tree):
            selection = tree.selection()
            if selection:
                tree.selection_remove(*selection)
        self._update_trade_review()

    def _reopen_saved_trade(self, as_counteroffer=False):
        #Restore saved teams and player selections in the live trade reviewer.
        selection = self.saved_trade_tree.selection()
        if not selection:
            messagebox.showinfo("Choose a Trade", "Select a saved trade to reopen.")
            return
        trade = self._find_saved_trade(selection[0])
        if not trade:
            messagebox.showinfo("Choose a Trade", "Select a saved trade row under an opposing team.")
            return
        left_index = next((i for i, (_label, roster_id) in enumerate(self.team_options)
                           if str(roster_id) == str(trade.get("give_roster_id"))), None)
        right_index = next((i for i, (_label, roster_id) in enumerate(self.team_options)
                            if str(roster_id) == str(trade.get("get_roster_id"))), None)
        if left_index is None or right_index is None:
            messagebox.showwarning("Team Not Available", "One or both teams from this saved trade are not in the currently loaded league.")
            return
        self.trade_left_team.current(left_index)
        self.trade_right_team.current(right_index)
        self.trade_note_var.set(trade.get("note", ""))
        self._populate_trade_rosters()
        missing = []
        for tree, side in ((self.trade_left_tree, "give_players"), (self.trade_right_tree, "get_players")):
            wanted = {str(asset.get("id")) for asset in trade.get(side, [])}
            found = set()
            for iid in tree.get_children():
                tags = tree.item(iid, "tags")
                if tags and str(tags[0]) in wanted:
                    tree.selection_add(iid)
                    tree.see(iid)
                    found.add(str(tags[0]))
            missing.extend(wanted - found)
        if as_counteroffer:
            self.active_saved_trade_id = None
            self.active_counteroffer_parent_id = str(trade["id"])
            self.save_trade_button.configure(text="Save Counteroffer")
        else:
            self.active_saved_trade_id = str(trade["id"])
            self.active_counteroffer_parent_id = trade.get("parent_trade_id")
            self.save_trade_button.configure(
                text="Update Counteroffer" if self.active_counteroffer_parent_id else "Update Saved Trade"
            )
        self._sync_save_as_new_trade_button()
        self.workspace.select(self.trade_tab)
        self._update_trade_review()
        if missing:
            self.status(f"Some saved players are no longer on the selected rosters ({len(missing)}). Adjust the trade and update it if needed.")

    def _delete_saved_trade(self):
        #Remove the selected proposal only after the user confirms the deletion.
        selection = self.saved_trade_tree.selection()
        if not selection:
            messagebox.showinfo("Choose a Trade", "Select a saved trade to delete.")
            return
        trade = self._find_saved_trade(selection[0])
        if not trade or not messagebox.askyesno("Delete Saved Trade", "Delete this saved trade from this computer?"):
            return
        located = self._locate_saved_trade(trade.get("id"))
        if not located:
            return
        records, index, _record, _parent = located
        removed_ids = set()
        def collect_ids(item):
            removed_ids.add(str(item.get("id")))
            for child in item.get("counter_offers", []):
                collect_ids(child)
        collect_ids(trade)
        del records[index]
        try:
            self._persist_saved_trades()
        except OSError as exc:
            messagebox.showerror("Could Not Delete Trade", f"The saved trade could not be removed.\n\n{exc}")
            return
        if (self.active_saved_trade_id in removed_ids or
                self.active_counteroffer_parent_id in removed_ids):
            self.active_saved_trade_id = None
            self.active_counteroffer_parent_id = None
            self.save_trade_button.configure(text="Save Trade")
            self._sync_save_as_new_trade_button()
        self._refresh_saved_trades_view()

    def _format_changed(self, _event=None):
        #Save the Dynasty/Redraft choice and request values for that format.
        self._save_settings()
        if self.dataframes is not None:
            self._load_fantasycalc()
            self._refresh_value_trends()
            self._update_trade_review()

    def _stats_guy_format(self):
        #Select the nearest supported Stats Guy format from the league and app settings.
        league = self.selected_league() or {}
        slots = [str(slot).upper() for slot in league.get("roster_positions") or []]
        superflex = (
            "SF" in slots or "SUPER_FLEX" in slots or
            slots.count("QB") > 1
        )
        return f"{'sf' if superflex else 'non_sf'}_{'dynasty' if self.format_var.get() == 'Dynasty' else 'redraft'}"

    def _load_stats_guy_values(self):
        #Load a once-daily cached market board from an independent value source.
        def task():
            try:
                payload = self.api.get_stats_guy_values()
                players = {str(player.get("id")): player for player in payload.get("players", [])
                           if player.get("id") not in (None, "")}
                return players, payload.get("valuesAsOf") or {}, None
            except Exception as exc:
                return {}, {}, str(exc)

        def loaded(result):
            self.stats_guy_players, self.stats_guy_values_asof, self.stats_guy_error = result
            self._refresh_value_trends()
            self._update_trade_review()

        self.run_background(task, loaded)

    def _refresh_value_trends(self, _event=None):
        #Compare current FantasyCalc and Stats Guy values for the selected league players.
        if not hasattr(self, "value_trend_tree"):
            return
        for item in self.value_trend_tree.get_children():
            self.value_trend_tree.delete(item)
        self.value_trend_rows = []
        self.value_history_points = []
        self.selected_history_player_id = None
        self.value_trend_detail.set("Select a player to see volatility and recent value history.")
        self._draw_value_history()
        if self.dataframes is None:
            self.value_trend_source_status.set("Load a league to compare value sources.")
            return

        value_format = self._stats_guy_format()
        supported_format = value_format.replace("_", " ").title()
        self.value_trend_format_note.set(
            f"Stats Guy format: {supported_format}. Its API doesn't adjust for PPR, TE premium, or team count; "
            "it covers QB/RB/WR/TE. FantasyCalc uses your league's closest supported settings."
        )
        if self.stats_guy_error:
            self.value_trend_source_status.set(f"Stats Guy unavailable: {self.stats_guy_error}")
        elif not self.stats_guy_players:
            self.value_trend_source_status.set("Loading Stats Guy values…")
        else:
            as_of = self.stats_guy_values_asof.get(value_format, "date unavailable")
            self.value_trend_source_status.set(f"Stats Guy values as of {as_of}")

        selected_team = self.value_trend_team_combo.get()
        if selected_team and selected_team != "All League Players":
            roster_id = next((rid for name, rid in self.team_options if name == selected_team), None)
            roster_data = self.dataframes["Rosters"]
            roster_data = roster_data[roster_data["Roster ID"].astype(str) == str(roster_id)]
        else:
            roster_data = self.dataframes["Rosters"]

        roster_owners = {
            str(roster_id): display_name
            for display_name, roster_id in self.team_options
            if roster_id is not None
        }

        for _, row in roster_data.iterrows():
            player_id = str(row["Player ID"])
            fc_value = self.fantasycalc_values.get(player_id)
            fc_trend = self.fantasycalc_trends.get(player_id)
            stats_card = self.stats_guy_players.get(player_id, {})
            stats_values = stats_card.get("value") or {}
            stats_value = stats_values.get(value_format) if isinstance(stats_values, dict) else None
            spread = (fc_value - stats_value) if fc_value is not None and stats_value is not None else None
            self.value_trend_rows.append({
                "id": player_id,
                "name": str(row["Player"]),
                "position": str(row["Position"]),
                "rostered_by": roster_owners.get(str(row.get("Roster ID")), "Free Agent"),
                "fc": fc_value,
                "fc30": fc_trend,
                "stats_guy": stats_value,
                "spread": spread,
            })
        self.value_trend_sort_reverse = {"fc": True}
        self._sort_value_trends("fc", toggle=False)

    @staticmethod
    def _display_value(value, signed=False):
        #Show absent source data consistently and mark changes with a sign.
        if value is None:
            return "—"
        try:
            number = float(value)
        except (TypeError, ValueError):
            return "—"
        if signed:
            return f"{number:+,.0f}"
        return f"{number:,.0f}"

    def _sort_value_trends(self, column, toggle=True):
        #Sort the value comparison table while keeping missing source values at the bottom.
        if not hasattr(self, "value_trend_tree"):
            return
        if toggle:
            reverse = not self.value_trend_sort_reverse.get(column, column in {"fc", "stats_guy"})
        else:
            reverse = self.value_trend_sort_reverse.get(column, column in {"fc", "stats_guy"})
        self.value_trend_sort_reverse[column] = reverse
        rows = list(self.value_trend_rows)
        for item in self.value_trend_tree.get_children():
            self.value_trend_tree.delete(item)
        if column == "name":
            rows.sort(key=lambda row: (row["name"].rsplit(" ", 1)[-1].casefold(), row["name"].casefold()), reverse=reverse)
        elif column == "position":
            rows.sort(key=lambda row: self._position_sort_key(row["position"], row["name"].rsplit(" ", 1)[-1]), reverse=reverse)
        elif column == "rostered_by":
            rows.sort(key=lambda row: row["rostered_by"].casefold(), reverse=reverse)
        else:
            valued = [row for row in rows if row.get(column) is not None]
            missing = [row for row in rows if row.get(column) is None]
            valued.sort(key=lambda row: float(row[column]), reverse=reverse)
            rows = valued + missing
        for row in rows:
            self.value_trend_tree.insert(
                "", "end", iid=row["id"], text=row["name"], tags=(row["id"],),
                values=(row["position"], row["rostered_by"], self._display_value(row["fc"]),
                        self._display_value(row["fc30"], signed=True),
                        self._display_value(row["stats_guy"]),
                        self._display_value(row["spread"], signed=True)),
            )

    def _select_value_trend_player(self, _event=None):
        #Load the selected player's recent Stats Guy history once, then reuse it for redraws.
        selection = self.value_trend_tree.selection()
        if not selection:
            return
        player_id = str(selection[0])
        self.selected_history_player_id = player_id
        value_format = self._stats_guy_format()
        window = self._history_window_limit()
        cache_key = (player_id, value_format, window)
        if cache_key in self.value_history_cache:
            self._render_value_history(player_id, value_format, self.value_history_cache[cache_key])
            return
        player = self.stats_guy_players.get(player_id, {})
        self.value_trend_detail.set(
            f"Loading recent history for {player.get('name', 'selected player')}…"
        )
        self.value_history_points = []
        self._draw_value_history()

        def task():
            try:
                return self.api.get_stats_guy_history(player_id, value_format, window=window), None
            except Exception as exc:
                return None, str(exc)

        def loaded(result):
            history, error = result
            if error:
                self.value_trend_detail.set(f"Could not load selected-player history: {error}")
                return
            self.value_history_cache[cache_key] = history
            if (self.selected_history_player_id == player_id and
                    self._stats_guy_format() == value_format and
                    self._history_window_limit() == window):
                self._render_value_history(player_id, value_format, history)

        self.run_background(task, loaded)

    def _select_roster_history(self, _event=None):
        #Load the selected roster player's value history into the roster panel.
        selection = self.roster_tree.selection()
        if selection and self.roster_tree.parent(selection[0]):
            self._load_history_for_view(str(selection[0]), "roster")

    def _select_weekly_history(self, _event=None):
        #Load the selected waiver player's value history into the Weekly Help panel.
        selection = self.weekly_tree.selection()
        if selection:
            self._load_history_for_view(str(selection[0]), "weekly")

    def _load_history_for_view(self, player_id, view):
        #Reuse the same Stats Guy history cache and summary for each player-list tab.
        detail = self.roster_history_detail if view == "roster" else self.weekly_history_detail
        canvas = self.roster_history_canvas if view == "roster" else self.weekly_history_canvas
        points_name = "roster_history_points" if view == "roster" else "weekly_history_points"
        selected_name = "roster_history_player_id" if view == "roster" else "weekly_history_player_id"
        setattr(self, selected_name, player_id)
        card = self.stats_guy_players.get(player_id, {})
        position_text = str(card.get("position", ""))
        if not position_text and view == "roster" and self.dataframes is not None:
            roster = self.dataframes["Rosters"]
            matching = roster[roster["Player ID"].astype(str) == player_id]
            if not matching.empty:
                position_text = str(matching.iloc[0].get("Position", ""))
        elif not position_text and view == "weekly":
            player_row = next((row for row in self.weekly_waiver_rows if row["id"] == player_id), {})
            position_text = str(player_row.get("position", ""))
        supported_positions = {part.strip().upper() for part in position_text.split(",")}
        if not supported_positions.intersection({"QB", "RB", "WR", "TE"}):
            setattr(self, points_name, [])
            detail.set("Stats Guy history is available for QB, RB, WR, and TE only.")
            self._draw_history_chart(canvas, [], "No history available for this position.")
            return
        value_format = self._stats_guy_format()
        window = self._history_window_limit()
        cache_key = (player_id, value_format, window)
        if cache_key in self.value_history_cache:
            self._render_value_history(
                player_id, value_format, self.value_history_cache[cache_key], canvas, detail, points_name,
            )
            return
        detail.set(f"Loading recent value history for {card.get('name', 'selected player')}…")
        setattr(self, points_name, [])
        self._draw_history_chart(canvas, [], "Loading a 90-day trend chart…")

        def task():
            try:
                return self.api.get_stats_guy_history(player_id, value_format, window=window), None
            except Exception as exc:
                return None, str(exc)

        def loaded(result):
            history, error = result
            if getattr(self, selected_name) != player_id or self._history_window_limit() != window:
                return
            if error:
                detail.set(f"Could not load selected-player history: {error}")
                return
            self.value_history_cache[cache_key] = history
            self._render_value_history(
                player_id, value_format, history, canvas, detail, points_name,
            )

        self.run_background(task, loaded)

    def _build_history_window_control(self, parent, metric=False, action=None):
        #Offer the same time scale in each player-history panel.
        controls = ttk.Frame(parent)
        controls.pack(fill="x", anchor="w", pady=(0, 4))
        ttk.Label(controls, text="Time scale:").pack(side="left")
        selector = ttk.Combobox(
            controls, textvariable=self.value_history_window_var,
            values=tuple(VALUE_HISTORY_WINDOWS), state="readonly", width=16,
        )
        selector.pack(side="left", padx=(7, 0))
        selector.bind("<<ComboboxSelected>>", self._history_window_changed)
        if metric:
            ttk.Label(controls, text="Metric:").pack(side="left", padx=(14, 0))
            metric_selector = ttk.Combobox(
                controls, textvariable=self.trade_history_metric_var,
                values=("Percent Change", "Value"), state="readonly", width=16,
            )
            metric_selector.pack(side="left", padx=(7, 0))
            metric_selector.bind("<<ComboboxSelected>>", self._trade_history_metric_changed)
        if action is not None:
            ttk.Button(controls, text="Open Graph", command=action).pack(side="right")
        return selector

    def _refresh_player_history_source(self, source):
        #Reload the selected player associated with a graph window.
        if source == "trends":
            self._select_value_trend_player()
        elif source == "roster":
            self._select_roster_history()
        elif source == "weekly":
            self._select_weekly_history()

    def _open_player_history_graph(self, source):
        #Open one of the compact player history graphs in a larger window.
        details = {
            "roster": ("roster_history_canvas", "roster_history_points", self.roster_history_detail,
                       "Roster Player Value History", "Select a player to load a 90-day value trend."),
            "weekly": ("weekly_history_canvas", "weekly_history_points", self.weekly_history_detail,
                       "Weekly Help Player Value History", "Select a player to load a 90-day value trend."),
            "trends": ("value_history_canvas", "value_history_points", self.value_trend_detail,
                       "Player Value History", "Select a player to load a 90-day trend chart."),
        }
        if source not in details:
            return
        if self.player_history_window and self.player_history_window.winfo_exists():
            if self.player_history_popup_source == source:
                self.player_history_window.deiconify()
                self.player_history_window.lift()
                self.player_history_window.focus_force()
                return
            self._close_player_history_window()
        _canvas_name, points_name, detail, title, empty_text = details[source]
        window = tk.Toplevel(self)
        window.title(title)
        width = min(1600, max(900, int(self.winfo_screenwidth() * 0.78)))
        height = min(1000, max(560, int(self.winfo_screenheight() * 0.72)))
        window.geometry(f"{width}x{height}")
        window.minsize(720, 420)
        self.player_history_window = window
        self.player_history_popup_source = source
        self.player_history_popup_points_name = points_name
        self.player_history_popup_empty_text = empty_text
        toolbar = ttk.Frame(window, padding=(12, 10, 12, 4))
        toolbar.pack(fill="x")
        self.player_history_popup_scale = self._build_history_window_control(toolbar)
        ttk.Label(window, textvariable=detail, wraplength=1500, justify="left").pack(
            anchor="w", padx=12, pady=(0, 4)
        )
        canvas = tk.Canvas(window, height=400, highlightthickness=0,
                           bg=self.theme_colors["surface"],
                           highlightbackground=self.theme_colors["border"])
        canvas.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        canvas.bind("<Configure>", lambda _event, target=canvas, name=points_name, text=empty_text:
                    self._render_history_chart(target, getattr(self, name, []), text))
        self.player_history_popup_canvas = canvas
        window.protocol("WM_DELETE_WINDOW", self._close_player_history_window)
        self._render_history_chart(canvas, getattr(self, points_name, []), empty_text)

    def _close_player_history_window(self):
        #Release popup state so the next graph button opens the requested view.
        if self.player_history_window and self.player_history_window.winfo_exists():
            self.player_history_window.destroy()
        self.player_history_window = None
        self.player_history_popup_canvas = None
        self.player_history_popup_source = None
        self.player_history_popup_points_name = None
        self.player_history_popup_empty_text = None
        self.player_history_popup_scale = None

    def _history_window_limit(self):
        #Translate the visible range label into the number of daily snapshots to request.
        return VALUE_HISTORY_WINDOWS.get(self.value_history_window_var.get(), 91)

    def _history_window_changed(self, event=None):
        #A popup's time scale always updates the graph it opened, even after tab changes.
        if (event is not None and self.player_history_popup_scale is event.widget
                and self.player_history_window and self.player_history_window.winfo_exists()):
            self._refresh_player_history_source(self.player_history_popup_source)
            return
        #Reload only the selected player on the active tab with the new history range.
        selected_tab = self.workspace.tab(self.workspace.select(), "text")
        if selected_tab == "Value Trends":
            self._select_value_trend_player()
        elif selected_tab == "Roster View":
            self._select_roster_history()
        elif selected_tab == "Weekly Help":
            self._select_weekly_history()
        elif selected_tab == "Trade Review":
            self._refresh_trade_history(self.trade_left_tree.selection(), self.trade_right_tree.selection())
        elif self.trade_history_popup_canvas and self.trade_history_popup_canvas.winfo_exists():
            self._refresh_trade_history(self.trade_left_tree.selection(), self.trade_right_tree.selection())

    @staticmethod
    def _history_points_from_payload(payload):
        #Convert an API history response to a clean, oldest-to-newest series.
        points = []
        for point in (payload or {}).get("history", []):
            try:
                points.append((str(point["date"]), float(point["value"])))
            except (KeyError, TypeError, ValueError):
                continue
        return sorted(points, key=lambda point: point[0])

    def _refresh_trade_history(self, give_selection, get_selection):
        #Load each selected player's history once, then compare their indexed movements.
        selected = [(iid, "Give", self.trade_left_tree) for iid in give_selection]
        selected.extend((iid, "Receive", self.trade_right_tree) for iid in get_selection)
        player_ids = tuple((str(iid), side) for iid, side, _tree in selected)
        value_format = self._stats_guy_format()
        window = self._history_window_limit()
        signature = (player_ids, value_format, window)
        if signature == self.trade_history_signature:
            return
        self.trade_history_signature = signature
        self.trade_history_request_id += 1
        request_id = self.trade_history_request_id
        if not selected:
            self.trade_history_series = []
            self.trade_history_status.set("Select one or more players to compare their value trends.")
            self._draw_trade_history()
            return

        requested = []
        series = []
        unsupported = 0
        for iid, side, tree in selected:
            player_id = str(tree.item(iid, "tags")[0])
            player_name = tree.item(iid, "text")
            values = tree.item(iid, "values")
            position_text = str(values[0] if values else "").upper()
            supported = {position.strip() for position in position_text.split(",")}
            if not supported.intersection({"QB", "RB", "WR", "TE"}):
                unsupported += 1
                continue
            cache_key = (player_id, value_format, window)
            payload = self.value_history_cache.get(cache_key)
            if payload is None:
                requested.append((player_id, player_name, side, cache_key))
            else:
                points = self._history_points_from_payload(payload)
                if points:
                    series.append({"id": player_id, "name": player_name, "side": side, "points": points})

        if not requested:
            self.trade_history_series = series
            note = "Indexed movement compares each player's change from the first snapshot."
            if unsupported:
                note += f" {unsupported} selected player(s) have no Stats Guy history."
            if series:
                self.trade_history_status.set(note)
            elif unsupported:
                self.trade_history_status.set(
                    "Stats Guy history covers QB, RB, WR, and TE; selected player(s) outside those positions were skipped."
                )
            else:
                self.trade_history_status.set("No history is available for the selected players.")
            self._draw_trade_history()
            return

        self.trade_history_series = series
        self.trade_history_status.set(f"Loading value history for {len(requested)} selected player(s)…")
        self._draw_trade_history()

        def task():
            loaded, errors = [], []
            for player_id, player_name, side, cache_key in requested:
                try:
                    payload = self.api.get_stats_guy_history(player_id, value_format, window=window)
                    loaded.append((player_id, player_name, side, cache_key, payload))
                except Exception as exc:
                    errors.append(f"{player_name}: {exc}")
            return loaded, errors

        def loaded(result):
            fetched, errors = result
            if request_id != self.trade_history_request_id:
                return
            current_give = tuple((str(iid), "Give") for iid in self.trade_left_tree.selection())
            current_receive = tuple((str(iid), "Receive") for iid in self.trade_right_tree.selection())
            if ((current_give + current_receive, self._stats_guy_format(), self._history_window_limit())
                    != signature):
                return
            for player_id, player_name, side, cache_key, payload in fetched:
                self.value_history_cache[cache_key] = payload
                points = self._history_points_from_payload(payload)
                if points:
                    series.append({"id": player_id, "name": player_name, "side": side, "points": points})
            self.trade_history_series = series
            note = "Indexed movement compares each player's change from the first snapshot."
            if unsupported:
                note += f" {unsupported} selected player(s) have no Stats Guy history."
            if errors:
                note += f" Could not load {len(errors)} player history item(s)."
            if not series:
                note = "No history is available for the selected players."
            self.trade_history_status.set(note)
            self._draw_trade_history()

        self.run_background(task, loaded)

    def _open_trade_history_window(self):
        #Show the selected trade histories in a larger resizable window.
        if self.trade_history_window and self.trade_history_window.winfo_exists():
            self.trade_history_window.deiconify()
            self.trade_history_window.lift()
            self.trade_history_window.focus_force()
            return
        window = tk.Toplevel(self)
        window.title("Selected Players' Value Trends")
        width = min(1600, max(900, int(self.winfo_screenwidth() * 0.78)))
        height = min(1000, max(560, int(self.winfo_screenheight() * 0.72)))
        window.geometry(f"{width}x{height}")
        window.minsize(720, 420)
        self.trade_history_window = window
        toolbar = ttk.Frame(window, padding=(12, 10, 12, 4))
        toolbar.pack(fill="x")
        self._build_history_window_control(toolbar, metric=True)
        ttk.Label(window, textvariable=self.trade_history_status, wraplength=1500,
                  justify="left").pack(anchor="w", padx=12, pady=(0, 4))
        canvas = tk.Canvas(
            window, height=400, highlightthickness=0,
            bg=self.theme_colors["surface"], highlightbackground=self.theme_colors["border"],
        )
        canvas.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        canvas.bind("<Configure>", lambda _event, target=canvas: self._draw_trade_history(target))
        canvas.bind("<Motion>", lambda event, target=canvas: self._show_trade_history_hover(event, target))
        canvas.bind("<Leave>", lambda _event, target=canvas: self._clear_trade_history_hover(target))
        self.trade_history_popup_canvas = canvas
        window.protocol("WM_DELETE_WINDOW", self._close_trade_history_window)
        self._draw_trade_history(canvas)

    def _close_trade_history_window(self):
        #Release the popup references so reopening creates a fresh graph window.
        if self.trade_history_window and self.trade_history_window.winfo_exists():
            self.trade_history_window.destroy()
        self.trade_history_window = None
        self.trade_history_popup_canvas = None

    def _draw_trade_history(self, target_canvas=None):
        #Overlay selected players as indexed percentage changes on a common axis.
        if target_canvas is None and not hasattr(self, "trade_history_canvas"):
            return
        canvases = (
            [target_canvas] if target_canvas is not None else
            [self.trade_history_canvas] + (
                [self.trade_history_popup_canvas]
                if self.trade_history_popup_canvas and self.trade_history_popup_canvas.winfo_exists() else []
            )
        )
        for canvas in canvases:
            self.trade_history_plot_points[str(canvas)] = []
            self._draw_trade_history_canvas(canvas)

    def _trade_history_metric_changed(self, _event=None):
        #Switch axes locally without reloading the selected player histories.
        self._draw_trade_history()

    def _draw_trade_history_canvas(self, canvas):
        #Render one canvas so inline and popup graphs always use the same selected data.
        canvas.delete("all")
        canvas.configure(bg=self.theme_colors["surface"], highlightbackground=self.theme_colors["border"])
        series = [item for item in self.trade_history_series if item.get("points")]
        if not series:
            canvas.create_text(12, 18, anchor="nw", text="Select players in Trade Review to compare their histories.",
                               fill=self.theme_colors["muted"])
            return
        colors = (
            ("#66c2ff", "#ffad70", "#65d69a", "#d28cff", "#ff7285", "#73d5e8",
             "#e4c45e", "#8b9dff", "#b7d968", "#ff8fd0", "#d0d0d0", "#50d7ca")
            if self.dark_mode_var.get() else
            ("#2878b5", "#d06b28", "#2f8f5b", "#a14ca6", "#c44252", "#558c9e",
             "#97712f", "#5266b3", "#6f7d3b", "#cf4c9a", "#555555", "#20a5a3")
        )
        width = max(canvas.winfo_width(), 420)
        height = max(canvas.winfo_height(), 120)
        percent_mode = self.trade_history_metric_var.get() == "Percent Change"
        legend_x, legend_y = 12, 10
        for index, item in enumerate(series):
            color = colors[index % len(colors)]
            text = f"{item['side']}: {item['name']}"
            item_width = min(360, len(text) * 6.2 + 26)
            if legend_x + item_width > width - 10:
                legend_x = 12
                legend_y += 15
            canvas.create_line(legend_x, legend_y, legend_x + 18, legend_y,
                               fill=color, width=2, dash=(4, 2) if item["side"] == "Receive" else ())
            canvas.create_text(legend_x + 23, legend_y, anchor="w", text=text,
                               fill=self.theme_colors["foreground"], font=("TkDefaultFont", 8))
            legend_x += item_width
        top = min(legend_y + 13, height - 46)
        left, right, bottom = 55, width - 12, height - 27
        movements = []
        for item in series:
            first_value = item["points"][0][1]
            if first_value:
                if percent_mode:
                    movements.extend((value / first_value - 1) * 100 for _date, value in item["points"])
                else:
                    movements.extend(value for _date, value in item["points"])
        if not movements:
            canvas.create_text(12, top, anchor="nw", text="No comparable value history is available.",
                               fill=self.theme_colors["muted"])
            return
        data_min, data_max = min(movements), max(movements)
        padding = max(abs(data_max) * 0.05, 1) if data_min == data_max else (data_max - data_min) * 0.12
        minimum, maximum = data_min - padding, data_max + padding
        spread = maximum - minimum or 1
        for fraction in (0, 0.5, 1):
            y = bottom - fraction * (bottom - top)
            label = minimum + fraction * spread
            canvas.create_line(left, y, right, y, fill=self.theme_colors["border"], dash=(2, 3))
            axis_label = f"{label:+.1f}%" if percent_mode else f"{label:,.0f}"
            canvas.create_text(2, y, anchor="w", text=axis_label,
                               fill=self.theme_colors["muted"], font=("TkDefaultFont", 8))
        hover_points = []
        for index, item in enumerate(series):
            points = item["points"]
            first_value = points[0][1]
            if not first_value:
                continue
            coords = []
            for point_index, (_date, value) in enumerate(points):
                x = left if len(points) == 1 else left + point_index * (right - left) / (len(points) - 1)
                change_pct = (value / first_value - 1) * 100
                graph_value = change_pct if percent_mode else value
                y = bottom - (graph_value - minimum) / spread * (bottom - top)
                coords.extend((x, y))
                hover_points.append({
                    "x": x, "y": y, "date": _date, "value": value,
                    "change_pct": change_pct, "name": item["name"], "side": item["side"],
                })
            if len(coords) >= 4:
                canvas.create_line(*coords, fill=colors[index % len(colors)], width=2, smooth=True,
                                   dash=(5, 3) if item["side"] == "Receive" else ())
        reference = max(series, key=lambda item: len(item["points"]))["points"]
        tick_count = min(len(reference), max(2, int((right - left) / 78) + 1))
        tick_indexes = sorted({round(index * (len(reference) - 1) / (tick_count - 1))
                               for index in range(tick_count)})
        for index in tick_indexes:
            x = left if len(reference) == 1 else left + index * (right - left) / (len(reference) - 1)
            canvas.create_line(x, top, x, bottom, fill=self.theme_colors["border"], dash=(1, 4))
            canvas.create_text(x, height - 2, anchor="s", text=reference[index][0][5:],
                               fill=self.theme_colors["muted"], font=("TkDefaultFont", 7))
        self.trade_history_plot_points[str(canvas)] = hover_points

    def _show_trade_history_hover(self, event, canvas):
        #Show the closest snapshot's date, raw value, and indexed movement.
        points = self.trade_history_plot_points.get(str(canvas), [])
        if not points:
            self._clear_trade_history_hover(canvas)
            return
        point = min(points, key=lambda item: (item["x"] - event.x) ** 2 + (item["y"] - event.y) ** 2)
        if (point["x"] - event.x) ** 2 + (point["y"] - event.y) ** 2 > 14 ** 2:
            self._clear_trade_history_hover(canvas)
            return
        self._clear_trade_history_hover(canvas)
        if self.trade_history_metric_var.get() == "Percent Change":
            detail = f"{point['change_pct']:+.1f}% from start · value {point['value']:,.0f}"
        else:
            detail = f"value {point['value']:,.0f} · {point['change_pct']:+.1f}% from start"
        text = f"{point['side']}: {point['name']}\n{point['date']} · {detail}"
        x = min(max(event.x + 14, 4), max(4, canvas.winfo_width() - 280))
        y = min(max(event.y + 14, 4), max(4, canvas.winfo_height() - 48))
        background = "#292f36" if self.dark_mode_var.get() else "#ffffff"
        foreground = "#edf1f5" if self.dark_mode_var.get() else "#20252b"
        border = self.theme_colors["border"]
        canvas.create_rectangle(x, y, x + 276, y + 44, fill=background, outline=border,
                                tags=("trade-history-hover",))
        canvas.create_text(x + 6, y + 5, anchor="nw", text=text, width=264,
                           fill=foreground, font=("TkDefaultFont", 8),
                           tags=("trade-history-hover",))

    @staticmethod
    def _clear_trade_history_hover(canvas):
        #Remove the transient detail callout when the pointer moves away from a data point.
        if canvas.winfo_exists():
            canvas.delete("trade-history-hover")

    def _render_value_history(self, player_id, value_format, history_payload,
                              canvas=None, detail_var=None, points_name="value_history_points"):
        #Summarize the recent series and compute volatility from daily percentage changes.
        history = []
        for point in history_payload.get("history", []):
            try:
                history.append((str(point["date"]), float(point["value"])))
            except (KeyError, TypeError, ValueError):
                continue
        history.sort(key=lambda point: point[0])
        canvas = canvas or self.value_history_canvas
        detail_var = detail_var or self.value_trend_detail
        setattr(self, points_name, history)
        player = self.stats_guy_players.get(player_id, {})
        if not history:
            detail_var.set(f"No Stats Guy history is available for {player.get('name', 'this player')}.")
            self._draw_history_chart(canvas, history, "No history is available for this player.")
            return
        daily_changes = []
        for previous, current in zip(history, history[1:]):
            if previous[1] > 0:
                daily_changes.append((current[1] - previous[1]) / previous[1] * 100)
        volatility = statistics.pstdev(daily_changes) if len(daily_changes) > 1 else None
        largest_swing = max((abs(change) for change in daily_changes), default=None)
        current_value = history[-1][1]
        week_change = current_value - history[-8][1] if len(history) >= 8 else None
        month_change = current_value - history[-31][1] if len(history) >= 31 else None
        low_value = min(value for _date, value in history)
        high_value = max(value for _date, value in history)
        volatility_text = f"{volatility:.2f}%" if volatility is not None else "not enough snapshots"
        swing_text = f"{largest_swing:.2f}%" if largest_swing is not None else "—"
        detail_var.set(
            f"{player.get('name', 'Player')} · {value_format.replace('_', ' ').title()} · "
            f"{len(history)} daily snapshots ({history[0][0]} to {history[-1][0]}). "
            f"7-day change: {self._display_value(week_change, signed=True)}; "
            f"30-day change: {self._display_value(month_change, signed=True)}. "
            f"Daily volatility: {volatility_text} standard deviation; "
            f"largest daily move: {swing_text}. Range: {low_value:,.0f}–{high_value:,.0f}. "
            "Historical variability, not a forecast."
        )
        self._draw_history_chart(canvas, history, "No history is available for this player.")

    def _draw_value_history(self):
        #Draw the Value Trends history through the shared history chart renderer.
        if not hasattr(self, "value_history_canvas"):
            return
        self._draw_history_chart(
            self.value_history_canvas, self.value_history_points,
            "Select a player to load a 90-day trend chart.",
        )

    def _draw_history_chart(self, canvas, points, empty_text):
        #Keep an open pop-out synchronized with redraws of its compact source graph.
        self._render_history_chart(canvas, points, empty_text)
        source_canvas = {
            "roster": getattr(self, "roster_history_canvas", None),
            "weekly": getattr(self, "weekly_history_canvas", None),
            "trends": getattr(self, "value_history_canvas", None),
        }.get(self.player_history_popup_source)
        popup = self.player_history_popup_canvas
        if popup and popup.winfo_exists() and canvas is source_canvas:
            self._render_history_chart(
                popup, getattr(self, self.player_history_popup_points_name, []),
                self.player_history_popup_empty_text,
            )

    def _render_history_chart(self, canvas, points, empty_text):
        #Render a compact shared chart with date ticks sized to the visible width.
        canvas.delete("all")
        canvas.configure(bg=self.theme_colors["surface"], highlightbackground=self.theme_colors["border"])
        if not points:
            canvas.create_text(
                12, 18, anchor="nw", text=empty_text,
                fill=self.theme_colors["muted"],
            )
            return
        width = max(canvas.winfo_width(), 400)
        height = max(canvas.winfo_height(), 100)
        left, right, top, bottom = 56, width - 14, 14, height - 32
        values = [value for _date, value in points]
        data_minimum, data_maximum = min(values), max(values)
        data_spread = data_maximum - data_minimum
        padding = max(abs(data_maximum) * 0.05, 1) if data_spread == 0 else data_spread * 0.12
        minimum = data_minimum - padding
        maximum = data_maximum + padding
        spread = maximum - minimum
        for fraction in (0, 0.5, 1):
            y = bottom - fraction * (bottom - top)
            label = minimum + fraction * spread
            canvas.create_line(left, y, right, y, fill=self.theme_colors["border"], dash=(2, 3))
            canvas.create_text(2, y, anchor="w", text=f"{label:,.0f}",
                               fill=self.theme_colors["muted"], font=("TkDefaultFont", 8))
        coords = []
        for index, (_date, value) in enumerate(points):
            x = left if len(points) == 1 else left + index * (right - left) / (len(points) - 1)
            y = bottom - (value - minimum) / spread * (bottom - top)
            coords.extend((x, y))
        if len(coords) >= 4:
            canvas.create_line(*coords, fill=self.theme_colors["accent"], width=2, smooth=True)
        elif coords:
            x, y = coords
            canvas.create_oval(x - 3, y - 3, x + 3, y + 3, fill=self.theme_colors["accent"], outline="")
        tick_count = min(len(points), max(2, int((right - left) / 72) + 1))
        tick_indexes = sorted({round(index * (len(points) - 1) / (tick_count - 1))
                               for index in range(tick_count)})
        for index in tick_indexes:
            x = left if len(points) == 1 else left + index * (right - left) / (len(points) - 1)
            canvas.create_line(x, top, x, bottom, fill=self.theme_colors["border"], dash=(1, 4))
            canvas.create_text(x, height - 2, anchor="s", text=points[index][0][5:],
                               fill=self.theme_colors["muted"], font=("TkDefaultFont", 7))

    def _load_weekly_context(self):
        #Fetch Sleeper's current NFL week and this league's matchup scores in the background.
        league = self.selected_league() or {}
        league_id = league.get("league_id")
        if not league_id:
            self.weekly_context_error = "Load a league to see current matchup context."
            self._refresh_weekly_help()
            return
        self.weekly_context_error = "Loading current NFL week and matchup data..."
        self.weekly_schedule_error = ""
        self.weekly_bye_weeks = {}
        self._refresh_weekly_help()

        def task():
            try:
                state = self.api.get_nfl_state()
                season = state.get("league_season") or state.get("season")
                league_season = league.get("season")
                if season and league_season and str(season) != str(league_season):
                    return state, [], {}, (
                        f"This league is for {league_season}; Sleeper's current NFL season is {season}. "
                        "No current-season matchup is available for this league."
                    ), ""
                week = state.get("display_week") or state.get("week") or state.get("leg")
                if not week:
                    matchup_rows = []
                    matchup_error = "Sleeper has not published a current NFL week yet."
                else:
                    try:
                        matchup_rows = self.api.get_matchups(league_id, int(week))
                        matchup_error = ""
                    except Exception as exc:
                        matchup_rows = []
                        matchup_error = f"Could not load weekly matchup data: {exc}"
                schedule_season = league_season or season
                try:
                    schedule = self.api.get_nfl_schedule(schedule_season)
                    bye_weeks = self._bye_weeks_from_schedule(schedule)
                    schedule_error = "" if bye_weeks else "The schedule did not include bye-week data for this season."
                except Exception as exc:
                    bye_weeks = {}
                    schedule_error = f"Could not load bye-week schedule: {exc}"
                return state, matchup_rows, bye_weeks, matchup_error, schedule_error
            except Exception as exc:
                return {}, [], {}, f"Could not load weekly matchup data: {exc}", ""

        def loaded(result):
            #Ignore a slow response if the user has already loaded a different league.
            if str((self.selected_league() or {}).get("league_id", "")) != str(league_id):
                return
            (self.weekly_state, self.weekly_matchups, self.weekly_bye_weeks,
             self.weekly_context_error, self.weekly_schedule_error) = result
            self._refresh_weekly_help()
            self._refresh_trade_bye_weeks()
            self._load_roster_projections()

        self.run_background(task, loaded)

    @staticmethod
    def _normalize_nfl_team_code(team_code):
        #Align schedule abbreviations with Sleeper's NFL team codes.
        team_code = str(team_code or "").strip().upper()
        aliases = {"JAC": "JAX", "LA": "LAR", "STL": "LAR", "SD": "LAC",
                   "OAK": "LV", "WSH": "WAS"}
        return aliases.get(team_code, team_code)

    @classmethod
    def _bye_weeks_from_schedule(cls, schedule):
        #Find each team's bye as the regular-season week where it has no scheduled game.
        teams_by_week = {}
        all_teams = set()
        for game in schedule:
            try:
                week = int(game.get("week", ""))
            except (TypeError, ValueError):
                continue
            if not 1 <= week <= 18:
                continue
            teams = {
                cls._normalize_nfl_team_code(game.get("home_team")),
                cls._normalize_nfl_team_code(game.get("away_team")),
            } - {""}
            teams_by_week.setdefault(week, set()).update(teams)
            all_teams.update(teams)
        bye_weeks = {}
        for week, playing_teams in teams_by_week.items():
            for team in all_teams - playing_teams:
                bye_weeks[team] = week
        return bye_weeks

    def _refresh_weekly_help(self, _event=None):
        #Summarize the selected team's live matchup and identify rostered players on bye.
        if _event is not None and hasattr(self, "weekly_history_detail"):
            self.weekly_history_player_id = None
            self.weekly_history_points = []
            self.weekly_history_detail.set("Select a waiver option to see its recent value movement.")
            self._draw_history_chart(self.weekly_history_canvas, [], "Select a player to load a 90-day value trend.")
        if self.dataframes is None:
            self.weekly_summary_var.set("Load a league to see this week's matchup and bye coverage.")
            self._refresh_weekly_waivers()
            return
        if self.weekly_context_error:
            self.weekly_summary_var.set(self.weekly_context_error)
            self.weekly_priority_positions = set()
            self._refresh_weekly_waivers()
            return
        roster_id = self._selected_roster_id(self.weekly_team_combo)
        if roster_id is None:
            self.weekly_summary_var.set("Choose a team to see its matchup and bye coverage.")
            self._refresh_weekly_waivers()
            return

        week = self.weekly_state.get("display_week") or self.weekly_state.get("week") or self.weekly_state.get("leg")
        team_frame = self.dataframes["Teams"]
        team_row = team_frame[team_frame["Roster ID"].astype(str) == str(roster_id)]
        team_name = str(team_row.iloc[0]["Team Name"]) if not team_row.empty else "Your team"
        roster = self.dataframes["Rosters"]
        team_roster = roster[roster["Roster ID"].astype(str) == str(roster_id)]

        matchup_rows = [row for row in self.weekly_matchups
                        if str(row.get("roster_id")) == str(roster_id)]
        matchup = matchup_rows[0] if matchup_rows else None
        opponent = None
        if matchup and matchup.get("matchup_id") is not None:
            opponent = next((row for row in self.weekly_matchups
                             if row.get("matchup_id") == matchup.get("matchup_id")
                             and str(row.get("roster_id")) != str(roster_id)), None)
        opponent_name = "opponent not found"
        score_text = "No matchup scores posted yet."
        if opponent is not None:
            opponent_row = team_frame[team_frame["Roster ID"].astype(str) == str(opponent.get("roster_id"))]
            if not opponent_row.empty:
                opponent_name = str(opponent_row.iloc[0]["Team Name"])
            own_score = float(matchup.get("points") or 0)
            opponent_score = float(opponent.get("points") or 0)
            if own_score or opponent_score:
                margin = own_score - opponent_score
                if abs(margin) < 0.05:
                    score_text = f"Score: {own_score:.1f}–{opponent_score:.1f} (tied)."
                elif margin > 0:
                    score_text = f"Score: {own_score:.1f}–{opponent_score:.1f} (leading by {margin:.1f})."
                else:
                    score_text = f"Score: {own_score:.1f}–{opponent_score:.1f} (trailing by {abs(margin):.1f})."

        bye_rows = []
        bye_data_available = bool(self.weekly_bye_weeks)
        if week and bye_data_available:
            for _, player in team_roster.iterrows():
                team_code = self._normalize_nfl_team_code(player.get("NFL Team", ""))
                player_bye = self.weekly_bye_weeks.get(team_code)
                if player_bye == int(week):
                    bye_rows.append(player)
        starter_byes = [player for player in bye_rows
                        if str(player.get("Slot", "")).casefold() not in {"bench", "reserve", "ir"}]
        starter_bye_ids = {str(player.get("Player ID", "")) for player in starter_byes}
        self.weekly_priority_positions = {
            position.strip().upper()
            for player in starter_byes
            for position in str(player.get("Position", "")).split(",")
            if position.strip()
        }
        if not bye_data_available:
            bye_text = self.weekly_schedule_error or "Bye-week information is not available for this season."
        elif not week:
            bye_text = "The current NFL week is unavailable, so bye coverage cannot be checked."
        else:
            bye_text = "No rostered players are on bye this week."
        if bye_rows:
            affected = [
                f"{player['Player']} ({player['Position']}{', starter' if str(player.get('Player ID', '')) in starter_bye_ids else ''})"
                for player in bye_rows
            ]
            focus = ", ".join(sorted(self.weekly_priority_positions)) or "No starting-slot byes"
            bye_text = "On bye: " + ", ".join(affected) + f". Coverage focus: {focus}."
        state_text = f"NFL Week {week}" if week else "Current NFL week unavailable"
        matchup_text = f"{team_name} vs {opponent_name}. {score_text}"
        if matchup is None and self.weekly_matchups:
            matchup_text = f"No matchup entry found for {team_name} in Week {week}."
        self.weekly_team_roster_id = str(roster_id)
        self.weekly_summary_var.set(
            f"{state_text}\n{matchup_text}\n{bye_text}\n"
            "Matchup scores update from Sleeper; FantasyCalc values are market estimates, not weekly projections."
        )
        if self.weekly_schedule_error and bye_data_available:
            self.weekly_summary_var.set(
                self.weekly_summary_var.get() + f"\nSchedule note: {self.weekly_schedule_error}"
            )
        self._refresh_weekly_waivers()

    def _refresh_weekly_waivers(self, _event=None):
        #List unrostered NFL players, favoring eligible positions affected by this week's byes.
        if not hasattr(self, "weekly_tree"):
            return
        for item in self.weekly_tree.get_children():
            self.weekly_tree.delete(item)
        if self.dataframes is None:
            self.weekly_waiver_count.set("Load a league to view free agents.")
            return
        agents = self.dataframes.get("Free Agents", pd.DataFrame())
        if agents.empty:
            self.weekly_waiver_count.set("No free-agent players were found in the Sleeper player directory.")
            return
        focus = self.weekly_filter_combo.get() or "Bye Coverage"
        rows = []
        for _, player in agents.iterrows():
            positions = {part.strip().upper() for part in str(player.get("Position", "")).split(",") if part.strip()}
            if focus in {"QB", "RB", "WR", "TE", "K", "DEF"} and focus not in positions:
                continue
            if focus == "Bye Coverage" and self.weekly_priority_positions and not positions.intersection(self.weekly_priority_positions):
                continue
            #Skip unsigned or inactive NFL entries; they are not useful waiver options.
            if str(player.get("Status", "")).casefold() not in {"active", ""}:
                continue
            player_id = str(player.get("Player ID", ""))
            value = self.fantasycalc_values.get(player_id)
            rows.append({
                "id": player_id,
                "name": str(player.get("Player", "")),
                "position": str(player.get("Position", "")),
                "nfl_team": str(player.get("NFL Team", "")),
                "status": self._player_list_status(player),
                "bye": str(self.weekly_bye_weeks.get(
                    self._normalize_nfl_team_code(player.get("NFL Team", "")), "—"
                )),
                "value": value,
            })
        self.weekly_waiver_rows = rows
        self.weekly_sort_reverse = {"value": True}
        self._sort_weekly_waivers("value", toggle=False)
        focus_text = "bye-affected positions" if focus == "Bye Coverage" else focus
        if focus == "Bye Coverage" and not self.weekly_priority_positions:
            focus_text = "all positions (no starting players marked on bye)"
        self.weekly_waiver_count.set(
            f"Showing {len(rows)} available player(s) for {focus_text}. "
            "Click a column heading to sort; N/A means FantasyCalc has no value for that player."
        )

    def _sort_weekly_waivers(self, column, toggle=True):
        #Sort free agents by the clicked column while keeping unavailable values at the bottom.
        if not hasattr(self, "weekly_tree"):
            return
        for item in self.weekly_tree.get_children():
            self.weekly_tree.delete(item)
        if toggle:
            reverse = not self.weekly_sort_reverse.get(column, column == "value")
        else:
            reverse = self.weekly_sort_reverse.get(column, column == "value")
        self.weekly_sort_reverse[column] = reverse
        rows = list(self.weekly_waiver_rows)
        if column == "value":
            known = [row for row in rows if row["value"] is not None]
            missing = [row for row in rows if row["value"] is None]
            known.sort(key=lambda row: row["value"], reverse=reverse)
            rows = known + missing
        elif column == "name":
            rows.sort(key=lambda row: (row["name"].rsplit(" ", 1)[-1].casefold(), row["name"].casefold()), reverse=reverse)
        elif column == "position":
            rows.sort(key=lambda row: self._position_sort_key(row["position"], row["name"].rsplit(" ", 1)[-1]), reverse=reverse)
        elif column == "bye":
            rows.sort(key=lambda row: (row["bye"] == "—", int(row["bye"]) if row["bye"].isdigit() else 99), reverse=reverse)
        else:
            rows.sort(key=lambda row: row[column].casefold(), reverse=reverse)
        for index, row in enumerate(rows):
            shown_value = f"{row['value']:,.0f}" if row["value"] is not None else "N/A"
            iid = row["id"] or f"weekly-{index}"
            self.weekly_tree.insert(
                "", "end", iid=iid, text=row["name"],
                values=(row["position"], row["nfl_team"], row["status"], row["bye"], shown_value),
            )

    def _load_fantasycalc(self):
        #Request FantasyCalc values using the league's size/scoring and the chosen league type.
        league = self.selected_league() or {}
        scoring = league.get("scoring_settings") or {}
        roster_positions = [str(p).upper() for p in league.get("roster_positions") or []]
        num_qbs = 2 if ("SF" in roster_positions or "SUPER_FLEX" in roster_positions or
                        roster_positions.count("QB") > 1) else 1
        teams = int(league.get("total_rosters") or len(self.team_options) or 12)
        rec = float(scoring.get("rec", 0) or 0)
        ppr = min((0, 0.5, 1), key=lambda value: abs(value - rec))
        params = {
            "isDynasty": str(self.format_var.get() == "Dynasty").lower(),
            "numQbs": num_qbs,
            "numTeams": teams,
            "ppr": ppr,
        }
        te_premium = any("te" in str(key).lower() and "rec" in str(key).lower() and value
                         for key, value in scoring.items())
        note = " · TE premium not reflected" if te_premium else ""
        note += " · closest supported PPR" if abs(rec - ppr) > 0.001 else ""
        self.market_settings_var.set(f"{teams} teams · {'Superflex' if num_qbs == 2 else '1QB'} · {ppr:g} PPR{note}")

        def task():
            try:
                rows = self.api.get_fantasycalc_values(params)
                values = {}
                trends = {}
                player_positions = {}
                for entry in rows:
                    player = entry.get("player") or {}
                    sleeper_id = player.get("sleeperId")
                    if sleeper_id not in (None, "") and entry.get("value") is not None:
                        player_id = str(sleeper_id)
                        values[player_id] = float(entry["value"])
                        raw_positions = player.get("fantasyPositions") or player.get("position") or []
                        if isinstance(raw_positions, str):
                            raw_positions = raw_positions.replace("/", ",").split(",")
                        elif not isinstance(raw_positions, (list, tuple, set)):
                            raw_positions = [raw_positions]
                        player_positions[player_id] = tuple(
                            str(position).strip().upper() for position in raw_positions if str(position).strip()
                        )
                        if entry.get("trend30Day") is not None:
                            trends[player_id] = float(entry["trend30Day"])
                overall_ranks = {}
                sorted_players = sorted(values.items(), key=lambda item: (-item[1], item[0]))
                previous_value = None
                for index, (player_id, value) in enumerate(sorted_players, start=1):
                    if value != previous_value:
                        current_rank = index
                    overall_ranks[player_id] = current_rank
                    previous_value = value
                by_position = {}
                for player_id, positions in player_positions.items():
                    for position in positions:
                        by_position.setdefault(position, []).append(player_id)
                position_ranks = {}
                for position, player_ids in by_position.items():
                    previous_value = None
                    sorted_ids = sorted(player_ids, key=lambda item: (-values[item], item))
                    for index, player_id in enumerate(sorted_ids, start=1):
                        value = values[player_id]
                        if value != previous_value:
                            current_rank = index
                        position_ranks.setdefault(player_id, {})[position] = current_rank
                        previous_value = value
                return values, trends, overall_ranks, position_ranks, None
            except Exception as exc:
                return {}, {}, {}, {}, str(exc)

        def loaded(result):
            values, trends, overall_ranks, position_ranks, error = result
            self.fantasycalc_values = values
            self.fantasycalc_trends = trends
            self.fantasycalc_overall_ranks = overall_ranks
            self.fantasycalc_position_ranks = position_ranks
            self._refresh_trade_builder_values()
            self._populate_trade_rosters()
            self._refresh_saved_trades_view()
            self._refresh_trade_ideas()
            self._refresh_weekly_waivers()
            self._refresh_value_trends()
            if error:
                self.status(f"League loaded. FantasyCalc values could not be loaded: {error}")
            else:
                self.status(f"FantasyCalc market values loaded for {len(values)} players.")
        self.run_background(task, loaded)

    def run_background(self, function, on_success=None):
        #Run network work off the UI thread and return the result to Tk's event loop.
        self._background_task_count += 1
        if self._background_task_count == 1:
            self.progress.configure(mode="indeterminate", value=0)
            self.progress.pack(fill="x")
            self.progress.start(10)

        def worker():
            try:
                result = function()
                self.after(0, lambda: self._background_success(result, on_success))
            except Exception as exc:
                self.after(0, lambda: self._background_error(exc))

        threading.Thread(target=worker, daemon=True).start()

    def _background_success(self, result, callback):
        #Stop the activity indicator before passing successful data to its UI handler.
        self._finish_background_task()
        if callback:
            callback(result)

    def _background_error(self, exc):
        #Stop the activity indicator and show a readable error if background work fails.
        self._finish_background_task()
        self.status("Operation failed.")
        messagebox.showerror("Fantasy Trade Calculator", str(exc))

    def _finish_background_task(self):
        #Keep the activity indicator visible until every overlapping task has finished.
        self._background_task_count = max(0, self._background_task_count - 1)
        if self._background_task_count == 0:
            self.progress.stop()
            self.progress.configure(mode="determinate", value=0)
            self.progress.pack_forget()

    def load_leagues(self):
        #Look up the username and fetch that account's leagues for the chosen season.
        username = self.username_var.get().strip()
        season = self.season_var.get().strip()

        if not username:
            messagebox.showwarning("Missing Username", "Enter your Sleeper username.")
            return

        def task():
            self.status(f"Looking up Sleeper user '{username}'...")
            user = self.api.get_user(username)

            self.status(f"Loading {season} leagues...")
            leagues = self.api.get_leagues(user["user_id"], season)
            return user, leagues

        self.run_background(task, self._leagues_loaded)

    def _leagues_loaded(self, result):
        #Populate the league picker and reopen the last-used league when possible.
        self.user, self.leagues = result

        if not self.leagues:
            self.league_combo["values"] = []
            self.league_var.set("")
            self._update_window_context()
            self.status("No leagues were found for that user and season.")
            messagebox.showinfo(
                "No Leagues",
                "No NFL leagues were found for that username and season."
            )
            return

        values = [
            f'{league.get("name", "Unnamed League")}'
            for league in self.leagues
        ]

        self.league_combo["values"] = values
        self.league_combo.current(0)
        self._update_window_context()

        if self.startup_autoload_pending:
            self.startup_autoload_pending = False
            saved_league_id = str(self.saved_settings.get("league_id", ""))
            matching_index = next(
                (index for index, league in enumerate(self.leagues)
                 if str(league.get("league_id", "")) == saved_league_id),
                None,
            )
            if matching_index is not None:
                self.league_combo.current(matching_index)
                self.after_idle(self.load_selected_league)
            else:
                self.status(
                    "Your saved league was not found for this season. "
                    "Choose a league and click Load Selected League."
                )
                return

        self.status(
            f"Found {len(self.leagues)} league(s) for "
            f"{self.user.get('display_name') or self.user.get('username')}."
        )

    def selected_league(self):
        #Return the league object corresponding to the visible picker selection.
        index = self.league_combo.current()
        if index < 0 or index >= len(self.leagues):
            return None
        return self.leagues[index]

    def _save_selected_league(self, _event=None):
        #Remember the selected league ID so the app can reopen it on the next launch.
        league = self.selected_league()
        if not league:
            self._update_window_context()
            return
        self.saved_settings["league_id"] = str(league.get("league_id", ""))
        self._save_settings()
        self._update_window_context()

    def load_selected_league(self):
        #Save the current choices and fetch league, roster, and player data in the background.
        league = self.selected_league()

        if not league:
            messagebox.showwarning("No League", "Load your leagues and select one first.")
            return

        league_id = league["league_id"]
        self.saved_settings["league_id"] = str(league_id)
        self.saved_settings["username"] = self.username_var.get().strip()
        self.saved_settings["season"] = self.season_var.get().strip()
        self._save_settings()
        self._update_window_context()

        def task():
            return build_data(
                self.api,
                league_id,
                self.status
            )

        self.run_background(task, self._league_loaded)

    def _league_loaded(self, dataframes):
        #Enable analysis tools, choose the logged-in user's team, and refresh market data.
        self.dataframes = dataframes
        self.loaded_username = self.username_var.get().strip()
        self._update_window_context()
        self.fantasycalc_values = {}
        self.fantasycalc_trends = {}
        self.fantasycalc_overall_ranks = {}
        self.fantasycalc_position_ranks = {}
        self._new_trade()
        self.export_button.configure(state="normal")

        self.team_options = []
        for _, team in dataframes["Teams"].iterrows():
            manager = str(team.get("Manager", "")).strip()
            team_name = str(team.get("Team Name", "Team")).strip()
            display = f"{team_name} ({manager})" if manager and manager.casefold() != team_name.casefold() else team_name
            self.team_options.append((display, team.get("Roster ID")))
        options = [display for display, _roster_id in self.team_options]
        self.needs_team_combo["values"] = options
        self.suggestion_team_combo["values"] = options
        self.weekly_team_combo["values"] = options
        self.roster_team_combo["values"] = options
        self.value_trend_team_combo["values"] = ["All League Players"] + options
        self.trade_builder_team_combo["values"] = options
        self.trade_left_team["values"] = options
        self.trade_right_team["values"] = options
        own_roster_id = None
        if self.user:
            own = dataframes["Teams"][dataframes["Teams"]["Owner ID"].astype(str) == str(self.user.get("user_id", ""))]
            if not own.empty:
                own_roster_id = str(own.iloc[0]["Roster ID"])
        default_index = next((i for i, (_name, roster_id) in enumerate(self.team_options)
                              if own_roster_id is not None and str(roster_id) == own_roster_id), 0)
        if options:
            self.needs_team_combo.current(default_index)
            self.suggestion_team_combo.current(default_index)
            self.weekly_team_combo.current(default_index)
            self.roster_team_combo.current(default_index)
            self.value_trend_team_combo.current(0)
            self.trade_builder_team_combo.current(default_index)
            self.trade_left_team.current(default_index)
            other_index = 1 if len(options) > 1 and default_index == 0 else 0
            self.trade_right_team.current(other_index)
        self._show_position_needs()
        self.weekly_state = {}
        self.weekly_matchups = []
        self.weekly_bye_weeks = {}
        self.weekly_context_error = ""
        self.weekly_schedule_error = ""
        self.roster_projection_values = {}
        self.roster_projection_metric = ""
        self.roster_projection_note.set("Loading this week's Sleeper projections…")
        self.roster_history_player_id = None
        self.roster_history_points = []
        self.roster_history_detail.set("Select a player to see recent value movement and volatility.")
        self.weekly_history_player_id = None
        self.weekly_history_points = []
        self.weekly_history_detail.set("Select a waiver option to see its recent value movement.")
        self._refresh_weekly_help()
        self._refresh_roster_view()
        self._refresh_value_trends()
        self._refresh_trade_ideas()
        self._trade_builder_team_changed()
        self._populate_trade_rosters()
        self._refresh_saved_trades_view()
        self.workspace.select(self.position_needs_tab)

        league_df = dataframes["League Info"]
        league_name = league_df.iloc[0]["League Name"]
        teams = len(dataframes["Teams"])
        players = len(dataframes["Rosters"])

        self.status(
            f"League loaded: {league_name} | "
            f"{teams} teams | {players} roster entries."
        )
        self._load_fantasycalc()
        self._load_stats_guy_values()
        self._load_weekly_context()

    def refresh_players(self):
        #Force a fresh Sleeper player map download, bypassing its 24-hour cache.
        def task():
            self.status("Downloading the latest Sleeper NFL player database...")
            self.api.get_players(force_refresh=True)
            return True

        self.run_background(task, lambda _: self.status(
            "Player database refreshed. Load the league again to use the new data."
        ))

    def export(self):
        #Ask where to save a team-by-team workbook and optionally open it after creation.
        if self.dataframes is None:
            messagebox.showwarning("Nothing to Export", "Load a league first.")
            return

        league_name = self.dataframes["League Info"].iloc[0]["League Name"]
        safe_name = "".join(
            c if c.isalnum() or c in " _-" else "_"
            for c in str(league_name)
        ).strip() or "Sleeper_League"

        filename = filedialog.asksaveasfilename(
            title="Save Excel Workbook",
            defaultextension=".xlsx",
            initialfile=f"{safe_name}.xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )

        if not filename:
            return

        try:
            self.status("Creating Excel workbook...")
            export_excel(self.dataframes, filename)
            self.status(f"Export complete: {filename}")
            if messagebox.askyesno(
                "Export Complete",
                f"Excel workbook created successfully:\n\n{filename}\n\nWould you like to open the file?"
            ):
                os.startfile(filename)
        except Exception as exc:
            messagebox.showerror("Export Failed", str(exc))


#Only create the desktop window when this file is launched directly.
if __name__ == "__main__":
    App().mainloop()
