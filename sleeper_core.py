# Contains the desktop app, its Sleeper and FantasyCalc API helpers, and the Excel export logic.
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
APP_VERSION = "1.1.5"
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
