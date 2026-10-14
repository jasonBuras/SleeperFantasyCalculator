# App behavior mixins. Shared libraries and application constants come from sleeper_core.
from sleeper_core import *

class InsightsFeatures:
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

