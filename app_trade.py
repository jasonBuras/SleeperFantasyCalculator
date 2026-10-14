# App behavior mixins. Shared libraries and application constants come from sleeper_core.
from sleeper_core import *

class TradeFeatures:
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
        #Use the page scrollbar to reach lower Trade Review sections on short displays.
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

    def _initialize_trade_review_split(self):
        #Start with room for a useful number of roster rows and keep the details fully reachable.
        self.update_idletasks()
        panes = self.trade_review_split.panes()
        roster_pane = self.nametowidget(panes[0]) if panes else None
        roster_height = max(320, roster_pane.winfo_reqheight() if roster_pane else 0)
        self._trade_review_details_height = max(320, self.trade_review_details.winfo_reqheight())
        self.trade_review_split.configure(height=roster_height + self._trade_review_details_height)
        self.trade_review_split.sashpos(0, roster_height)
        self.after_idle(self._update_trade_review_scrollregion)

    def _resize_trade_review_split(self, _event=None):
        #Grow the scrollable page with the roster pane so the summary and graph never get clipped.
        details_height = getattr(self, "_trade_review_details_height", None)
        if details_height is None:
            return
        roster_height = self.trade_review_split.sashpos(0)
        if roster_height >= 0:
            self.trade_review_split.configure(height=roster_height + details_height)
            self.after_idle(self._update_trade_review_scrollregion)

    def _update_trade_review_scrollregion(self):
        #Refresh the single page scroll range after its resizable pane changes height.
        self.trade_tab_canvas.configure(scrollregion=self.trade_tab_canvas.bbox("all"))

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

    def _add_tree_expand_controls(self, parent, tree):
        #Add consistent controls for opening or closing every expandable row in a tree.
        controls = ttk.Frame(parent)
        controls.pack(side="right")
        ttk.Button(
            controls, text="Expand All",
            command=lambda: self._set_tree_expanded(tree, True),
        ).pack(side="left", padx=(4, 0))
        ttk.Button(
            controls, text="Collapse All",
            command=lambda: self._set_tree_expanded(tree, False),
        ).pack(side="left", padx=(4, 0))

    @staticmethod
    def _set_tree_expanded(tree, expanded):
        #Apply the requested state recursively so nested counteroffers are included.
        def set_children(parent):
            for item in tree.get_children(parent):
                tree.item(item, open=expanded)
                set_children(item)

        set_children("")

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

    def _clear_trade_builder_matches(self, message="Select players to see near-value returns from other teams."):
        #Remove suggestions as soon as their offer inputs change.
        pending_search = getattr(self, "_trade_builder_search_after", None)
        if pending_search is not None:
            self.after_cancel(pending_search)
            self._trade_builder_search_after = None
        for item in self.trade_builder_results_tree.get_children():
            self.trade_builder_results_tree.delete(item)
        self.trade_builder_matches.clear()
        self.trade_builder_all_matches = []
        self.trade_builder_summary.set(message)

    def _trade_builder_options_changed(self, _event=None):
        #Refresh suggestions when the user changes the team or value tolerance.
        self._update_trade_builder_selection()

    def _trade_builder_team_changed(self, _event=None):
        #Reload the offer roster and refresh suggestions for the selected team.
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
        #Summarize the offer and refresh suggestions for the latest selection.
        if hasattr(self, "trade_builder_results_tree"):
            self._clear_trade_builder_matches("Updating near-value returns…")
        selected = self.trade_builder_give_tree.selection()
        if not selected:
            self.trade_builder_selection_var.set("Select one or more players to send.")
            self.trade_builder_summary.set("Select players to see near-value returns from other teams.")
            return
        missing = [iid for iid in selected if str(iid) not in self.fantasycalc_values]
        if missing:
            self.trade_builder_selection_var.set(
                f"{len(selected)} selected · value unavailable for {len(missing)} player(s)."
            )
            self.trade_builder_summary.set("Waiting for FantasyCalc values before showing return packages.")
            return
        total = sum(self.fantasycalc_values[str(iid)] for iid in selected)
        self.trade_builder_selection_var.set(f"{len(selected)} selected · total FC value {total:,.0f}.")
        #Debounce multi-select input so one search runs after the user finishes selecting players.
        self._trade_builder_search_after = self.after(250, self._run_scheduled_trade_builder_search)

    def _run_scheduled_trade_builder_search(self):
        self._trade_builder_search_after = None
        self._build_trade_suggestions()

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

