# Main desktop application entry point.
from sleeper_core import *
from app_trade import TradeFeatures
from app_insights import InsightsFeatures


class App(TradeFeatures, InsightsFeatures, tk.Tk):
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
        self.value_history_window_var = tk.StringVar(value="1 month")
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
        self.update_release_url_var = tk.StringVar(value="")
        self.update_release_link_text_var = tk.StringVar(value="")
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
            self.after(1000, self._check_for_app_updates)

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
        if hasattr(self, "update_notes_text"):
            self.update_notes_text.configure(
                background=surface, foreground=fg, insertbackground=fg,
            )
        if hasattr(self, "update_release_link"):
            self.update_release_link.configure(foreground=accent)

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
        self.update_release_url_var.set("")
        self.update_release_link_text_var.set("")
        self._set_update_notes("")
        self.run_background(self._fetch_latest_app_release_safely, self._show_app_update_result)

    def _set_update_notes(self, notes):
        #Show release notes as selectable plain text so users can copy them.
        if not hasattr(self, "update_notes_text"):
            self.update_notes_var.set(str(notes or ""))
            return
        self.update_notes_text.configure(state="normal")
        self.update_notes_text.delete("1.0", "end")
        self.update_notes_text.insert("1.0", str(notes or ""))
        self.update_notes_text.configure(state="disabled")

    def _copy_update_notes(self):
        #Copy the complete plain-text release notes to the clipboard.
        notes = self.update_notes_text.get("1.0", "end-1c")
        if notes.strip():
            self.clipboard_clear()
            self.clipboard_append(notes)

    def _open_update_release_page(self, _event=None):
        #Only open release links hosted in this project's GitHub Releases page.
        url = self.update_release_url_var.get().strip()
        prefix = "https://github.com/jasonBuras/SleeperFantasyCalculator/releases/"
        if url.startswith(prefix):
            __import__("webbrowser").open(url)

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
            self.update_release_url_var.set("")
            self.update_release_link_text_var.set("")
            self._set_update_notes(result["error"])
            return
        if result.get("not_published"):
            self.update_status_var.set(f"Installed version {APP_VERSION} · No release has been published yet.")
            self.update_release_url_var.set("")
            self.update_release_link_text_var.set("")
            self._set_update_notes("When a new version is released, its changes will appear here.")
            return
        release_url = result.get("release_url", "")
        self.update_release_url_var.set(release_url)
        self.update_release_link_text_var.set(
            "Open release notes on GitHub"
            if release_url.startswith("https://github.com/jasonBuras/SleeperFantasyCalculator/releases/")
            else ""
        )
        self._set_update_notes(result["notes"])
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
        self.update_release_link = ttk.Label(
            updates, textvariable=self.update_release_link_text_var, cursor="hand2",
            foreground=self.theme_colors["accent"], underline=True,
        )
        self.update_release_link.pack(anchor="w", pady=(7, 3))
        self.update_release_link.bind("<Button-1>", self._open_update_release_page)
        ttk.Label(updates, text="Release notes (plain text):").pack(anchor="w")
        notes_frame = ttk.Frame(updates)
        notes_frame.pack(fill="x", pady=(3, 0))
        self.update_notes_text = tk.Text(
            notes_frame, height=6, wrap="word", state="disabled",
            background=self.theme_colors["surface"], foreground=self.theme_colors["foreground"],
            insertbackground=self.theme_colors["foreground"], relief="solid", borderwidth=1,
        )
        notes_scroll = ttk.Scrollbar(notes_frame, orient="vertical", command=self.update_notes_text.yview)
        self.update_notes_text.configure(yscrollcommand=notes_scroll.set)
        self.update_notes_text.pack(side="left", fill="both", expand=True)
        notes_scroll.pack(side="right", fill="y")
        ttk.Button(notes_frame, text="Copy Notes", command=self._copy_update_notes).pack(side="right", padx=(6, 0))
        self._set_update_notes(self.update_notes_var.get())

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
        self._add_tree_expand_controls(roster_controls, self.roster_tree)
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

        self.trade_review_split = ttk.PanedWindow(trade_tab, orient="vertical")
        self.trade_review_split.pack(fill="x", pady=(0, 2))
        sides = ttk.PanedWindow(self.trade_review_split, orient="horizontal")
        left = ttk.LabelFrame(sides, text="You give", padding=8)
        right = ttk.LabelFrame(sides, text="You receive", padding=8)
        sides.add(left, weight=1)
        sides.add(right, weight=1)
        self.trade_review_split.add(sides, weight=3)
        trade_details = ttk.Frame(self.trade_review_split)
        self.trade_review_details = trade_details
        self.trade_review_split.add(trade_details, weight=2)
        self.trade_review_split.bind("<ButtonRelease-1>", self._resize_trade_review_split, add="+")
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
        copy_preview = ttk.LabelFrame(trade_details, text="Trade Overview", padding=5)
        copy_preview.pack(fill="x", pady=(5, 0))
        ttk.Label(copy_preview, textvariable=self.copy_trade_preview_var,
                  wraplength=1300, justify="left").pack(anchor="w")
        trade_actions = ttk.Frame(trade_details)
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
        note_row = ttk.Frame(trade_details)
        note_row.pack(fill="x", pady=(4, 0))
        ttk.Label(note_row, text="Note:").pack(side="left")
        self.trade_note_var = tk.StringVar()
        ttk.Entry(note_row, textvariable=self.trade_note_var).pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.trade_result_var = tk.StringVar(value="Load a league to review a trade.")
        ttk.Label(trade_details, textvariable=self.trade_result_var, font=("TkDefaultFont", 12, "bold"), wraplength=850).pack(anchor="w", pady=(6, 0))
        balance_row = ttk.Frame(trade_details)
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
        ttk.Label(trade_details, textvariable=self.trade_stats_guy_result_var, font=("TkDefaultFont", 10, "bold"),
                  wraplength=850).pack(anchor="w", pady=(3, 0))
        self.trade_impact_var = tk.StringVar(value="Roster and lineup impact details will appear here.")
        impact_row = ttk.Frame(trade_details)
        impact_row.pack(fill="x", pady=(5, 0))
        self.trade_impact_summary = tk.StringVar(value="Roster impact appears when you select players.")
        ttk.Label(impact_row, textvariable=self.trade_impact_summary, wraplength=760).pack(side="left", anchor="w", fill="x", expand=True)
        ttk.Button(impact_row, text="View Lineup Impact", command=self._show_lineup_impact).pack(side="right", padx=(8, 0))
        trade_history_panel = ttk.LabelFrame(trade_details, text="Selected Players' Value Trends", padding=4)
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
        self.after_idle(self._initialize_trade_review_split)

        #Saved Trades keeps proposals available for reopening and editing later.
        self.saved_trade_tab = ttk.Frame(self.workspace, padding=12)
        self.workspace.add(self.saved_trade_tab, text="Saved Trades")
        ttk.Label(self.saved_trade_tab, text="Saved trades refresh their totals whenever current FantasyCalc values load.", wraplength=850).pack(anchor="w", pady=(0, 4))
        saved_credit = ttk.Label(self.saved_trade_tab, text="Trade values provided by FantasyCalc · fantasycalc.com", foreground="#1f5a92", cursor="hand2")
        self.attribution_labels.append(saved_credit)
        saved_credit.pack(anchor="w", pady=(0, 8))
        saved_credit.bind("<Button-1>", lambda _event: __import__("webbrowser").open("https://fantasycalc.com/"))
        saved_tree_controls = ttk.Frame(self.saved_trade_tab)
        saved_tree_controls.pack(fill="x", pady=(0, 4))
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
        self._add_tree_expand_controls(saved_tree_controls, self.saved_trade_tree)
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
        self._add_tree_expand_controls(ideas_controls, self.trade_ideas_tree)
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
        builder_content = ttk.PanedWindow(builder_tab, orient="horizontal")
        builder_content.pack(fill="both", expand=True)
        give_box = ttk.LabelFrame(builder_content, text="Players You Would Send", padding=6)
        return_box = ttk.LabelFrame(builder_content, text="Near-Value Returns", padding=6)
        builder_content.add(give_box, weight=1)
        builder_content.add(return_box, weight=2)
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
        builder_tree_controls = ttk.Frame(return_box)
        builder_tree_controls.pack(fill="x", pady=(0, 4))
        self._add_tree_expand_controls(builder_tree_controls, self.trade_builder_results_tree)
        results_scroll = ttk.Scrollbar(return_box, orient="vertical", command=self.trade_builder_results_tree.yview)
        self.trade_builder_results_tree.configure(yscrollcommand=results_scroll.set)
        self.trade_builder_results_tree.pack(side="left", fill="both", expand=True)
        results_scroll.pack(side="right", fill="y")
        self.trade_builder_results_tree.bind("<Double-1>", lambda _event: self._open_trade_builder_match())
        self.trade_builder_summary = tk.StringVar(value="Select players to see near-value returns from other teams.")
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
        self.workspace.select(self.main_menu_tab)

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

if __name__ == "__main__":
    App().mainloop()
