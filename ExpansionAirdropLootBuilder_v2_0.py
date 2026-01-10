import json
import os
import shutil
import time
import xml.etree.ElementTree as ET
import webbrowser
from dataclasses import dataclass

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

# Optional logo support (safe if PIL not installed or logo missing)
try:
    from PIL import Image
except Exception:
    Image = None

"""
KPTools - Expansion Airdrop Loot Builder
v2.0
- Airdrops tool: unchanged from v1.9 (your working base)
- Added Market tab:
  - Load / New / Save / Save As for Expansion Market category JSON
  - Add items from loaded types.xml
  - Edit/remove market entries
  - Remembers last market folder in %APPDATA%
"""

APP_VERSION = "v2.0"

# KPTools palette (dark + green)
KP_BG = "#0b0f0c"
KP_PANEL = "#0f1712"
KP_PANEL_2 = "#0d1410"
KP_TEXT = "#d7ffe0"
KP_MUTED = "#8fd3a0"
KP_GREEN = "#00ff66"
KP_GREEN_DARK = "#00cc55"

# Treeview zebra / selection colors
TV_ROW_EVEN = "#101a14"
TV_ROW_ODD = "#0d1510"
TV_HEAD_BG = "#0d1410"
TV_BG = "#0f1712"
TV_SEL = "#1a2a20"

DISCORD_URL = "https://discord.gg/F9mTFPubhg"


# ---------------- Types XML helpers ----------------

@dataclass
class TypeItem:
    name: str
    source: str
    category: str = ""


def find_xml_files_in_folder(folder: str):
    out = []
    for root, _dirs, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(".xml"):
                out.append(os.path.join(root, f))
    return out


def load_types_xml(path: str):
    """
    Returns (items_list, meta_dict) where:
      items_list: [classname,...]
      meta_dict[classname] = {"source": <types file basename>, "category": ...}
    """
    tree = ET.parse(path)
    root = tree.getroot()

    items = []
    meta = {}
    src = os.path.basename(path)

    for t in root.findall("type"):
        name = t.get("name")
        if not name:
            continue

        cat = ""
        cat_node = t.find("category")
        if cat_node is not None and cat_node.text:
            cat = cat_node.text.strip()

        items.append(name)
        meta[name] = {"source": src, "category": cat}

    return items, meta


def safe_int(val, default=0):
    try:
        return int(val)
    except Exception:
        return default


# ---------------- Main App ----------------

class ExpansionAirdropLootBuilder(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")

        # set icon if it exists next to the script/exe
        try:
            icon_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "kp_icon.ico")
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass

        self.title(f"KPTools – Expansion Airdrop Loot Builder {APP_VERSION}")
        self.geometry("1320x860")
        self.minsize(1120, 720)
        self.configure(fg_color=KP_BG)

        # ---------- Config in %APPDATA% ----------
        self.config_dir = os.path.join(
            os.environ.get("APPDATA", os.path.expanduser("~")),
            "KPTools",
            "ExpansionAirdropLootBuilder"
        )
        self.config_path = os.path.join(self.config_dir, "config.json")

        self.cfg = {
            "last_airdrop_path": "",
            "types_folders": [],
            "last_market_path": "",
            "last_market_dir": ""
        }
        self._load_config()

        # State
        self.airdrop_path = ""
        self.airdrop_data = None

        # Market State
        self.market_path = ""
        self.market_data = None
        self.market_dirty = False

        self.types_folders = []
        self.types_files_loaded = []
        self.all_items = []
        self.item_meta = {}
        self.source_to_items = {}

        self.container_names = []
        self.container_index = {}

        self.current_container_key = None

        self.types_file_label = "None"
        self.airdrop_file_label = "None"

        self.selected_loot_index = None

        # Throttle jobs
        self._item_job = None
        self._loot_job = None

        # Search state
        self.var_search = ctk.StringVar(value="")
        self.var_source = ctk.StringVar(value="All Types Files")

        # dropdown for containers
        self.var_container = ctk.StringVar(value="Load airdrop file first")

        # Container-level settings (pending)
        self.var_item_count = ctk.StringVar(value="")
        self.var_infected_count = ctk.StringVar(value="")
        self.var_infected_enabled = ctk.BooleanVar(value=True)

        # Loot editor vars
        self.var_name = ctk.StringVar(value="")
        self.var_chance = ctk.StringVar(value="")
        self.var_min = ctk.StringVar(value="")
        self.var_max = ctk.StringVar(value="")
        self.var_qty = ctk.StringVar(value="")

        self._build_ui()

        self._apply_tree_style()
        self._update_paths_label()

        # Restore last session (if paths still exist)
        self._restore_last_session()

    # ---------------- Config ----------------

    def _load_config(self):
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    self.cfg.update(data)
        except Exception:
            pass

    def _save_config(self):
        try:
            os.makedirs(self.config_dir, exist_ok=True)
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(self.cfg, f, indent=2)
        except Exception:
            pass

    def _remember_airdrop_path(self, path: str):
        self.cfg["last_airdrop_path"] = path
        self._save_config()

    def _remember_types_folders(self, folders: list):
        # de-dupe, keep order
        out = []
        seen = set()
        for p in folders:
            if p and p not in seen:
                out.append(p)
                seen.add(p)
        self.cfg["types_folders"] = out
        self._save_config()

    def _restore_last_session(self):
        # Types folders
        folders = [p for p in (self.cfg.get("types_folders") or []) if isinstance(p, str) and os.path.isdir(p)]
        if folders:
            self.types_folders = []
            for folder in folders:
                self._load_types_folder_silent(folder)

        # Airdrop file
        last_air = self.cfg.get("last_airdrop_path") or ""
        if isinstance(last_air, str) and last_air and os.path.isfile(last_air):
            # don't auto-open dialogs; just load silently
            self._load_airdrop_path(last_air, silent=True)

        # Market file
        last_mkt = self.cfg.get("last_market_path") or ""
        if isinstance(last_mkt, str) and last_mkt and os.path.isfile(last_mkt):
            self._load_market_path(last_mkt, silent=True)

    # ---------------- UI ----------------

    def _build_ui(self):
        # Top bar
        top = ctk.CTkFrame(self, fg_color=KP_PANEL, corner_radius=12)
        top.pack(fill="x", padx=12, pady=(12, 8))

        ctk.CTkButton(
            top, text="Load AirdropSettings.json",
            fg_color=KP_GREEN_DARK, hover_color=KP_GREEN,
            text_color="black", command=self.load_airdrop
        ).pack(side="left", padx=10, pady=10)

        ctk.CTkButton(
            top, text="Load Types Folder(s)",
            fg_color=KP_PANEL_2, hover_color="#16251a",
            text_color=KP_TEXT, command=self.load_types
        ).pack(side="left", padx=0, pady=10)

        ctk.CTkButton(
            top, text="Save (Backup + Write)",
            fg_color=KP_GREEN_DARK, hover_color=KP_GREEN,
            text_color="black", command=self.save_airdrop
        ).pack(side="left", padx=10, pady=10)

        self.lbl_paths = ctk.CTkLabel(top, text="No files loaded yet.", text_color=KP_MUTED)
        self.lbl_paths.pack(side="left", padx=14)

        # Tabs
        tabs = ctk.CTkTabview(self, fg_color=KP_BG)
        tabs.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        tabs.add("Airdrops")
        tabs.add("Market")
        tabs.add("Info")

        # --- Airdrops tab root ---
        body = ctk.CTkFrame(tabs.tab("Airdrops"), fg_color=KP_BG)
        body.pack(fill="both", expand=True)

        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        # Left panel (types browser + loot list)
        left = ctk.CTkFrame(body, fg_color=KP_PANEL, corner_radius=12)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.grid_rowconfigure(4, weight=1)
        left.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(left, text="Item Browser", text_color=KP_TEXT, font=("Segoe UI", 16, "bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(12, 6)
        )

        self.dd_source = ctk.CTkOptionMenu(
            left,
            values=["All Types Files"],
            variable=self.var_source,
            fg_color=KP_PANEL_2,
            button_color=KP_GREEN_DARK,
            button_hover_color=KP_GREEN,
            text_color=KP_TEXT,
            dropdown_fg_color=KP_PANEL,
            dropdown_text_color=KP_TEXT,
            command=lambda _: self.refresh_item_filter()
        )
        self.dd_source.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))

        search_row = ctk.CTkFrame(left, fg_color="transparent")
        search_row.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 8))
        search_row.grid_columnconfigure(0, weight=1)

        ent = ctk.CTkEntry(
            search_row, textvariable=self.var_search,
            fg_color=KP_PANEL_2, text_color=KP_TEXT,
            placeholder_text="Search classnames…",
            placeholder_text_color=KP_MUTED
        )
        ent.grid(row=0, column=0, sticky="ew")
        ent.bind("<KeyRelease>", lambda _e: self.refresh_item_filter())

        ctk.CTkButton(
            search_row, text="Clear",
            fg_color=KP_PANEL_2, hover_color="#16251a",
            text_color=KP_TEXT,
            command=self.clear_search
        ).grid(row=0, column=1, padx=(8, 0))

        ctk.CTkButton(
            left, text="Add Selected → Loot",
            fg_color=KP_GREEN_DARK, hover_color=KP_GREEN,
            text_color="black", command=self.add_selected_item_to_loot
        ).grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 8))

        # Item list (FAST Treeview)
        item_host = ctk.CTkFrame(left, fg_color=KP_PANEL_2, corner_radius=10)
        item_host.grid(row=4, column=0, sticky="nsew", padx=12, pady=(0, 8))
        item_host.grid_rowconfigure(0, weight=1)
        item_host.grid_columnconfigure(0, weight=1)

        self.item_tree = ttk.Treeview(
            item_host,
            columns=("Name", "Source"),
            show="headings",
            selectmode="browse"
        )
        self.item_tree.heading("Name", text="Classname")
        self.item_tree.heading("Source", text="Types File")
        self.item_tree.column("Name", width=280, anchor="w")
        self.item_tree.column("Source", width=160, anchor="w")
        self.item_tree.grid(row=0, column=0, sticky="nsew")

        item_scroll = ttk.Scrollbar(item_host, orient="vertical", command=self.item_tree.yview)
        item_scroll.grid(row=0, column=1, sticky="ns")
        self.item_tree.configure(yscrollcommand=item_scroll.set)

        # Right panel (airdrop containers + loot editor)
        right = ctk.CTkFrame(body, fg_color=KP_PANEL, corner_radius=12)
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_rowconfigure(6, weight=1)
        right.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(right, text="Airdrop Containers", text_color=KP_TEXT, font=("Segoe UI", 16, "bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(12, 6)
        )

        # Container picker row
        picker = ctk.CTkFrame(right, fg_color="transparent")
        picker.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))
        picker.grid_columnconfigure(0, weight=1)

        self.dd_container = ctk.CTkOptionMenu(
            picker,
            values=["Load airdrop file first"],
            variable=self.var_container,
            fg_color=KP_PANEL_2,
            button_color=KP_GREEN_DARK,
            button_hover_color=KP_GREEN,
            text_color=KP_TEXT,
            dropdown_fg_color=KP_PANEL,
            dropdown_text_color=KP_TEXT,
            command=self.on_container_change
        )
        self.dd_container.grid(row=0, column=0, sticky="ew")

        ctk.CTkButton(
            picker, text="+ Add Airdrop",
            fg_color=KP_PANEL_2, hover_color="#16251a",
            text_color=KP_TEXT,
            command=self.add_airdrop
        ).grid(row=0, column=1, padx=(8, 0))

        ctk.CTkButton(
            picker, text="Clone",
            fg_color=KP_PANEL_2, hover_color="#16251a",
            text_color=KP_TEXT,
            command=self.clone_airdrop
        ).grid(row=0, column=2, padx=(8, 0))

        ctk.CTkButton(
            picker, text="Remove",
            fg_color="#4a1a1a", hover_color="#6a2222",
            text_color=KP_TEXT,
            command=self.remove_airdrop
        ).grid(row=0, column=3, padx=(8, 0))

        # Container settings
        settings = ctk.CTkFrame(right, fg_color=KP_PANEL_2, corner_radius=10)
        settings.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 8))
        settings.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(settings, text="ItemCount:", text_color=KP_MUTED).grid(row=0, column=0, sticky="w", padx=12, pady=10)
        ctk.CTkEntry(settings, textvariable=self.var_item_count, fg_color=KP_PANEL, text_color=KP_TEXT).grid(
            row=0, column=1, sticky="ew", padx=12, pady=10
        )

        ctk.CTkLabel(settings, text="Infected Enabled:", text_color=KP_MUTED).grid(row=1, column=0, sticky="w", padx=12, pady=(0, 10))
        ctk.CTkSwitch(settings, text="", variable=self.var_infected_enabled).grid(row=1, column=1, sticky="w", padx=12, pady=(0, 10))

        ctk.CTkLabel(settings, text="InfectedCount:", text_color=KP_MUTED).grid(row=2, column=0, sticky="w", padx=12, pady=(0, 10))
        ctk.CTkEntry(settings, textvariable=self.var_infected_count, fg_color=KP_PANEL, text_color=KP_TEXT).grid(
            row=2, column=1, sticky="ew", padx=12, pady=(0, 10)
        )

        ctk.CTkButton(
            settings, text="Apply Container Settings (Pending)",
            fg_color=KP_GREEN_DARK, hover_color=KP_GREEN,
            text_color="black", command=self.apply_container_settings
        ).grid(row=3, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 12))

        # Loot list
        loot_host = ctk.CTkFrame(right, fg_color=KP_PANEL_2, corner_radius=10)
        loot_host.grid(row=3, column=0, sticky="nsew", padx=12, pady=(0, 8))
        loot_host.grid_rowconfigure(0, weight=1)
        loot_host.grid_columnconfigure(0, weight=1)

        self.loot_tree = ttk.Treeview(
            loot_host,
            columns=("Name", "Chance", "Min", "Max", "Qty%"),
            show="headings",
            selectmode="browse"
        )
        for col, w in [("Name", 240), ("Chance", 70), ("Min", 50), ("Max", 50), ("Qty%", 60)]:
            self.loot_tree.heading(col, text=col)
            self.loot_tree.column(col, width=w, anchor="w" if col == "Name" else "center")
        self.loot_tree.grid(row=0, column=0, sticky="nsew")

        loot_scroll = ttk.Scrollbar(loot_host, orient="vertical", command=self.loot_tree.yview)
        loot_scroll.grid(row=0, column=1, sticky="ns")
        self.loot_tree.configure(yscrollcommand=loot_scroll.set)

        self.loot_tree.bind("<<TreeviewSelect>>", lambda _e: self.on_loot_select())

        # Loot editor
        editor = ctk.CTkFrame(right, fg_color=KP_PANEL_2, corner_radius=10)
        editor.grid(row=4, column=0, sticky="ew", padx=12, pady=(0, 12))
        editor.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(editor, text="Name:", text_color=KP_MUTED).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))
        ctk.CTkEntry(editor, textvariable=self.var_name, fg_color=KP_PANEL, text_color=KP_TEXT).grid(row=0, column=1, sticky="ew", padx=12, pady=(12, 6))

        ctk.CTkLabel(editor, text="Chance:", text_color=KP_MUTED).grid(row=1, column=0, sticky="w", padx=12, pady=6)
        ctk.CTkEntry(editor, textvariable=self.var_chance, fg_color=KP_PANEL, text_color=KP_TEXT).grid(row=1, column=1, sticky="ew", padx=12, pady=6)

        ctk.CTkLabel(editor, text="Min:", text_color=KP_MUTED).grid(row=2, column=0, sticky="w", padx=12, pady=6)
        ctk.CTkEntry(editor, textvariable=self.var_min, fg_color=KP_PANEL, text_color=KP_TEXT).grid(row=2, column=1, sticky="ew", padx=12, pady=6)

        ctk.CTkLabel(editor, text="Max:", text_color=KP_MUTED).grid(row=3, column=0, sticky="w", padx=12, pady=6)
        ctk.CTkEntry(editor, textvariable=self.var_max, fg_color=KP_PANEL, text_color=KP_TEXT).grid(row=3, column=1, sticky="ew", padx=12, pady=6)

        ctk.CTkLabel(editor, text="Qty%:", text_color=KP_MUTED).grid(row=4, column=0, sticky="w", padx=12, pady=6)
        ctk.CTkEntry(editor, textvariable=self.var_qty, fg_color=KP_PANEL, text_color=KP_TEXT).grid(row=4, column=1, sticky="ew", padx=12, pady=6)

        row_btns = ctk.CTkFrame(editor, fg_color="transparent")
        row_btns.grid(row=5, column=0, columnspan=2, sticky="ew", padx=12, pady=(6, 10))

        ctk.CTkButton(
            row_btns, text="Apply to Selected Loot",
            fg_color=KP_GREEN_DARK, hover_color=KP_GREEN,
            text_color="black",
            command=self.apply_loot_edit
        ).pack(side="left")

        ctk.CTkButton(
            row_btns, text="Remove Selected Loot",
            fg_color="#4a1a1a", hover_color="#6a2222",
            text_color=KP_TEXT,
            command=self.remove_loot
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            row_btns, text="Move Up",
            fg_color=KP_PANEL, hover_color="#16251a",
            text_color=KP_TEXT,
            command=lambda: self.move_loot(-1)
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            row_btns, text="Move Down",
            fg_color=KP_PANEL, hover_color="#16251a",
            text_color=KP_TEXT,
            command=lambda: self.move_loot(1)
        ).pack(side="left")

        self.lbl_status = ctk.CTkLabel(
            editor,
            text="Tip: ItemCount/InfectedCount/Infected commit ONLY when you click Save.",
            text_color=KP_MUTED
        )
        self.lbl_status.grid(row=7, column=0, columnspan=2, sticky="w", padx=12, pady=(0, 10))

        # ---------------- Market tab ----------------
        mkt = ctk.CTkFrame(tabs.tab("Market"), fg_color=KP_BG)
        mkt.pack(fill="both", expand=True)

        mkt.grid_rowconfigure(0, weight=1)
        mkt.grid_columnconfigure(0, weight=1)
        mkt.grid_columnconfigure(1, weight=2)

        # Left: Types browser for Market
        m_left = ctk.CTkFrame(mkt, fg_color=KP_PANEL, corner_radius=12)
        m_left.grid(row=0, column=0, sticky="nsew", padx=(12, 6), pady=12)
        m_left.grid_rowconfigure(2, weight=1)
        m_left.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            m_left,
            text="Types → Market",
            font=("Segoe UI", 16, "bold"),
            text_color=KP_TEXT
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        m_search_row = ctk.CTkFrame(m_left, fg_color="transparent")
        m_search_row.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))
        m_search_row.grid_columnconfigure(0, weight=1)

        self.var_market_search = tk.StringVar(value="")
        m_ent = ctk.CTkEntry(
            m_search_row, textvariable=self.var_market_search,
            fg_color=KP_PANEL_2, text_color=KP_TEXT,
            placeholder_text="Search classnames…",
            placeholder_text_color=KP_MUTED
        )
        m_ent.grid(row=0, column=0, sticky="ew")
        m_ent.bind("<KeyRelease>", lambda _e: self._refresh_market_types_list())

        ctk.CTkButton(
            m_search_row, text="Clear",
            width=70,
            fg_color=KP_GREEN_DARK, hover_color=KP_GREEN,
            text_color="black",
            command=lambda: (self.var_market_search.set(""), self._refresh_market_types_list())
        ).grid(row=0, column=1, padx=(8, 0))

        m_types_host = ctk.CTkFrame(m_left, fg_color=KP_PANEL_2, corner_radius=10)
        m_types_host.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 12))
        m_types_host.grid_rowconfigure(0, weight=1)
        m_types_host.grid_columnconfigure(0, weight=1)

        self.mkt_types_tree = ttk.Treeview(
            m_types_host,
            columns=("Name", "Source"),
            show="headings",
            selectmode="browse"
        )
        self.mkt_types_tree.heading("Name", text="Classname")
        self.mkt_types_tree.heading("Source", text="Types File")
        self.mkt_types_tree.column("Name", width=280, anchor="w")
        self.mkt_types_tree.column("Source", width=160, anchor="w")
        self.mkt_types_tree.grid(row=0, column=0, sticky="nsew")

        m_types_scroll = ttk.Scrollbar(m_types_host, orient="vertical", command=self.mkt_types_tree.yview)
        m_types_scroll.grid(row=0, column=1, sticky="ns")
        self.mkt_types_tree.configure(yscrollcommand=m_types_scroll.set)

        # Double click to add to market
        self.mkt_types_tree.bind("<Double-1>", lambda _e: self.market_add_selected_from_types())

        # Right: Market editor
        m_right = ctk.CTkFrame(mkt, fg_color=KP_PANEL, corner_radius=12)
        m_right.grid(row=0, column=1, sticky="nsew", padx=(6, 12), pady=12)
        m_right.grid_rowconfigure(3, weight=1)
        m_right.grid_columnconfigure(0, weight=1)

        btns = ctk.CTkFrame(m_right, fg_color="transparent")
        btns.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 8))

        ctk.CTkButton(
            btns, text="Load Market JSON",
            fg_color=KP_GREEN_DARK, hover_color=KP_GREEN,
            text_color="black",
            command=self.load_market
        ).pack(side="left")

        ctk.CTkButton(
            btns, text="New Market Category",
            fg_color=KP_PANEL_2, hover_color="#16251a",
            text_color=KP_TEXT,
            command=self.new_market
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            btns, text="Save Market",
            fg_color=KP_GREEN_DARK, hover_color=KP_GREEN,
            text_color="black",
            command=self.save_market
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            btns, text="Save As…",
            fg_color=KP_PANEL_2, hover_color="#16251a",
            text_color=KP_TEXT,
            command=self.save_market_as
        ).pack(side="left")

        self.lbl_market_path = ctk.CTkLabel(m_right, text="No market loaded.", text_color=KP_MUTED)
        self.lbl_market_path.grid(row=1, column=0, sticky="w", padx=12, pady=(0, 8))

        # Market items table
        table_host = ctk.CTkFrame(m_right, fg_color=KP_PANEL_2, corner_radius=10)
        table_host.grid(row=3, column=0, sticky="nsew", padx=12, pady=(0, 8))
        table_host.grid_rowconfigure(0, weight=1)
        table_host.grid_columnconfigure(0, weight=1)

        self.market_tree = ttk.Treeview(
            table_host,
            columns=("ClassName", "MinPrice", "MaxPrice", "MinStock", "MaxStock"),
            show="headings",
            selectmode="browse"
        )
        for col, w in [("ClassName", 260), ("MinPrice", 90), ("MaxPrice", 90), ("MinStock", 90), ("MaxStock", 90)]:
            self.market_tree.heading(col, text=col)
            self.market_tree.column(col, width=w, anchor="w" if col == "ClassName" else "center")
        self.market_tree.grid(row=0, column=0, sticky="nsew")

        m_scroll = ttk.Scrollbar(table_host, orient="vertical", command=self.market_tree.yview)
        m_scroll.grid(row=0, column=1, sticky="ns")
        self.market_tree.configure(yscrollcommand=m_scroll.set)

        self.market_tree.bind("<<TreeviewSelect>>", lambda _e: self._on_market_select())

        # Editor row
        edit = ctk.CTkFrame(m_right, fg_color="transparent")
        edit.grid(row=4, column=0, sticky="ew", padx=12, pady=(0, 8))
        for c in range(10):
            edit.grid_columnconfigure(c, weight=1 if c in (1, 3, 5, 7, 9) else 0)

        self.var_m_class = tk.StringVar(value="")
        self.var_m_minp = tk.StringVar(value="5")
        self.var_m_maxp = tk.StringVar(value="10")
        self.var_m_minst = tk.StringVar(value="0")
        self.var_m_maxst = tk.StringVar(value="100")

        def _m_lab(text, col):
            ctk.CTkLabel(edit, text=text, text_color=KP_MUTED).grid(row=0, column=col, sticky="w")

        def _m_ent(var, col):
            e = ctk.CTkEntry(edit, textvariable=var, fg_color=KP_PANEL_2, text_color=KP_TEXT, width=90)
            e.grid(row=1, column=col, sticky="ew", padx=(0, 8))
            return e

        _m_lab("Class", 0); _m_ent(self.var_m_class, 1)
        _m_lab("Min$", 2); _m_ent(self.var_m_minp, 3)
        _m_lab("Max$", 4); _m_ent(self.var_m_maxp, 5)
        _m_lab("MinStock", 6); _m_ent(self.var_m_minst, 7)
        _m_lab("MaxStock", 8); _m_ent(self.var_m_maxst, 9)

        m_actions = ctk.CTkFrame(m_right, fg_color="transparent")
        m_actions.grid(row=5, column=0, sticky="ew", padx=12, pady=(0, 12))

        ctk.CTkButton(
            m_actions, text="Add Selected Types → Market",
            fg_color=KP_GREEN_DARK, hover_color=KP_GREEN,
            text_color="black",
            command=self.market_add_selected_from_types
        ).pack(side="left")

        ctk.CTkButton(
            m_actions, text="Apply Edit",
            fg_color=KP_PANEL_2, hover_color="#16251a",
            text_color=KP_TEXT,
            command=self.market_apply_edit
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            m_actions, text="Remove Selected",
            fg_color="#4a1a1a", hover_color="#6a2222",
            text_color=KP_TEXT,
            command=self.market_remove_selected
        ).pack(side="left")

        # ---------------- Info tab ----------------
        info = ctk.CTkFrame(tabs.tab("Info"), fg_color=KP_BG)
        info.pack(fill="both", expand=True)

        panel = ctk.CTkFrame(info, fg_color=KP_PANEL, corner_radius=12)
        panel.pack(fill="both", expand=True, padx=12, pady=12)

        # Optional logo
        logo_row = ctk.CTkFrame(panel, fg_color="transparent")
        logo_row.pack(fill="x", pady=(14, 6), padx=14)

        if Image:
            try:
                logo_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "kp_logo.png")
                if os.path.exists(logo_path):
                    img = Image.open(logo_path)
                    img = img.resize((180, 180))
                    self.logo_ctk = ctk.CTkImage(light_image=img, dark_image=img, size=(180, 180))
                    ctk.CTkLabel(logo_row, image=self.logo_ctk, text="").pack(side="left", padx=(0, 12))
            except Exception:
                pass

        text_col = ctk.CTkFrame(logo_row, fg_color="transparent")
        text_col.pack(side="left", fill="both", expand=True)

        ctk.CTkLabel(
            text_col,
            text="KPTools – Expansion Airdrop Loot Builder",
            font=("Segoe UI", 20, "bold"),
            text_color=KP_TEXT
        ).pack(anchor="w")

        ctk.CTkLabel(
            text_col,
            text=f"Version: {APP_VERSION}",
            text_color=KP_MUTED
        ).pack(anchor="w", pady=(2, 0))

        ctk.CTkLabel(
            panel,
            text="Discord / Support",
            font=("Segoe UI", 16, "bold"),
            text_color=KP_TEXT
        ).pack(anchor="w", padx=14, pady=(14, 6))

        ctk.CTkLabel(
            panel,
            text="If you need custom DayZ work done, contact KP_Creates.\nJoin the Discord for support + updates:",
            text_color=KP_MUTED,
            justify="left"
        ).pack(anchor="w", padx=14)

        ctk.CTkButton(
            panel,
            text="Open Discord Invite",
            fg_color=KP_GREEN_DARK, hover_color=KP_GREEN,
            text_color="black",
            command=lambda: webbrowser.open(DISCORD_URL)
        ).pack(anchor="w", padx=14, pady=(10, 0))

        ctk.CTkLabel(
            panel,
            text="Notes:\n- types.xml browsing is optimized, but resizing can still cause minor UI redraw lag.\n- Save creates a timestamped .bak backup automatically.",
            text_color=KP_MUTED,
            justify="left"
        ).pack(anchor="w", padx=14, pady=(14, 0))

    # ---------------- Styling ----------------

    def _apply_tree_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("Treeview",
                        background=TV_BG,
                        fieldbackground=TV_BG,
                        foreground=KP_TEXT,
                        rowheight=22,
                        borderwidth=0)
        style.map("Treeview",
                  background=[("selected", TV_SEL)],
                  foreground=[("selected", KP_TEXT)])

        style.configure("Treeview.Heading",
                        background=TV_HEAD_BG,
                        foreground=KP_TEXT,
                        relief="flat")
        style.map("Treeview.Heading",
                  background=[("active", TV_HEAD_BG)])

        # tag colors
        self._tv_tags_ready = True

    # ---------------- Helpers ----------------

    def _cancel_item_job(self):
        if self._item_job is not None:
            try:
                self.after_cancel(self._item_job)
            except Exception:
                pass
            self._item_job = None

    def _cancel_loot_job(self):
        if self._loot_job is not None:
            try:
                self.after_cancel(self._loot_job)
            except Exception:
                pass
            self._loot_job = None

    def _update_paths_label(self):
        ap = os.path.basename(self.airdrop_path) if self.airdrop_path else "None"
        tf = f"{len(self.types_files_loaded)} file(s)" if self.types_files_loaded else "None"
        mk = os.path.basename(self.market_path) if getattr(self, "market_path", "") else "None"
        self.lbl_paths.configure(
            text=f"Airdrop: {ap} | Types: {tf} | Market: {mk} | Config: %APPDATA%\\KPTools\\ExpansionAirdropLootBuilder"
        )
        self._update_market_path_label()

    # ---------------- Types loading ----------------

    def load_types(self):
        folder = filedialog.askdirectory(
            title="Select a folder containing types.xml (can include subfolders)"
        )
        if not folder:
            return
        # allow picking multiple times; merges into the same database
        self._load_types_folder_silent(folder)

    def _load_types_folder_silent(self, folder: str, show_warnings: bool = False):
        if folder in self.types_folders:
            return

        xmls = find_xml_files_in_folder(folder)
        if not xmls:
            if show_warnings:
                messagebox.showwarning("No XML found", "No .xml files found in that folder.")
            return

        self.types_folders.append(folder)
        self._merge_types_files(xmls, show_warnings=show_warnings)
        self._remember_types_folders(self.types_folders)
        self._update_paths_label()

    def clear_types(self):
        self._cancel_item_job()
        self.types_folders = []
        self.types_files_loaded = []
        self.all_items = []
        self.item_meta.clear()
        self.source_to_items.clear()

        self.dd_source.configure(values=["All Types Files"])
        self.var_source.set("All Types Files")
        self.clear_search()
        self._remember_types_folders([])
        self._update_paths_label()
        self._refresh_market_types_list()

    def _merge_types_files(self, paths, show_warnings: bool = True):
        merged_items = set(self.all_items)
        errors = []

        for p in paths:
            try:
                items, meta = load_types_xml(p)
                src = os.path.basename(p)

                self.source_to_items[src] = items
                if p not in self.types_files_loaded:
                    self.types_files_loaded.append(p)

                for k, v in meta.items():
                    if k not in self.item_meta:
                        self.item_meta[k] = v

                for it in items:
                    merged_items.add(it)
            except Exception as e:
                errors.append(f"{os.path.basename(p)}: {e}")

        self.all_items = sorted(merged_items, key=str.lower)

        sources = ["All Types Files"] + sorted(self.source_to_items.keys(), key=str.lower)
        self.dd_source.configure(values=sources)
        self.var_source.set("All Types Files")
        self.refresh_item_filter()
        self._refresh_market_types_list()

        if errors and show_warnings:
            messagebox.showwarning("Some files failed to load", "A few types files failed:\n\n" + "\n".join(errors))

    def clear_search(self):
        self.var_search.set("")
        self.refresh_item_filter()

    def refresh_item_filter(self):
        self._cancel_item_job()
        # debounce redraw
        self._item_job = self.after(75, self._item_insert_step)

    def _get_filtered_items(self):
        q = (self.var_search.get() or "").strip().lower()
        src = self.var_source.get() or "All Types Files"

        if not self.all_items:
            return []

        if src == "All Types Files":
            base = self.all_items
        else:
            base = self.source_to_items.get(src, [])

        if not q:
            return base

        out = []
        for name in base:
            if q in name.lower():
                out.append(name)
        return out

    def _reset_tree_zebra(self, tree):
        # re-tag all rows for zebra striping
        for i, iid in enumerate(tree.get_children("")):
            tree.item(iid, tags=("odd" if i % 2 else "even",))

    def _item_insert_step(self):
        self._item_job = None

        # clear
        self.item_tree.delete(*self.item_tree.get_children())

        items = self._get_filtered_items()
        if not items:
            return

        # render limited rows for perf (search narrows it)
        max_rows = 1200 if (self.var_search.get() or "").strip() else 600
        count = 0
        for name in items:
            meta = self.item_meta.get(name, {})
            src = meta.get("source", "")
            self.item_tree.insert("", "end", values=(name, src))
            count += 1
            if count >= max_rows:
                break

        self._reset_tree_zebra(self.item_tree)

    # ---------------- Airdrop file ----------------

    def load_airdrop(self):
        path = filedialog.askopenfilename(
            title="Select AirdropSettings.json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return
        self._load_airdrop_path(path, silent=False)

    def _load_airdrop_path(self, path: str, silent: bool = False):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.airdrop_data = data
            self.airdrop_path = path
            self._remember_airdrop_path(path)

            self._index_containers()
            self._refresh_container_dropdown()

            self._update_paths_label()
            if not silent:
                messagebox.showinfo("Loaded", f"Airdrop file loaded:\n{os.path.basename(path)}")
        except Exception as e:
            if not silent:
                messagebox.showerror("Load Error", str(e))

    def save_airdrop(self):
        if not self.airdrop_data or not self.airdrop_path:
            messagebox.showwarning("No file", "Load an AirdropSettings.json first.")
            return

        # Apply pending container settings before save
        self._commit_container_settings_to_data()

        # Backup
        ts = time.strftime("%Y%m%d_%H%M%S")
        bak = f"{self.airdrop_path}.bak_{ts}"
        try:
            shutil.copy2(self.airdrop_path, bak)
        except Exception:
            bak = None

        try:
            with open(self.airdrop_path, "w", encoding="utf-8") as f:
                json.dump(self.airdrop_data, f, indent=2)
            msg = "Saved successfully."
            if bak:
                msg += f"\nBackup: {os.path.basename(bak)}"
            messagebox.showinfo("Saved", msg)
        except Exception as e:
            messagebox.showerror("Save Error", str(e))

    # ---------------- Containers / Loot ----------------

    def _index_containers(self):
        """
        Builds:
          self.container_names: list of display names for dropdown
          self.container_index: mapping dropdown label -> actual container object reference info
        """
        self.container_names = []
        self.container_index = {}

        if not isinstance(self.airdrop_data, dict):
            return

        containers = None
        # Try common schema keys
        for key in ("AirdropContainers", "Containers", "m_AirdropContainers"):
            if key in self.airdrop_data and isinstance(self.airdrop_data[key], list):
                containers = self.airdrop_data[key]
                break

        if containers is None:
            # fallback: scan for list of dicts w/ Loot
            for k, v in self.airdrop_data.items():
                if isinstance(v, list) and v and isinstance(v[0], dict) and ("Loot" in v[0] or "m_Loot" in v[0]):
                    containers = v
                    break

        if containers is None:
            return

        for idx, c in enumerate(containers):
            if not isinstance(c, dict):
                continue

            name = c.get("Name") or c.get("ContainerName") or c.get("m_Name") or f"Container_{idx}"
            # Some Expansion setups use "Container" classname for the crate type
            container_class = c.get("Container") or c.get("ContainerType") or c.get("m_Container") or ""
            label = f"{name}"
            if container_class:
                label = f"{name}  ({container_class})"

            self.container_names.append(label)
            self.container_index[label] = {"ref": c, "idx": idx, "list_ref": containers}

        self.container_names.sort(key=str.lower)

    def _refresh_container_dropdown(self):
        if not self.container_names:
            self.dd_container.configure(values=["Load airdrop file first"])
            self.var_container.set("Load airdrop file first")
            self.current_container_key = None
            self._clear_loot_tree()
            return

        self.dd_container.configure(values=self.container_names)
        # keep selection if possible
        if self.current_container_key in self.container_names:
            self.var_container.set(self.current_container_key)
        else:
            self.var_container.set(self.container_names[0])
            self.current_container_key = self.container_names[0]

        self.on_container_change(self.var_container.get())

    def on_container_change(self, selected):
        if not selected or selected not in self.container_index:
            self.current_container_key = None
            self._clear_loot_tree()
            return

        self.current_container_key = selected
        container = self.container_index[selected]["ref"]

        # pull settings to UI (these are NOT applied until Save)
        itemcount = container.get("ItemCount", container.get("m_ItemCount", ""))
        infected = container.get("Infected", container.get("m_Infected", True))
        infcount = container.get("InfectedCount", container.get("m_InfectedCount", ""))

        self.var_item_count.set("" if itemcount is None else str(itemcount))
        self.var_infected_enabled.set(bool(infected))
        self.var_infected_count.set("" if infcount is None else str(infcount))

        self._load_loot_tree(container)

    def _clear_loot_tree(self):
        self.loot_tree.delete(*self.loot_tree.get_children())
        self.selected_loot_index = None
        self._clear_loot_editor()

    def _clear_loot_editor(self):
        self.var_name.set("")
        self.var_chance.set("")
        self.var_min.set("")
        self.var_max.set("")
        self.var_qty.set("")

    def _get_container_loot_list(self, container: dict):
        # loot key variations
        for key in ("Loot", "m_Loot"):
            if key in container and isinstance(container[key], list):
                return container[key], key
        # create if missing
        container["Loot"] = []
        return container["Loot"], "Loot"

    def _load_loot_tree(self, container: dict):
        self._cancel_loot_job()
        self._clear_loot_tree()
        loot, _loot_key = self._get_container_loot_list(container)
        self._loot_cache = loot
        self._loot_job = self.after(50, self._loot_insert_step)

    def _loot_insert_step(self):
        self._loot_job = None

        loot = getattr(self, "_loot_cache", [])
        self.loot_tree.delete(*self.loot_tree.get_children())

        if not loot:
            return

        for row in loot:
            if not isinstance(row, dict):
                continue
            name = row.get("Name", row.get("m_Name", ""))
            chance = row.get("Chance", row.get("m_Chance", 0))
            mn = row.get("Min", row.get("m_Min", 0))
            mx = row.get("Max", row.get("m_Max", 0))
            qty = row.get("QuantityPercent", row.get("m_QuantityPercent", 0))
            self.loot_tree.insert("", "end", values=(name, chance, mn, mx, qty))

        self._reset_tree_zebra(self.loot_tree)

    def on_loot_select(self):
        sel = self.loot_tree.selection()
        if not sel:
            return
        iid = sel[0]
        vals = self.loot_tree.item(iid, "values")
        if not vals:
            return

        # Find index in loot list by matching Name + fields (best-effort)
        name = str(vals[0])
        chance = str(vals[1])
        mn = str(vals[2])
        mx = str(vals[3])
        qty = str(vals[4])

        idx = None
        container = self._get_current_container()
        if container:
            loot, _ = self._get_container_loot_list(container)
            for i, row in enumerate(loot):
                if not isinstance(row, dict):
                    continue
                rname = str(row.get("Name", row.get("m_Name", "")))
                if rname != name:
                    continue
                # best effort match
                if str(row.get("Chance", row.get("m_Chance", ""))) == chance and str(row.get("Min", row.get("m_Min", ""))) == mn:
                    idx = i
                    break
            if idx is None:
                # fallback: first matching name
                for i, row in enumerate(loot):
                    if isinstance(row, dict) and str(row.get("Name", row.get("m_Name", ""))) == name:
                        idx = i
                        break

        self.selected_loot_index = idx

        self.var_name.set(name)
        self.var_chance.set(chance)
        self.var_min.set(mn)
        self.var_max.set(mx)
        self.var_qty.set(qty)

    def _get_current_container(self):
        if not self.current_container_key or self.current_container_key not in self.container_index:
            return None
        return self.container_index[self.current_container_key]["ref"]

    def add_selected_item_to_loot(self):
        container = self._get_current_container()
        if not container:
            messagebox.showwarning("No container", "Load an airdrop file and select a container first.")
            return

        sel = self.item_tree.selection()
        if not sel:
            messagebox.showinfo("Select item", "Select a classname in the Item Browser first.")
            return

        classname = self.item_tree.item(sel[0], "values")[0]
        if not classname:
            return

        loot, loot_key = self._get_container_loot_list(container)

        # Default loot row
        row = {
            "Name": classname,
            "Chance": 1.0,
            "Min": 0,
            "Max": 1,
            "QuantityPercent": 0
        }
        loot.append(row)
        container[loot_key] = loot

        self._load_loot_tree(container)

    def apply_loot_edit(self):
        container = self._get_current_container()
        if not container:
            return
        if self.selected_loot_index is None:
            messagebox.showinfo("No selection", "Select a loot entry first.")
            return

        loot, loot_key = self._get_container_loot_list(container)
        if self.selected_loot_index < 0 or self.selected_loot_index >= len(loot):
            return

        row = loot[self.selected_loot_index]
        if not isinstance(row, dict):
            return

        row["Name"] = self.var_name.get().strip()
        # Keep floats as floats where possible
        try:
            row["Chance"] = float(self.var_chance.get())
        except Exception:
            row["Chance"] = self.var_chance.get()

        row["Min"] = safe_int(self.var_min.get(), 0)
        row["Max"] = safe_int(self.var_max.get(), 1)
        row["QuantityPercent"] = safe_int(self.var_qty.get(), 0)

        loot[self.selected_loot_index] = row
        container[loot_key] = loot

        self._load_loot_tree(container)

    def remove_loot(self):
        container = self._get_current_container()
        if not container:
            return
        if self.selected_loot_index is None:
            return

        loot, loot_key = self._get_container_loot_list(container)
        if 0 <= self.selected_loot_index < len(loot):
            del loot[self.selected_loot_index]
        container[loot_key] = loot

        self.selected_loot_index = None
        self._clear_loot_editor()
        self._load_loot_tree(container)

    def move_loot(self, direction: int):
        container = self._get_current_container()
        if not container or self.selected_loot_index is None:
            return
        loot, loot_key = self._get_container_loot_list(container)
        i = self.selected_loot_index
        j = i + direction
        if j < 0 or j >= len(loot):
            return
        loot[i], loot[j] = loot[j], loot[i]
        container[loot_key] = loot
        self.selected_loot_index = j
        self._load_loot_tree(container)
        # reselect visually
        kids = self.loot_tree.get_children()
        if 0 <= j < len(kids):
            self.loot_tree.selection_set(kids[j])
            self.loot_tree.see(kids[j])

    def apply_container_settings(self):
        # This just updates the UI status (actual commit happens on Save)
        self.lbl_status.configure(text="Container settings staged. They will be written when you click Save.", text_color=KP_GREEN)

    def _commit_container_settings_to_data(self):
        container = self._get_current_container()
        if not container:
            return

        # write values to container dict
        try:
            ic = self.var_item_count.get().strip()
            if ic == "":
                # remove if exists
                if "ItemCount" in container:
                    container.pop("ItemCount", None)
                if "m_ItemCount" in container:
                    container.pop("m_ItemCount", None)
            else:
                container["ItemCount"] = safe_int(ic, 0)
        except Exception:
            pass

        try:
            container["Infected"] = bool(self.var_infected_enabled.get())
        except Exception:
            pass

        try:
            infc = self.var_infected_count.get().strip()
            if infc == "":
                if "InfectedCount" in container:
                    container.pop("InfectedCount", None)
                if "m_InfectedCount" in container:
                    container.pop("m_InfectedCount", None)
            else:
                container["InfectedCount"] = safe_int(infc, 0)
        except Exception:
            pass

    def add_airdrop(self):
        if not isinstance(self.airdrop_data, dict):
            messagebox.showwarning("No file", "Load an AirdropSettings.json first.")
            return

        name = simpledialog.askstring("Add Airdrop", "New container Name:")
        if not name:
            return

        container_class = simpledialog.askstring("Add Airdrop", "Container classname (optional):", initialvalue="")
        # Find list of containers
        any_key = None
        containers = None
        for key in ("AirdropContainers", "Containers", "m_AirdropContainers"):
            if key in self.airdrop_data and isinstance(self.airdrop_data[key], list):
                any_key = key
                containers = self.airdrop_data[key]
                break
        if containers is None:
            messagebox.showerror("Schema", "Could not find container list in this AirdropSettings.json.")
            return

        new_c = {
            "Name": name,
            "Container": container_class or "",
            "ItemCount": 10,
            "Infected": True,
            "InfectedCount": 0,
            "Loot": []
        }
        containers.append(new_c)
        self.airdrop_data[any_key] = containers

        self._index_containers()
        self._refresh_container_dropdown()

    def clone_airdrop(self):
        container = self._get_current_container()
        if not container:
            return
        name = simpledialog.askstring("Clone Airdrop", "New cloned Name:", initialvalue=f"{container.get('Name','Clone')}_Copy")
        if not name:
            return

        # find container list ref
        list_ref = self.container_index[self.current_container_key]["list_ref"]
        new_c = json.loads(json.dumps(container))  # deep copy
        new_c["Name"] = name
        list_ref.append(new_c)

        self._index_containers()
        self._refresh_container_dropdown()

    def remove_airdrop(self):
        container = self._get_current_container()
        if not container:
            return
        if not messagebox.askyesno("Remove Airdrop", "Delete this airdrop container?"):
            return

        info = self.container_index.get(self.current_container_key)
        if not info:
            return
        list_ref = info["list_ref"]
        idx = info["idx"]
        try:
            del list_ref[idx]
        except Exception:
            # fallback remove by identity
            try:
                list_ref.remove(container)
            except Exception:
                return

        self.current_container_key = None
        self._index_containers()
        self._refresh_container_dropdown()
    # ---------------- Market methods ----------------

    def _remember_market_path(self, path: str):
        self.cfg["last_market_path"] = path
        try:
            self.cfg["last_market_dir"] = os.path.dirname(path)
        except Exception:
            pass
        self._save_config()

    def _update_market_path_label(self):
        if hasattr(self, "lbl_market_path"):
            if self.market_path:
                self.lbl_market_path.configure(text=f"Market: {os.path.basename(self.market_path)}")
            elif isinstance(self.market_data, dict):
                self.lbl_market_path.configure(text="Market: (unsaved new category)")
            else:
                self.lbl_market_path.configure(text="No market loaded.")

    def _load_market_path(self, path: str, silent: bool = False):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict) or "Items" not in data:
                raise ValueError("Not a valid Expansion Market category JSON (missing Items).")
            self.market_data = data
            self.market_path = path
            self.market_dirty = False
            self._remember_market_path(path)
            self._update_market_path_label()
            self.refresh_market_tree()
            if not silent:
                messagebox.showinfo("Market Loaded", os.path.basename(path))
        except Exception as e:
            if not silent:
                messagebox.showerror("Market Load Error", str(e))

    def load_market(self):
        initialdir = self.cfg.get("last_market_dir") or ""
        path = filedialog.askopenfilename(
            title="Select Expansion Market category JSON",
            initialdir=initialdir if os.path.isdir(initialdir) else None,
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return
        self._load_market_path(path, silent=False)

    def new_market(self):
        display = simpledialog.askstring("New Market", "DisplayName (what shows in trader):", initialvalue="New Category")
        if display is None:
            return
        file_hint = simpledialog.askstring("New Market", "Filename hint (e.g. Ammo.json):", initialvalue="NewCategory.json")
        if file_hint is None:
            return

        self.market_data = {
            "m_Version": 3,
            "DisplayName": display,
            "Icon": "Deliver",
            "Color": "FBFCFEFF",
            "IsExchange": False,
            "InitStockPercent": 75.0,
            "Items": []
        }
        self.market_path = ""
        self.market_dirty = True
        self.cfg["new_market_filename_hint"] = file_hint
        self._save_config()
        self._update_market_path_label()
        self.refresh_market_tree()

    def save_market(self):
        if not isinstance(self.market_data, dict):
            messagebox.showwarning("No Market", "Load or create a market category first.")
            return
        if not self.market_path:
            self.save_market_as()
            return
        try:
            if os.path.isfile(self.market_path):
                ts = time.strftime("%Y%m%d_%H%M%S")
                bak = f"{self.market_path}.bak_{ts}"
                try:
                    shutil.copy2(self.market_path, bak)
                except Exception:
                    bak = None
            else:
                bak = None

            with open(self.market_path, "w", encoding="utf-8") as f:
                json.dump(self.market_data, f, indent=2)

            self.market_dirty = False
            self._remember_market_path(self.market_path)
            self._update_market_path_label()
            msg = f"Saved: {os.path.basename(self.market_path)}"
            if bak:
                msg += f"\nBackup: {os.path.basename(bak)}"
            messagebox.showinfo("Market Saved", msg)
        except Exception as e:
            messagebox.showerror("Market Save Error", str(e))

    def save_market_as(self):
        if not isinstance(self.market_data, dict):
            messagebox.showwarning("No Market", "Load or create a market category first.")
            return

        start_dir = self.cfg.get("last_market_dir") or ""
        if not os.path.isdir(start_dir):
            try:
                if self.airdrop_path and os.path.isfile(self.airdrop_path):
                    start_dir = os.path.dirname(self.airdrop_path)
            except Exception:
                start_dir = ""

        hint = self.cfg.get("new_market_filename_hint") or "Market.json"
        path = filedialog.asksaveasfilename(
            title="Save Market JSON As",
            initialdir=start_dir if os.path.isdir(start_dir) else None,
            initialfile=hint,
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return
        self.market_path = path
        self._remember_market_path(path)
        self.save_market()

    def _refresh_market_types_list(self):
        if not hasattr(self, "mkt_types_tree"):
            return
        for it in self.mkt_types_tree.get_children():
            self.mkt_types_tree.delete(it)

        q = (self.var_market_search.get() or "").strip().lower()
        if not self.all_items:
            return

        max_rows = 1200 if q else 600
        count = 0
        for name in self.all_items:
            if q and q not in name.lower():
                continue
            src = ""
            meta = self.item_meta.get(name)
            if isinstance(meta, dict):
                src = meta.get("source", "")
            self.mkt_types_tree.insert("", "end", values=(name, src))
            count += 1
            if count >= max_rows:
                break

    def refresh_market_tree(self):
        if not hasattr(self, "market_tree"):
            return
        for it in self.market_tree.get_children():
            self.market_tree.delete(it)

        if not isinstance(self.market_data, dict):
            return
        items = self.market_data.get("Items") or []
        if not isinstance(items, list):
            return

        for row in items:
            if not isinstance(row, dict):
                continue
            cn = str(row.get("ClassName", ""))
            minp = row.get("MinPriceThreshold", "")
            maxp = row.get("MaxPriceThreshold", "")
            minst = row.get("MinStockThreshold", "")
            maxst = row.get("MaxStockThreshold", "")
            self.market_tree.insert("", "end", values=(cn, minp, maxp, minst, maxst))

    def _on_market_select(self):
        sel = self.market_tree.selection()
        if not sel:
            return
        vals = self.market_tree.item(sel[0], "values")
        if not vals:
            return
        self.var_m_class.set(vals[0])
        self.var_m_minp.set(str(vals[1]))
        self.var_m_maxp.set(str(vals[2]))
        self.var_m_minst.set(str(vals[3]))
        self.var_m_maxst.set(str(vals[4]))

    def _market_find_index_by_class(self, classname: str):
        if not isinstance(self.market_data, dict):
            return None
        items = self.market_data.get("Items") or []
        if not isinstance(items, list):
            return None
        target = classname.lower()
        for i, row in enumerate(items):
            if isinstance(row, dict) and str(row.get("ClassName", "")).lower() == target:
                return i
        return None

    def market_add_selected_from_types(self):
        if not isinstance(self.market_data, dict):
            messagebox.showwarning("No Market", "Load or create a market category first.")
            return
        sel = self.mkt_types_tree.selection()
        if not sel:
            messagebox.showinfo("Select Item", "Select a classname from the Types list first.")
            return
        classname = self.mkt_types_tree.item(sel[0], "values")[0]

        if self._market_find_index_by_class(classname) is not None:
            messagebox.showinfo("Already Exists", f"{classname} is already in this market category.")
            return

        maxp = safe_int(self.var_m_maxp.get(), 10)
        minp = safe_int(self.var_m_minp.get(), max(1, maxp // 2))
        maxst = safe_int(self.var_m_maxst.get(), 100)
        minst = safe_int(self.var_m_minst.get(), 0)

        entry = {
            "ClassName": classname,
            "MaxPriceThreshold": maxp,
            "MinPriceThreshold": minp,
            "SellPricePercent": -1,
            "MaxStockThreshold": maxst,
            "MinStockThreshold": minst,
            "QuantityPercent": -1,
            "SpawnAttachments": [],
            "Variants": []
        }
        self.market_data["Items"].append(entry)
        self.market_dirty = True
        self.refresh_market_tree()

    def market_apply_edit(self):
        if not isinstance(self.market_data, dict):
            return
        classname = (self.var_m_class.get() or "").strip()
        if not classname:
            messagebox.showinfo("No Item", "Select a market item first.")
            return
        idx = self._market_find_index_by_class(classname)
        if idx is None:
            messagebox.showinfo("Not Found", "Selected classname not found in market list.")
            return
        row = self.market_data["Items"][idx]
        row["ClassName"] = classname
        row["MinPriceThreshold"] = safe_int(self.var_m_minp.get(), 5)
        row["MaxPriceThreshold"] = safe_int(self.var_m_maxp.get(), 10)
        row["MinStockThreshold"] = safe_int(self.var_m_minst.get(), 0)
        row["MaxStockThreshold"] = safe_int(self.var_m_maxst.get(), 100)

        row.setdefault("SellPricePercent", -1)
        row.setdefault("QuantityPercent", -1)
        row.setdefault("SpawnAttachments", [])
        row.setdefault("Variants", [])

        self.market_dirty = True
        self.refresh_market_tree()

    def market_remove_selected(self):
        if not isinstance(self.market_data, dict):
            return
        sel = self.market_tree.selection()
        if not sel:
            return
        vals = self.market_tree.item(sel[0], "values")
        if not vals:
            return
        classname = vals[0]
        idx = self._market_find_index_by_class(classname)
        if idx is None:
            return
        del self.market_data["Items"][idx]
        self.market_dirty = True
        self.refresh_market_tree()

# ---------------- Run ----------------

if __name__ == "__main__":
    app = ExpansionAirdropLootBuilder()
    app.mainloop()
