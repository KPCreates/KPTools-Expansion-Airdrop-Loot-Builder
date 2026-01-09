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


APP_VERSION = "v1.8"

# KPTools-ish vibe
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


def ts_backup_name(path: str) -> str:
    stamp = time.strftime("%Y-%m-%d_%H%M%S")
    return f"{path}.bak_{stamp}"


def safe_float(s: str, default: float = 0.0) -> float:
    try:
        return float(s)
    except Exception:
        return default


def safe_int(s: str, default: int = 0) -> int:
    try:
        return int(float(s))
    except Exception:
        return default


@dataclass
class ItemMeta:
    category: str
    nominal: str
    min: str
    tags: list
    source: str


def load_types_xml(path: str):
    tree = ET.parse(path)
    root = tree.getroot()

    classnames = []
    meta = {}

    source_name = os.path.basename(path)

    for t in root.findall("type"):
        name = t.get("name")
        if not name:
            continue

        category = (t.findtext("category") or "").strip()
        nominal = (t.findtext("nominal") or "").strip()
        minimum = (t.findtext("min") or "").strip()
        tags = [x.text.strip() for x in t.findall("tag") if x.text and x.text.strip()]

        classnames.append(name)
        meta[name] = ItemMeta(
            category=category,
            nominal=nominal,
            min=minimum,
            tags=tags,
            source=source_name
        )

    classnames = sorted(set(classnames), key=str.lower)
    return classnames, meta


def load_airdrop_json(path: str):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if "Containers" not in data or not isinstance(data["Containers"], list):
        raise ValueError("This AirdropSettings.json does not have a 'Containers' array.")

    # Ensure Loot exists & is list (without touching any other fields)
    for c in data["Containers"]:
        if "Container" not in c:
            raise ValueError("A container entry is missing the 'Container' key.")
        if "Loot" not in c:
            c["Loot"] = []
        if not isinstance(c.get("Loot"), list):
            c["Loot"] = []

    return data


def default_loot_entry(classname: str) -> dict:
    return {
        "Name": classname,
        "Chance": 0.10,
        "Attachments": [],
        "QuantityPercent": -1.0,
        "Max": -1,
        "Min": 0,
        "Variants": []
    }


def find_xml_files_in_folder(folder: str):
    out = []
    for root, _dirs, files in os.walk(folder):
        for fn in files:
            if fn.lower().endswith(".xml"):
                out.append(os.path.join(root, fn))
    return sorted(out, key=str.lower)


class AirdropLootBuilder(ctk.CTk):
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

        # State
        self.airdrop_path = ""
        self.airdrop_data = None

        self.types_folders = []
        self.types_files_loaded = []
        self.all_items = []
        self.item_meta = {}
        self.source_to_items = {}

        self.container_names = []
        self.current_container_index = None
        self.selected_loot_index = None

        # Pending container-level edits (committed ONLY on Save)
        self.pending_container_settings = {}

        # Chunked insertion state
        self._loot_insert_job = None
        self._loot_insert_idx = 0
        self._loot_insert_list = None

        # UI vars
        self.var_search = ctk.StringVar(value="")
        self.var_source = ctk.StringVar(value="All Types Files")
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
        self._update_paths_label()
        self._render_items(["(Add a folder with types.xml files to browse items)"])

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
            top, text="Add Types Folder",
            fg_color=KP_GREEN_DARK, hover_color=KP_GREEN,
            text_color="black", command=self.add_types_folder
        ).pack(side="left", padx=10, pady=10)

        ctk.CTkButton(
            top, text="Clear Types",
            fg_color="#1a1a1a", hover_color="#2a2a2a",
            text_color=KP_TEXT, command=self.clear_types
        ).pack(side="left", padx=10, pady=10)

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
        tabs.add("Info")

        # --- Airdrops tab root ---
        body = ctk.CTkFrame(tabs.tab("Airdrops"), fg_color=KP_BG)
        body.pack(fill="both", expand=True)

        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=3)
        body.grid_rowconfigure(0, weight=1)

        # Left panel (items)
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
        ent.bind("<KeyRelease>", lambda e: self.refresh_item_filter())

        ctk.CTkButton(
            search_row, text="Clear",
            fg_color=KP_PANEL_2, hover_color="#16251a",
            text_color=KP_TEXT, width=80,
            command=self.clear_search
        ).grid(row=0, column=1, padx=(8, 0))

        self.items_frame = ctk.CTkScrollableFrame(left, fg_color=KP_PANEL_2, corner_radius=10)
        self.items_frame.grid(row=4, column=0, sticky="nsew", padx=12, pady=(0, 12))

        ctk.CTkLabel(left, text="Click + to add item to selected airdrop container.", text_color=KP_MUTED).grid(
            row=5, column=0, sticky="w", padx=12, pady=(0, 12)
        )

        # Right panel (containers + loot)
        right = ctk.CTkFrame(body, fg_color=KP_PANEL, corner_radius=12)
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_rowconfigure(3, weight=1)
        right.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(right, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 8))
        header.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            header, text="Expansion Containers & Loot",
            text_color=KP_TEXT, font=("Segoe UI", 16, "bold")
        ).grid(row=0, column=0, sticky="w")

        self.dd_container = ctk.CTkOptionMenu(
            header,
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
        self.dd_container.grid(row=0, column=1, sticky="e", padx=(12, 0))

        # Container controls
        cont_controls = ctk.CTkFrame(right, fg_color="transparent")
        cont_controls.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))

        ctk.CTkButton(
            cont_controls, text="+ Add Airdrop (clone)",
            fg_color=KP_GREEN_DARK, hover_color=KP_GREEN,
            text_color="black", command=self.add_container_clone
        ).pack(side="left")

        ctk.CTkButton(
            cont_controls, text="Duplicate Airdrop",
            fg_color="#1a1a1a", hover_color="#2a2a2a",
            text_color=KP_TEXT, command=self.duplicate_container
        ).pack(side="left", padx=(8, 0))

        ctk.CTkButton(
            cont_controls, text="Remove Airdrop",
            fg_color="#2a1010", hover_color="#3a1515",
            text_color=KP_TEXT, command=self.remove_container
        ).pack(side="left", padx=(8, 0))

        ctk.CTkButton(
            cont_controls, text="Clear Loot",
            fg_color="#1a1a1a", hover_color="#2a2a2a",
            text_color=KP_TEXT, command=self.clear_current_loot
        ).pack(side="left", padx=(8, 0))

        ctk.CTkButton(
            cont_controls, text="Remove Selected Loot",
            fg_color="#1a1a1a", hover_color="#2a2a2a",
            text_color=KP_TEXT, command=self.remove_selected_loot
        ).pack(side="right", padx=(8, 0))

        ctk.CTkButton(
            cont_controls, text="Duplicate Selected Loot",
            fg_color="#1a1a1a", hover_color="#2a2a2a",
            text_color=KP_TEXT, command=self.duplicate_selected_loot
        ).pack(side="right")

        # --- Container Settings panel (PENDING, applied on SAVE) ---
        settings = ctk.CTkFrame(right, fg_color=KP_PANEL_2, corner_radius=10)
        settings.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 10))
        settings.grid_columnconfigure(1, weight=1)
        settings.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(settings, text="Airdrop Settings (applies when you Save)", text_color=KP_TEXT,
                     font=("Segoe UI", 13, "bold")).grid(row=0, column=0, columnspan=4,
                                                        sticky="w", padx=12, pady=(10, 6))

        def add_setting(row, col_label, col_entry, label, var, width=120):
            ctk.CTkLabel(settings, text=label, text_color=KP_MUTED).grid(
                row=row, column=col_label, sticky="w", padx=12, pady=6
            )
            e = ctk.CTkEntry(settings, textvariable=var, fg_color=KP_PANEL, text_color=KP_TEXT, width=width)
            e.grid(row=row, column=col_entry, sticky="w", padx=12, pady=6)
            e.bind("<FocusOut>", lambda _e: self._stash_pending_container_settings())
            e.bind("<Return>", lambda _e: self._stash_pending_container_settings())
            return e

        add_setting(1, 0, 1, "ItemCount", self.var_item_count, width=140)
        add_setting(1, 2, 3, "InfectedCount", self.var_infected_count, width=140)

        chk = ctk.CTkCheckBox(
            settings, text="Infected Enabled",
            variable=self.var_infected_enabled,
            fg_color=KP_GREEN_DARK,
            hover_color=KP_GREEN,
            text_color=KP_TEXT
        )
        chk.grid(row=2, column=0, columnspan=2, sticky="w", padx=12, pady=(0, 10))
        chk.bind("<ButtonRelease-1>", lambda _e: self.after(1, self._stash_pending_container_settings))

        self.lbl_pending = ctk.CTkLabel(settings, text="", text_color=KP_MUTED)
        self.lbl_pending.grid(row=2, column=2, columnspan=2, sticky="e", padx=12, pady=(0, 10))

        # --- Loot table (FAST) ---
        tree_host = ctk.CTkFrame(right, fg_color=KP_PANEL_2, corner_radius=10)
        tree_host.grid(row=3, column=0, sticky="nsew", padx=12, pady=(0, 10))
        tree_host.grid_rowconfigure(0, weight=1)
        tree_host.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            tree_host,
            columns=("Name", "Chance", "Min", "Max", "Qty%", "Att", "Var"),
            show="headings",
            selectmode="browse"
        )

        col_defs = [
            ("Name", 520, "w", True),
            ("Chance", 90, "e", False),
            ("Min", 60, "e", False),
            ("Max", 60, "e", False),
            ("Qty%", 80, "e", False),
            ("Att", 60, "e", False),
            ("Var", 60, "e", False),
        ]
        for col, w, anchor, stretch in col_defs:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor=anchor, stretch=stretch)

        vsb = ttk.Scrollbar(tree_host, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=8)
        vsb.grid(row=0, column=1, sticky="ns", padx=(0, 8), pady=8)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "Treeview",
            background=TV_BG,
            fieldbackground=TV_BG,
            foreground=KP_TEXT,
            rowheight=26,
            borderwidth=0
        )
        style.configure(
            "Treeview.Heading",
            background=TV_HEAD_BG,
            foreground=KP_MUTED,
            font=("Segoe UI", 10, "bold")
        )
        style.map("Treeview", background=[("selected", TV_SEL)], foreground=[("selected", KP_TEXT)])

        self.tree.tag_configure("even", background=TV_ROW_EVEN)
        self.tree.tag_configure("odd", background=TV_ROW_ODD)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # Loot editor
        editor = ctk.CTkFrame(right, fg_color=KP_PANEL_2, corner_radius=10)
        editor.grid(row=4, column=0, sticky="ew", padx=12, pady=(0, 12))
        editor.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(editor, text="Selected Loot Editor", text_color=KP_TEXT,
                     font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 6)
        )

        def add_field(r, label, var):
            ctk.CTkLabel(editor, text=label, text_color=KP_MUTED).grid(
                row=r, column=0, sticky="w", padx=12, pady=4
            )
            ctk.CTkEntry(editor, textvariable=var, fg_color=KP_PANEL, text_color=KP_TEXT).grid(
                row=r, column=1, sticky="ew", padx=12, pady=4
            )

        add_field(1, "Name", self.var_name)
        add_field(2, "Chance", self.var_chance)
        add_field(3, "Min", self.var_min)
        add_field(4, "Max", self.var_max)
        add_field(5, "QuantityPercent", self.var_qty)

        btnrow = ctk.CTkFrame(editor, fg_color="transparent")
        btnrow.grid(row=6, column=0, columnspan=2, sticky="w", padx=12, pady=(8, 6))
        ctk.CTkButton(
            btnrow, text="Apply Loot Changes",
            fg_color=KP_GREEN_DARK, hover_color=KP_GREEN,
            text_color="black", command=self.apply_editor_to_selected
        ).pack(side="left")

        self.lbl_status = ctk.CTkLabel(
            editor,
            text="Tip: ItemCount/InfectedCount/Infected commit ONLY when you click Save.",
            text_color=KP_MUTED
        )
        self.lbl_status.grid(row=7, column=0, columnspan=2, sticky="w", padx=12, pady=(0, 10))

        # ---------------- Info tab ----------------
        info = ctk.CTkFrame(tabs.tab("Info"), fg_color=KP_BG)
        info.pack(fill="both", expand=True)

        panel = ctk.CTkFrame(info, fg_color=KP_PANEL, corner_radius=12)
        panel.pack(fill="both", expand=True, padx=12, pady=12)

        # Optional logo (expects kp_logo.png next to app)
        logo_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "kp_logo.png")
        if Image is not None and os.path.exists(logo_path):
            try:
                img = Image.open(logo_path)
                img.thumbnail((360, 360))
                logo = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
                ctk.CTkLabel(panel, text="", image=logo).pack(anchor="w", padx=16, pady=(16, 8))
                self._logo_ref = logo
            except Exception:
                pass

        ctk.CTkLabel(
            panel,
            text="KPTools – Expansion Airdrop Loot Builder",
            text_color=KP_TEXT,
            font=("Segoe UI", 22, "bold")
        ).pack(anchor="w", padx=16, pady=(6, 6))

        ctk.CTkLabel(
            panel,
            text=f"Version: {APP_VERSION}\n\nNeed help or want custom DayZ work done?",
            text_color=KP_MUTED,
            font=("Segoe UI", 13)
        ).pack(anchor="w", padx=16, pady=(0, 10))

        msg = (
            "Contact me on Discord for:\n"
            "• Custom DayZ mods (weapons, clothing, buildings)\n"
            "• Server setup help (Expansion, economy, airdrops)\n"
            "• Config work + loot balancing\n"
            "• Troubleshooting crashes / scripts\n\n"
            "When you message me, include:\n"
            "• What server/mod you’re using\n"
            "• What you want built/changed\n"
            "• Any files/screenshots/logs you have\n"
        )
        ctk.CTkLabel(panel, text=msg, text_color=KP_TEXT, justify="left", font=("Segoe UI", 13)).pack(
            anchor="w", padx=16, pady=(0, 12)
        )

        ctk.CTkButton(
            panel,
            text="Join My Discord",
            fg_color=KP_GREEN_DARK,
            hover_color=KP_GREEN,
            text_color="black",
            height=40,
            command=lambda: webbrowser.open(DISCORD_URL)
        ).pack(anchor="w", padx=16, pady=(0, 10))

        link_frame = ctk.CTkFrame(panel, fg_color=KP_PANEL_2, corner_radius=10)
        link_frame.pack(fill="x", padx=16, pady=(0, 16))

        ctk.CTkLabel(link_frame, text="Discord Invite:", text_color=KP_MUTED).pack(side="left", padx=(12, 8), pady=12)

        link_entry = ctk.CTkEntry(link_frame, fg_color=KP_PANEL, text_color=KP_TEXT, width=520)
        link_entry.pack(side="left", padx=(0, 12), pady=12, fill="x", expand=True)
        link_entry.insert(0, DISCORD_URL)
        link_entry.configure(state="readonly")

    # ---------------- Pending container settings ----------------

    def _stash_pending_container_settings(self):
        if self.airdrop_data is None or self.current_container_index is None:
            return

        idx = self.current_container_index

        item_count_str = (self.var_item_count.get() or "").strip()
        infected_count_str = (self.var_infected_count.get() or "").strip()
        infected_enabled = bool(self.var_infected_enabled.get())

        pending = self.pending_container_settings.get(idx, {})

        if item_count_str != "":
            pending["ItemCount"] = safe_int(item_count_str, 0)
        else:
            pending.pop("ItemCount", None)

        if infected_count_str != "":
            pending["InfectedCount"] = safe_int(infected_count_str, 0)
        else:
            pending.pop("InfectedCount", None)

        pending["Infected"] = infected_enabled

        if pending:
            self.pending_container_settings[idx] = pending
            self.lbl_pending.configure(text="Pending changes ✓")
        else:
            self.pending_container_settings.pop(idx, None)
            self.lbl_pending.configure(text="")

    def _load_container_settings_into_ui(self):
        c = self.get_current_container()
        if not c:
            self.var_item_count.set("")
            self.var_infected_count.set("")
            self.var_infected_enabled.set(True)
            self.lbl_pending.configure(text="")
            return

        idx = self.current_container_index
        pending = self.pending_container_settings.get(idx)

        item_count = c.get("ItemCount", "")
        infected_count = c.get("InfectedCount", "")
        infected_enabled = c.get("Infected", True)

        if pending:
            item_count = pending.get("ItemCount", item_count)
            infected_count = pending.get("InfectedCount", infected_count)
            infected_enabled = pending.get("Infected", infected_enabled)
            self.lbl_pending.configure(text="Pending changes ✓")
        else:
            self.lbl_pending.configure(text="")

        self.var_item_count.set("" if item_count == "" else str(item_count))
        self.var_infected_count.set("" if infected_count == "" else str(infected_count))
        self.var_infected_enabled.set(bool(infected_enabled))

    def _commit_pending_container_settings_to_json(self):
        if not self.airdrop_data:
            return

        for idx, changes in list(self.pending_container_settings.items()):
            if idx < 0 or idx >= len(self.airdrop_data["Containers"]):
                continue
            c = self.airdrop_data["Containers"][idx]
            for k, v in changes.items():
                c[k] = v

        self.pending_container_settings.clear()
        self.lbl_pending.configure(text="")

    # ---------------- File IO ----------------

    def load_airdrop(self):
        path = filedialog.askopenfilename(
            title="Select AirdropSettings.json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return

        try:
            data = load_airdrop_json(path)
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load airdrop JSON:\n\n{e}")
            return

        self.airdrop_path = path
        self.airdrop_data = data
        self.pending_container_settings.clear()

        self._rebuild_container_dropdown(select_index=0)
        self._update_paths_label()
        self._load_container_settings_into_ui()
        self.refresh_loot_table_chunked()

    def save_airdrop(self):
        if not self.airdrop_data or not self.airdrop_path:
            messagebox.showwarning("Nothing to Save", "Load AirdropSettings.json first.")
            return

        self._stash_pending_container_settings()
        self._commit_pending_container_settings_to_json()

        try:
            bak = ts_backup_name(self.airdrop_path)
            shutil.copy2(self.airdrop_path, bak)
        except Exception as e:
            messagebox.showerror("Backup Error", f"Could not create backup:\n\n{e}")
            return

        try:
            with open(self.airdrop_path, "w", encoding="utf-8") as f:
                json.dump(self.airdrop_data, f, indent=2)
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not write JSON:\n\n{e}")
            return

        messagebox.showinfo("Saved", f"Saved successfully.\nBackup created:\n{bak}")

    # ---------------- Types loading ----------------

    def add_types_folder(self):
        folder = filedialog.askdirectory(title="Select a folder containing types.xml files")
        if not folder:
            return

        self.types_folders.append(folder)
        xmls = find_xml_files_in_folder(folder)
        if not xmls:
            messagebox.showwarning("No XML found", "No .xml files found in that folder.")
            return

        self._merge_types_files(xmls)
        self._update_paths_label()

    def clear_types(self):
        self.types_folders = []
        self.types_files_loaded = []
        self.all_items = []
        self.item_meta.clear()
        self.source_to_items.clear()

        self.dd_source.configure(values=["All Types Files"])
        self.var_source.set("All Types Files")
        self._render_items(["(Add a folder with types.xml files to browse items)"])
        self._update_paths_label()

    def _merge_types_files(self, paths):
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

                merged_items.update(items)
            except Exception as e:
                errors.append(f"{os.path.basename(p)}: {e}")

        self.all_items = sorted(merged_items, key=str.lower)

        sources = ["All Types Files"] + sorted(self.source_to_items.keys(), key=str.lower)
        self.dd_source.configure(values=sources)
        self.var_source.set("All Types Files")

        self.refresh_item_filter()

        if errors:
            messagebox.showwarning("Some files failed to load", "A few types files failed:\n\n" + "\n".join(errors))

    # ---------------- Containers ----------------

    def _rebuild_container_dropdown(self, select_index=None):
        if not self.airdrop_data:
            self.container_names = []
            self.dd_container.configure(values=["Load airdrop file first"])
            self.var_container.set("Load airdrop file first")
            self.current_container_index = None
            return

        self.container_names = [c.get("Container", f"Container_{i}") for i, c in enumerate(self.airdrop_data["Containers"])]
        values = self.container_names if self.container_names else ["(No containers)"]
        self.dd_container.configure(values=values)

        if not self.container_names:
            self.var_container.set("(No containers)")
            self.current_container_index = None
            return

        if select_index is None:
            select_index = min(max(self.current_container_index or 0, 0), len(self.container_names) - 1)

        self.current_container_index = select_index
        self.var_container.set(self.container_names[select_index])

    def add_container_clone(self):
        if not self.airdrop_data or not self.airdrop_data["Containers"]:
            messagebox.showwarning("No template", "Load a JSON with at least 1 container to clone.")
            return

        template_name = simpledialog.askstring(
            "Add Airdrop (clone)",
            "Template container to clone from (exact name):\nTip: copy/paste from dropdown."
        )
        if not template_name:
            return

        try:
            template_idx = self.container_names.index(template_name)
        except ValueError:
            messagebox.showwarning("Template not found", "That template name is not in the current container list.")
            return

        new_name = simpledialog.askstring("Add Airdrop (clone)", "New container name (unique):")
        if not new_name:
            return

        existing = {c.get("Container", "") for c in self.airdrop_data["Containers"]}
        if new_name in existing:
            messagebox.showwarning("Name exists", "That container name already exists. Pick a unique name.")
            return

        clear_loot = messagebox.askyesno("Start empty?", "Clear Loot in the new container?\n\nYes = empty Loot\nNo = copy Loot too")

        src = self.airdrop_data["Containers"][template_idx]
        clone = json.loads(json.dumps(src))
        clone["Container"] = new_name
        clone.setdefault("Loot", [])
        if clear_loot:
            clone["Loot"] = []

        self.airdrop_data["Containers"].append(clone)
        self._rebuild_container_dropdown(select_index=len(self.airdrop_data["Containers"]) - 1)

        self.selected_loot_index = None
        self._clear_editor()
        self._load_container_settings_into_ui()
        self.refresh_loot_table_chunked()

    def duplicate_container(self):
        if not self.airdrop_data or self.current_container_index is None:
            return

        src = self.airdrop_data["Containers"][self.current_container_index]
        base_name = src.get("Container", "Container")
        new_name = simpledialog.askstring("Duplicate Airdrop", f"New name for copy of '{base_name}':")
        if not new_name:
            return

        existing = {c.get("Container", "") for c in self.airdrop_data["Containers"]}
        if new_name in existing:
            messagebox.showwarning("Name exists", "That container name already exists. Pick a unique name.")
            return

        clear_loot = messagebox.askyesno("Start empty?", "Clear Loot in the new container?\n\nYes = empty Loot\nNo = copy Loot too")

        clone = json.loads(json.dumps(src))
        clone["Container"] = new_name
        clone.setdefault("Loot", [])
        if clear_loot:
            clone["Loot"] = []

        self.airdrop_data["Containers"].append(clone)
        self._rebuild_container_dropdown(select_index=len(self.airdrop_data["Containers"]) - 1)

        self.selected_loot_index = None
        self._clear_editor()
        self._load_container_settings_into_ui()
        self.refresh_loot_table_chunked()

    def remove_container(self):
        if not self.airdrop_data or self.current_container_index is None:
            return

        name = self.container_names[self.current_container_index]
        if not messagebox.askyesno("Remove Airdrop", f"Delete container '{name}'?\n\nThis cannot be undone (except via backup)."):
            return

        removed_idx = self.current_container_index
        self.pending_container_settings.pop(removed_idx, None)
        self.airdrop_data["Containers"].pop(removed_idx)

        new_pending = {}
        for idx, changes in self.pending_container_settings.items():
            if idx < removed_idx:
                new_pending[idx] = changes
            elif idx > removed_idx:
                new_pending[idx - 1] = changes
        self.pending_container_settings = new_pending

        if len(self.airdrop_data["Containers"]) == 0:
            self.current_container_index = None
            self._rebuild_container_dropdown(select_index=None)
            self.tree.delete(*self.tree.get_children())
            self._clear_editor()
            self._load_container_settings_into_ui()
            return

        new_index = min(removed_idx, len(self.airdrop_data["Containers"]) - 1)
        self._rebuild_container_dropdown(select_index=new_index)

        self.selected_loot_index = None
        self._clear_editor()
        self._load_container_settings_into_ui()
        self.refresh_loot_table_chunked()

    def clear_current_loot(self):
        c = self.get_current_container()
        if not c:
            return
        if not messagebox.askyesno("Clear Loot", "Remove ALL loot entries for this container?"):
            return
        c["Loot"] = []
        self.selected_loot_index = None
        self._clear_editor()
        self.refresh_loot_table_chunked()

    def on_container_change(self, _value=None):
        if not self.airdrop_data:
            return
        self._stash_pending_container_settings()

        selected = self.var_container.get()
        if selected in self.container_names:
            self.current_container_index = self.container_names.index(selected)
            self.selected_loot_index = None
            self._clear_editor()
            self._load_container_settings_into_ui()
            self.refresh_loot_table_chunked()

    def get_current_container(self):
        if self.airdrop_data is None or self.current_container_index is None:
            return None
        return self.airdrop_data["Containers"][self.current_container_index]

    # ---------------- Loot Treeview (FAST + chunked) ----------------

    def _cancel_loot_job(self):
        if self._loot_insert_job is not None:
            try:
                self.after_cancel(self._loot_insert_job)
            except Exception:
                pass
        self._loot_insert_job = None
        self._loot_insert_idx = 0
        self._loot_insert_list = None

    def refresh_loot_table_chunked(self):
        self._cancel_loot_job()
        self.tree.delete(*self.tree.get_children())

        c = self.get_current_container()
        if not c:
            return

        loot = c.get("Loot", [])
        if not isinstance(loot, list):
            loot = []
            c["Loot"] = loot

        self._loot_insert_list = loot
        self._loot_insert_idx = 0
        self.lbl_status.configure(text=f"Loading loot… (0 / {len(loot)})")
        self._loot_insert_job = self.after(1, self._loot_insert_step)

    def _loot_insert_step(self):
        loot = self._loot_insert_list or []
        n = len(loot)
        i = self._loot_insert_idx

        BATCH = 250
        end = min(i + BATCH, n)

        for idx in range(i, end):
            entry = loot[idx]
            name = str(entry.get("Name", ""))
            chance = entry.get("Chance", 0.0)
            mn = entry.get("Min", 0)
            mx = entry.get("Max", -1)
            qty = entry.get("QuantityPercent", -1.0)
            att = len(entry.get("Attachments", []) or [])
            var = len(entry.get("Variants", []) or [])

            tag = "even" if idx % 2 == 0 else "odd"
            self.tree.insert(
                "", "end",
                iid=str(idx),
                values=(name, chance, mn, mx, qty, att, var),
                tags=(tag,)
            )

        self._loot_insert_idx = end
        if end >= n:
            self._loot_insert_job = None
            self.lbl_status.configure(text=f"Loaded loot: {n} entries. (Settings commit on Save)")
            return

        self.lbl_status.configure(text=f"Loading loot… ({end} / {n})")
        self._loot_insert_job = self.after(1, self._loot_insert_step)

    def _on_tree_select(self, _evt=None):
        sel = self.tree.selection()
        if not sel:
            return
        try:
            idx = int(sel[0])
        except Exception:
            return
        self.selected_loot_index = idx
        self.populate_editor_from_selected()

    # ---------------- Loot operations ----------------

    def add_item_to_loot(self, classname: str):
        c = self.get_current_container()
        if not c:
            messagebox.showwarning("No Container", "Load AirdropSettings.json and select a container first.")
            return

        c.setdefault("Loot", [])
        c["Loot"].append(default_loot_entry(classname))

        idx = len(c["Loot"]) - 1
        entry = c["Loot"][idx]
        tag = "even" if idx % 2 == 0 else "odd"
        self.tree.insert("", "end", iid=str(idx), values=(
            entry.get("Name", ""),
            entry.get("Chance", 0.0),
            entry.get("Min", 0),
            entry.get("Max", -1),
            entry.get("QuantityPercent", -1.0),
            0,
            0,
        ), tags=(tag,))

        self.tree.selection_set(str(idx))
        self.tree.see(str(idx))
        self.selected_loot_index = idx
        self.populate_editor_from_selected()

    def remove_selected_loot(self):
        c = self.get_current_container()
        if not c or self.selected_loot_index is None:
            return

        loot = c.get("Loot", [])
        idx = self.selected_loot_index
        if idx < 0 or idx >= len(loot):
            return

        loot.pop(idx)
        self.selected_loot_index = None
        self._clear_editor()
        self.refresh_loot_table_chunked()

    def duplicate_selected_loot(self):
        c = self.get_current_container()
        if not c or self.selected_loot_index is None:
            return

        loot = c.get("Loot", [])
        idx = self.selected_loot_index
        if idx < 0 or idx >= len(loot):
            return

        entry = json.loads(json.dumps(loot[idx]))
        loot.append(entry)

        new_idx = len(loot) - 1
        tag = "even" if new_idx % 2 == 0 else "odd"
        att = len(entry.get("Attachments", []) or [])
        var = len(entry.get("Variants", []) or [])
        self.tree.insert("", "end", iid=str(new_idx), values=(
            entry.get("Name", ""),
            entry.get("Chance", 0.0),
            entry.get("Min", 0),
            entry.get("Max", -1),
            entry.get("QuantityPercent", -1.0),
            att,
            var
        ), tags=(tag,))

        self.tree.selection_set(str(new_idx))
        self.tree.see(str(new_idx))
        self.selected_loot_index = new_idx
        self.populate_editor_from_selected()

    # ---------------- Loot editor ----------------

    def _clear_editor(self):
        self.var_name.set("")
        self.var_chance.set("")
        self.var_min.set("")
        self.var_max.set("")
        self.var_qty.set("")

    def populate_editor_from_selected(self):
        c = self.get_current_container()
        if not c or self.selected_loot_index is None:
            self._clear_editor()
            return

        loot = c.get("Loot", [])
        idx = self.selected_loot_index
        if idx < 0 or idx >= len(loot):
            self._clear_editor()
            return

        entry = loot[idx]
        self.var_name.set(entry.get("Name", ""))
        self.var_chance.set(str(entry.get("Chance", 0.0)))
        self.var_min.set(str(entry.get("Min", 0)))
        self.var_max.set(str(entry.get("Max", -1)))
        self.var_qty.set(str(entry.get("QuantityPercent", -1.0)))

    def apply_editor_to_selected(self):
        c = self.get_current_container()
        if not c or self.selected_loot_index is None:
            return

        loot = c.get("Loot", [])
        idx = self.selected_loot_index
        if idx < 0 or idx >= len(loot):
            return

        name = self.var_name.get().strip()
        if not name:
            messagebox.showwarning("Invalid Name", "Name cannot be empty.")
            return

        if self.item_meta and name not in self.item_meta:
            if not messagebox.askyesno(
                "Unknown Classname",
                "This classname was not found in the loaded types.xml files.\n\nSave anyway?"
            ):
                return

        entry = loot[idx]
        entry["Name"] = name
        entry["Chance"] = safe_float(self.var_chance.get(), entry.get("Chance", 0.1))
        entry["Min"] = safe_int(self.var_min.get(), entry.get("Min", 0))
        entry["Max"] = safe_int(self.var_max.get(), entry.get("Max", -1))
        entry["QuantityPercent"] = safe_float(self.var_qty.get(), entry.get("QuantityPercent", -1.0))

        att = len(entry.get("Attachments", []) or [])
        var = len(entry.get("Variants", []) or [])
        self.tree.item(str(idx), values=(
            entry["Name"], entry["Chance"], entry["Min"], entry["Max"], entry["QuantityPercent"], att, var
        ))



    # ---------------- Item filter + rendering ----------------

    def clear_search(self):
        self.var_search.set("")
        self.refresh_item_filter()

    def refresh_item_filter(self):
        q = self.var_search.get().strip().lower()
        src = self.var_source.get()

        if not self.all_items:
            self._render_items(["(Add a folder with types.xml files to browse items)"])
            return

        base = self.all_items if src == "All Types Files" else self.source_to_items.get(src, [])
        filtered = base if not q else [x for x in base if q in x.lower()]
        # cap to keep UI snappy
        self._render_items(filtered[:900])

    def _render_items(self, items):
        for child in self.items_frame.winfo_children():
            child.destroy()

        if not items:
            ctk.CTkLabel(self.items_frame, text="No matches.", text_color=KP_MUTED).pack(anchor="w", padx=10, pady=10)
            return

        for name in items:
            if name.startswith("(") and name.endswith(")"):
                ctk.CTkLabel(self.items_frame, text=name, text_color=KP_MUTED).pack(anchor="w", padx=10, pady=10)
                continue

            meta = self.item_meta.get(name)
            src = meta.source if meta else ""

            row = ctk.CTkFrame(self.items_frame, fg_color=KP_PANEL, corner_radius=10)
            row.pack(fill="x", padx=8, pady=4)

            left = ctk.CTkFrame(row, fg_color="transparent")
            left.pack(side="left", fill="x", expand=True, padx=10, pady=8)

            ctk.CTkLabel(left, text=name, text_color=KP_TEXT, font=("Segoe UI", 12, "bold")).pack(anchor="w")
            if src:
                ctk.CTkLabel(left, text=src, text_color=KP_MUTED, font=("Segoe UI", 11)).pack(anchor="w")

            ctk.CTkButton(
                row, text="+",
                width=44,
                fg_color=KP_GREEN_DARK, hover_color=KP_GREEN,
                text_color="black",
                command=lambda n=name: self.add_item_to_loot(n)
            ).pack(side="right", padx=10, pady=10)

    def _update_paths_label(self):
        ap = os.path.basename(self.airdrop_path) if self.airdrop_path else "None"
        tf = len(self.types_files_loaded)
        td = len(self.types_folders)
        types_label = "None" if tf == 0 else f"{tf} xml files (from {td} folder(s))"
        self.lbl_paths.configure(text=f"Airdrop: {ap} | Types: {types_label}")


if __name__ == "__main__":
    app = AirdropLootBuilder()
    app.mainloop()
