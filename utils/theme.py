from tkinter import ttk

COLORS = {
    "bg":       "#1a1b2e",  "surface":  "#222340",  "surface2": "#2a2b4a",
    "bar":      "#16172b",  "accent":   "#6c63ff",  "accent_dk":"#5a52e0",
    "text":     "#e2e4f0",  "dim":      "#6b6d85",  "border":   "#2f3055",
    "sel":      "#6c63ff",  "alt":      "#272845",  "hdr_bg":   "#1e1f38",
    "green":    "#4ade80",  "orange":   "#fb923c",  "red":      "#f87171",
    "cyan":     "#67e8f9",  "purple":   "#c084fc",  "chk_sel":  "#222340",
    "input_bg": "#2a2b4a",  "input_bd": "#3d3e65",  "badge_bg": "#2a2b4a",
}

FONT_NAME  = "Segoe UI"
F          = (FONT_NAME, 10)
F_BOLD     = (FONT_NAME, 10, "bold")
F_TITLE    = (FONT_NAME, 18, "bold")
F_SUB      = (FONT_NAME, 11)
F_SMALL    = (FONT_NAME, 8)
F_STAT     = (FONT_NAME, 11, "bold")


def apply_styles():
    s = ttk.Style()
    s.theme_use("clam")
    s.configure("T.Treeview", background=COLORS["surface"], foreground=COLORS["text"],
                fieldbackground=COLORS["surface"], font=F, rowheight=30,
                borderwidth=0, relief="flat")
    s.configure("T.Treeview.Heading", background=COLORS["hdr_bg"],
                foreground=COLORS["cyan"], font=F_BOLD, relief="flat",
                padding=(8, 6))
    s.map("T.Treeview.Heading", background=[("active", COLORS["surface2"])])
    s.map("T.Treeview", background=[("selected", COLORS["sel"])],
          foreground=[("selected", "#fff")])
    s.configure("A.TButton", background=COLORS["accent"], foreground="#fff",
                font=F_BOLD, padding=(24, 10), borderwidth=0, relief="flat")
    s.map("A.TButton", background=[("active", COLORS["accent_dk"])])
    s.configure("TScrollbar", background=COLORS["surface2"], troughcolor=COLORS["bg"],
                bordercolor=COLORS["bg"], arrowcolor=COLORS["dim"], borderwidth=0,
                relief="flat")
    s.configure("TCombobox", fieldbackground=COLORS["input_bg"],
                background=COLORS["surface2"], foreground=COLORS["text"],
                bordercolor=COLORS["input_bd"], arrowcolor=COLORS["dim"],
                relief="flat")
    s.map("TCombobox", fieldbackground=[("readonly", COLORS["input_bg"])],
          foreground=[("readonly", COLORS["text"])])
