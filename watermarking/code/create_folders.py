"""
create_folders.py
-----------------
Paste a numbered folder structure into the left pane.
The right pane shows a live preview of how the folders will nest.
Click "Create Folders" to build the structure in watermarking/output/.

Numbering determines depth:
  1          -> level 1
  1.1        -> level 2
  1.1.1      -> level 3
  1.1.1.1    -> level 4
  (and so on)

Folder names include the number prefix as-is.
Existing folders are left in place; new ones are created alongside them.
"""

import os
import re
import tkinter as tk
from tkinter import messagebox

# ── PATHS ─────────────────────────────────────────────────────────────────────
SCRIPT_DIR    = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR    = os.path.dirname(SCRIPT_DIR)
OUTPUT_FOLDER = os.path.join(PARENT_DIR, "output")
# ─────────────────────────────────────────────────────────────────────────────

# ── BRAND COLOURS ─────────────────────────────────────────────────────────────
NAVY      = "#0B2340"
GOLD      = "#C8A84B"
BG        = "#F4F4F2"
CARD      = "#FFFFFF"
LOG_BG    = "#0D1B2A"
LOG_FG    = "#A8C4E0"
TEXT_DARK = "#1A1A2E"
MUTED     = "#6B7280"
# ─────────────────────────────────────────────────────────────────────────────


def parse_structure(raw_text):
    """
    Parse pasted text into a list of (depth, folder_name) tuples.
    Depth is 0-indexed: top-level = 0.

    Accepts lines like:
      1 Juridisch
      1.1. Vennootschapsrecht
      1.1.1. Structuur
      1.1.1.1. Oprichtingsakte

    Lines that don't start with a recognisable numeric prefix are skipped.
    """
    pattern = re.compile(r'^(\d+(?:\.\d+)*\.?)\s+(.+)')
    results = []
    for raw_line in raw_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        m = pattern.match(line)
        if not m:
            continue
        numeric_part = m.group(1).rstrip(".")   # e.g. "1.1.2"
        label        = line                      # full line kept as folder name
        # Sanitise characters Windows won't allow in folder names
        label = re.sub(r'[<>:"/\\|?*]', '-', label)
        label = label.strip()
        depth = numeric_part.count(".")          # "1" -> 0, "1.1" -> 1, etc.
        results.append((depth, label))
    return results


def build_folder_paths(entries):
    """
    Convert (depth, name) pairs into full absolute paths under OUTPUT_FOLDER.
    Uses a stack to track the current path at each depth level.
    """
    stack = {}   # depth -> folder name at that depth
    paths = []
    for depth, name in entries:
        stack[depth] = name
        # Drop any deeper levels that are now stale
        stale = [k for k in stack if k > depth]
        for k in stale:
            del stack[k]
        parts = [stack[d] for d in sorted(stack)]
        full_path = os.path.join(OUTPUT_FOLDER, *parts)
        paths.append(full_path)
    return paths


def preview_text(entries):
    """Build a readable indented preview string."""
    lines = []
    for depth, name in entries:
        indent = "    " * depth
        lines.append(f"{indent}{name}")
    return "\n".join(lines)


def create_folders(paths, log_widget):
    """Create all folders. Logs each action to the log widget."""
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    created = 0
    skipped = 0
    for path in paths:
        rel = os.path.relpath(path, PARENT_DIR)
        if os.path.isdir(path):
            log_widget.insert(tk.END, f"  EXISTS   {rel}\n")
            skipped += 1
        else:
            os.makedirs(path, exist_ok=True)
            log_widget.insert(tk.END, f"  CREATED  {rel}\n")
            created += 1
        log_widget.see(tk.END)
    log_widget.insert(tk.END, f"\nDone. {created} created, {skipped} already existed.\n")
    log_widget.see(tk.END)


# ── GUI ───────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Folder Structure Creator")
        self.configure(bg=BG)
        self.resizable(True, True)
        self._build()
        # Minimum usable size
        self.minsize(900, 600)
        self.geometry("1100x680")

    def _build(self):
        # ── Header ──
        header = tk.Frame(self, bg=NAVY, height=52)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)
        tk.Label(
            header, text="Kumulus Partners",
            font=("Segoe UI", 15, "bold"),
            bg=NAVY, fg=CARD
        ).pack(side=tk.LEFT, padx=18, pady=10)
        tk.Label(
            header, text="Folder Structure Creator",
            font=("Segoe UI", 11),
            bg=NAVY, fg="#7BAFD4"
        ).pack(side=tk.LEFT, padx=4, pady=10)

        # ── Body ──
        body = tk.Frame(self, bg=BG)
        body.pack(fill=tk.BOTH, expand=True, padx=16, pady=14)

        # Left card: paste input
        left_card = tk.Frame(body, bg=CARD, bd=0, highlightthickness=1,
                             highlightbackground="#D1D5DB")
        left_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))

        tk.Label(left_card, text="Paste structure here",
                 font=("Segoe UI", 10, "bold"), bg=CARD, fg=TEXT_DARK
                 ).pack(anchor="w", padx=12, pady=(10, 2))
        tk.Label(left_card,
                 text="One item per line, no blank lines between entries.",
                 font=("Segoe UI", 9), bg=CARD, fg=MUTED
                 ).pack(anchor="w", padx=12, pady=(0, 6))

        self.input_text = tk.Text(
            left_card, font=("Consolas", 9), bg="#FAFAFA", fg=TEXT_DARK,
            relief=tk.FLAT, bd=0, wrap=tk.NONE,
            insertbackground=NAVY, selectbackground=NAVY, selectforeground=CARD
        )
        self.input_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        self.input_text.bind("<KeyRelease>", self._on_input_change)

        # Right card: preview
        right_card = tk.Frame(body, bg=CARD, bd=0, highlightthickness=1,
                              highlightbackground="#D1D5DB")
        right_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0))

        tk.Label(right_card, text="Folder preview",
                 font=("Segoe UI", 10, "bold"), bg=CARD, fg=TEXT_DARK
                 ).pack(anchor="w", padx=12, pady=(10, 2))
        tk.Label(right_card, text="How the folders will nest in output/",
                 font=("Segoe UI", 9), bg=CARD, fg=MUTED
                 ).pack(anchor="w", padx=12, pady=(0, 6))

        self.preview_text_widget = tk.Text(
            right_card, font=("Consolas", 9), bg="#FAFAFA", fg=TEXT_DARK,
            relief=tk.FLAT, bd=0, wrap=tk.NONE, state=tk.DISABLED,
            selectbackground=NAVY, selectforeground=CARD
        )
        self.preview_text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # ── Bottom: log + button ──
        bottom = tk.Frame(self, bg=BG)
        bottom.pack(fill=tk.X, padx=16, pady=(0, 14))

        btn_frame = tk.Frame(bottom, bg=BG)
        btn_frame.pack(side=tk.RIGHT, padx=(8, 0))

        tk.Button(
            btn_frame, text="Create Folders",
            font=("Segoe UI", 10, "bold"),
            bg=NAVY, fg=CARD, activebackground="#1a3a5c", activeforeground=CARD,
            relief=tk.FLAT, padx=18, pady=8, cursor="hand2",
            command=self._run
        ).pack()

        log_frame = tk.Frame(bottom, bg=LOG_BG, bd=0, highlightthickness=1,
                             highlightbackground="#1E3A5F")
        log_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.log = tk.Text(
            log_frame, font=("Consolas", 9), bg=LOG_BG, fg=LOG_FG,
            relief=tk.FLAT, bd=0, height=6, wrap=tk.WORD, state=tk.NORMAL
        )
        self.log.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

    def _on_input_change(self, _event=None):
        raw = self.input_text.get("1.0", tk.END)
        entries = parse_structure(raw)
        preview = preview_text(entries)
        self.preview_text_widget.config(state=tk.NORMAL)
        self.preview_text_widget.delete("1.0", tk.END)
        self.preview_text_widget.insert(tk.END, preview)
        self.preview_text_widget.config(state=tk.DISABLED)

    def _run(self):
        raw = self.input_text.get("1.0", tk.END)
        entries = parse_structure(raw)
        if not entries:
            messagebox.showwarning("No input", "Paste a numbered folder structure first.")
            return
        paths = build_folder_paths(entries)
        self.log.delete("1.0", tk.END)
        self.log.insert(tk.END, f"Output: {OUTPUT_FOLDER}\n\n")
        create_folders(paths, self.log)


if __name__ == "__main__":
    app = App()
    app.mainloop()
