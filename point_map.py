"""
Tkinter window with an embedded matplotlib "map" of points relative to a datum.

- Marker SHAPE  = point type
- Marker COLOUR = point set
- Pan / zoom / save via the standard matplotlib toolbar
- Hover over a point for a tooltip (name, type, set, offsets, distance)
- Tick boxes on the right to show/hide sets and types
"""
import math
import tkinter as tk
from tkinter import ttk

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.lines import Line2D
from matplotlib.patches import Rectangle
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

# ---- Styling: one place to change how things look ---------------------------
TYPE_MARKERS = {"500KG": "^",
                "2000KG": "s",
                "1000KG": "o",
                "REF": "X"}
SET_COLOURS = {"On Stage": "#00ff00",
               "Downstage": "#f5910f",
               "Advance": "#aa00aa",
               "Upstage": "#00aaaa",
               "Stage Left": "#0000ff",
               "Stage Right": "#ff0000"}

# ---- Example data (absolute coordinates; plotted relative to the datum) -----
DATUM = (0,0)
POINTS = [
    {"name": "BM1", "x": 1012.0, "y": 2005.0, "type": "500KG", "set": "Onstage"},
    {"name": "BM2", "x": 985.5,  "y": 2018.2, "type": "500KG", "set": "Onstage"},
    {"name": "C1",  "x": 10000, "y": -2200, "type": "2000KG",   "set": "Downstage"},
    {"name": "C2",  "x": -10000,  "y": -2200, "type": "2000KG",   "set": "Downstage"},
    {"name": "D1",  "x": 5000, "y": 11500, "type": "1000KG",    "set": "Upstage"},
    {"name": "D2",  "x": -5000,  "y": 11500, "type": "1000KG",    "set": "Upstage"},
    {"name": "D3",  "x": 1016.8, "y": 2022.5, "type": "REF",    "set": "Stage Left"},
    {"name": "H1",  "x": 990.1,  "y": 2003.4, "type": "REF",    "set": "Stage Right"},
]


class PointMap(ttk.Frame):
    def __init__(self, master, points, 
                 datum=(0.0, 0.0), 
                 datum_label="Datum", 
                 stage=None):
        super().__init__(master)
        self.points = points
        self.datum = datum
        self.datum_label = datum_label
        self.stage = stage  # dict with keys: width, depth, x, y (absolute coords)
        self.hits = []          # (scatter artist, list of points) for hover lookup
        self.legends = []
        self._first_draw = True

        # --- Plot area (left) ---
        plot_frame = ttk.Frame(self)
        plot_frame.pack(side="left", fill="both", expand=True)

        self.fig = Figure(figsize=(7, 6), dpi=100, layout="constrained")
        self.ax = self.fig.add_subplot()
        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        toolbar = NavigationToolbar2Tk(self.canvas, plot_frame, pack_toolbar=False)
        toolbar.update()
        toolbar.pack(side="bottom", fill="x")
        self.canvas.get_tk_widget().pack(side="top", fill="both", expand=True)
        self.canvas.mpl_connect("motion_notify_event", self._on_hover)

        # --- Filter panel (right) ---
        panel = ttk.Frame(self, padding=8)
        panel.pack(side="right", fill="y")
        self.set_vars = self._checkboxes(panel, "Sets", SET_COLOURS)
        self.type_vars = self._checkboxes(panel, "Types", TYPE_MARKERS)

        self.redraw()

    def _checkboxes(self, parent, title, keys):
        box = ttk.LabelFrame(parent, text=title, padding=6)
        box.pack(fill="x", pady=(0, 8))
        vars_ = {}
        for key in keys:
            v = tk.BooleanVar(value=True)
            ttk.Checkbutton(box, text=key, variable=v, command=self.redraw).pack(anchor="w")
            vars_[key] = v
        return vars_

    def _rel(self, p):
        return p["x"] - self.datum[0], p["y"] - self.datum[1]

    def _draw_static(self):
        """Things drawn once: datum, axes styling, tooltip."""
        ax = self.ax
        ax.axhline(0, color="0.6", lw=0.8, zorder=0)
        ax.axvline(0, color="0.6", lw=0.8, zorder=0)
        ax.plot(0, 0, marker="P", ms=14, color="black", zorder=3)
        ax.annotate(self.datum_label, (0, 0), xytext=(8, 8), textcoords="offset points")
        ax.set_aspect("equal", adjustable="box")   # 1 m east looks like 1 m north
        ax.grid(True, ls=":", alpha=0.6)
        ax.set_xlabel("Onstage/Offstage from datum (m)")
        ax.set_ylabel("Upstage/Downstage from datum (m)")
        if self.stage:
            w = self.stage["width"]/1000
            d = self.stage["depth"]/1000
            cx = self.stage["x-shift"]/1000 - self.datum[0]
            cy = self.stage["y-shift"]/1000 - self.datum[1]
            rect = Rectangle((cx - w / 2, cy), w, d,
                              linewidth=1.2, edgecolor="0.4", facecolor="0.93", zorder=0)
            ax.add_patch(rect)
        self.tip = ax.annotate("", (0, 0), xytext=(12, 12), textcoords="offset points",
                               bbox=dict(boxstyle="round", fc="lightyellow", ec="0.4"),
                               zorder=10, visible=False)

    def _fit_all(self, pad=0.15):
        """Square view centred on the datum that contains every point."""
        r = max([max(abs(a), abs(b)) for a, b in map(self._rel, self.points)] or [1.0])
        r *= 1 + pad
        self.ax.set_xlim(-r, r)
        self.ax.set_ylim(-r, r)
        
    def redraw_with_points(self, points, stage=None):
        print('got here (rwp)')
        print(points)
        self._first_draw = True
        self.stage = stage
        self.points = points
        self.redraw()

    def redraw(self):
        """(Re)draw the point layers and legends; keeps the current zoom/pan."""
        ax = self.ax
        if self._first_draw:
            self._draw_static()
        for sc, _ in self.hits:
            sc.remove()
        for leg in self.legends:
            leg.remove()
        self.hits, self.legends = [], []

        # One scatter per (type, set) combination
        for t, marker in TYPE_MARKERS.items():
            if not self.type_vars[t].get():
                continue
            for s, colour in SET_COLOURS.items():
                if not self.set_vars[s].get():
                    continue
                group = [p for p in self.points if p["type"] == t and p["set"] == s]
                if not group:
                    continue
                xs, ys = zip(*(self._rel(p) for p in group))
                sc = ax.scatter(xs, ys, marker=marker, color=colour, s=70,
                                edgecolors="black", linewidths=0.6, zorder=4)
                self.hits.append((sc, group))

        # Two separate legends: colour = set, shape = type
        set_handles = [Line2D([], [], ls="", marker="o", color=c, mec="black", ms=9, label=s)
                       for s, c in SET_COLOURS.items() if self.set_vars[s].get()]
        type_handles = [Line2D([], [], ls="", marker=m, color="0.5", mec="black", ms=9, label=t)
                        for t, m in TYPE_MARKERS.items() if self.type_vars[t].get()]
        if set_handles:
            leg = ax.legend(handles=set_handles,
                            title="Markout",
                            loc="upper right")
            ax.add_artist(leg)
            self.legends.append(leg)
        if type_handles:
            self.legends.append(ax.legend(handles=type_handles,
                                          title="Type",
                                          loc="lower right"))

        if self._first_draw:
            self._fit_all()  # fixed limits, so filtering doesn't make the view jump
            self._first_draw = False
        self.canvas.draw_idle()

    def _on_hover(self, event):
        if event.inaxes is self.ax:
            for sc, group in self.hits:
                hit, info = sc.contains(event)
                if hit:
                    p = group[info["ind"][0]]
                    dx, dy = self._rel(p)
                    self.tip.xy = (dx, dy)
                    self.tip.set_text(
                        f"{p['name']}  ({p['type']}, {p['set']})\n"
                        f"ΔE {dx:+.2f}  ΔN {dy:+.2f}\n")
                        #f"Dist {math.hypot(dx, dy):.2f}  "
                        #f"Brg {math.degrees(math.atan2(dx, dy)) % 360:.1f}°")
                    self.tip.set_visible(True)
                    self.canvas.draw_idle()
                    return
        if self.tip.get_visible():
            self.tip.set_visible(False)
            self.canvas.draw_idle()


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Point map")
    PointMap(root,
             POINTS,
             datum=DATUM,
             datum_label="Datum",
             stage={"width": 18000,
                    "depth": 12000,
                    'x-shift': 0,
                    'y-shift': 0}
             ).pack(fill="both", expand=True)
    root.mainloop()