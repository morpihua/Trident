import tkinter as tk

class Tooltip:
    def __init__(self, widget, text, background="#313335", foreground="#EAEAEA"):
        self.widget = widget
        self.text = text
        self.background = background
        self.foreground = foreground
        self.tooltip_window = None
        widget.bind("<Enter>", self.enter)
        widget.bind("<Leave>", self.leave)

    def enter(self, event=None):
        try:
            x, y, _, _ = self.widget.bbox("insert")
        except Exception:
            x, y = 0, 0
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tooltip_window, text=self.text, justify='left',
                         background=self.background, foreground=self.foreground, relief='solid', borderwidth=1,
                         font=("Courier New", 8, "normal"))
        label.pack(ipadx=2, ipady=2)

    def leave(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
        self.tooltip_window = None
