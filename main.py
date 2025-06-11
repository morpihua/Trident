import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import copy
import re
from typing import List, Dict, Any, Optional
import csv
import json
import zipfile
import random
import math
import uuid
from datetime import datetime, timezone
import xml.etree.ElementTree as ET

# Сторонні бібліотеки (переконайтесь, що вони встановлені: pip install openpyxl xlsxwriter)
try:
    import openpyxl
    import xlsxwriter
except ImportError:
    messagebox.showerror("Відсутні бібліотеки", "Будь ласка, встановіть 'openpyxl' та 'xlsxwriter' для роботи з файлами Excel.\n\npip install openpyxl xlsxwriter")
    sys.exit()

from apq import ApqFile
from colors import COLORS, COLORS_EN_UA
from utils import xml_escape, calculate_distance
from tooltip import Tooltip

class Main:
    MAX_FILES = 100
    CSV_CHUNK_SIZE = 2000  # Ліміт рядків на один CSV файл

    def __init__(self):
        self.program_version = "8.2.0_new_palette"
        self.empty = "Не вибрано"
        self.file_ext, self.file_name = None, None

        # --- ОНОВЛЕНІ СПИСКИ ТА ПАЛІТРИ КОЛЬОРІВ ---
        self.list_of_formats = [".geojson", ".kml", ".kmz", ".kme", ".gpx", ".xlsx", ".csv", ".csv(макет)"]
        self.supported_read_formats = [".kml", ".kmz", ".kme", ".gpx", ".xlsx", ".csv", ".scene", ".wpt", ".set",
                                       ".rte", ".are", ".trk", ".ldk"]
        self.numerations = ["За найближчими сусідями", "За змійкою", "За відстаню від кута", "За відстаню від границі",
                            "За випадковістю"]
        self.translations = ["Не повертати", "На 90 градусів", "На 180 градусів", "На 270 градусів"]

        # Новий словник з кольорами та їх HEX-кодами
        self.colors = {
            "Red": "#f44336", "Pink": "#e91e63", "Purple": "#9c27b0", "DeepPurple": "#673ab7",
            "Indigo": "#3f51b5", "Blue": "#2196f3", "Cyan": "#00bcd4", "Teal": "#009688",
            "Green": "#4caf50", "LightGreen": "#8bc34a", "Lime": "#cddc39", "Yellow": "#ffeb3b",
            "Amber": "#ffc107", "Orange": "#ff9800", "DeepOrange": "#ff5722", "Brown": "#795548",
            "BlueGrey": "#607d8b", "Black": "#010101", "White": "#ffffff"
        }
        
        # ICON_ID_COLOR_MAP та extended_color_names_map для розпізнавання кольору
        # Вони визначені тут, як члени класу, щоб бути доступними в `convert_color`
        self.ICON_ID_COLOR_MAP = {
            0: (255,255,255), # white
            1: (102,51,153), # violet
            2: (255,0,0),    # red
            3: (0,128,0),    # green
            4: (0,0,255),    # blue
            5: (255,255,0),  # yellow
            6: (255,165,0),  # orange
            7: (128,128,128),# gray
            8: (0,0,0),      # black
            9: (255,192,203),# pink
            10: (0,255,255), # cyan
            # ДОДАЙТЕ СЮДИ ІНШІ icon_id та їх RGB значення, якщо вони є у ваших файлах WPT
        }

        self.extended_color_names_map = {
            "червоний": "Red", "фіолетовий": "Purple", "білий": "White", "чорний": "Black",
            "жовтий": "Yellow", "помаранчевий": "Orange", "сірий": "BlueGrey", "рожевий": "Pink",
            "блакитний": "Cyan", "зелений": "Green", "синій": "Blue",
            "салатовий": "LightGreen", "лаймовий": "Lime", "бурштиновий": "Amber",
            "насичено-помаранчевий": "DeepOrange", "коричневий": "Brown",
            "темно-фіолетовий": "DeepPurple", "індиго": "Indigo", "бірюзовий": "Teal",
            "красный": "Red", "фиолетовый": "Purple", "розовый": "Pink",
            "темно-фиолетовый": "DeepPurple", "индиго": "Indigo", "синий": "Blue",
            "бирюзовый": "Teal", "зеленый": "Green", "салатовый": "LightGreen",
            "лаймовый": "Lime", "желтый": "Yellow", "янтарный": "Amber",
            "оранжевый": "Orange", "насыщенно-оранжевый": "DeepOrange",
            "коричневый": "Brown", "сине-серый": "BlueGrey", "черный": "Black",
            "белый": "White", "голубой": "Cyan",
            # ДОДАЙТЕ СЮДИ ІНШІ АЛІАСИ КОЛЬОРІВ, ЯКЩО ПОТРІБНО
        }

        # Нові опції для випадаючого меню
        self.color_options = ["Без змін"] + list(self.colors.keys())

        # Нові переклади для кольорів
        self.colors_en_ua = {
            "Red": "Червоний", "Pink": "Рожевий", "Purple": "Фіолетовий", "DeepPurple": "Темно-фіолетовий",
            "Indigo": "Індиго", "Blue": "Синій", "Cyan": "Блакитний", "Teal": "Бірюзовий",
            "Green": "Зелений", "LightGreen": "Салатовий", "Lime": "Лаймовий", "Yellow": "Жовтий",
            "Amber": "Бурштиновий", "Orange": "Помаранчевий", "DeepOrange": "Насичено-помаранчевий",
            "Brown": "Коричневий",
            "BlueGrey": "Синьо-сірий", "Black": "Чорний", "White": "Білий"
        }
        # --- КІНЕЦЬ ОНОВЛЕННЯ КОЛЬОРІВ ---

        self.file_list = []
        self.list_is_visible = False
        self.output_directory_path = self.empty

        self.font, self.C_BACKGROUND, self.C_SIDEBAR, self.C_BUTTON, self.C_BUTTON_HOVER, self.C_TEXT = ("Courier New",
                                                                                                         11,
                                                                                                         "bold"), "#2B2B2B", "#3C3C3C", "#556B2F", "#6B8E23", "#F5F5F5"
        self.C_ACCENT_SUCCESS, self.C_ACCENT_DONE, self.C_STATUS_DEFAULT, self.C_ACCENT_ERROR = "#6B8E23", "#FFBF00", "#4F4F4F", "#8B0000"

        self.main_window = tk.Tk()

        self.names_agree = tk.BooleanVar(value=False)
        self.exceptions_agree = tk.BooleanVar(value=False)
        self.chosen_numeration = tk.StringVar(value="За найближчими сусідями")
        self.chosen_translation = tk.StringVar(value="Не повертати")

        self.main_window.title(f"Nexus v{self.program_version}")
        self.main_window.configure(background=self.C_BACKGROUND)
        self.main_window.minsize(450, 120)
        self.main_window.geometry("450x120")
        try:
            base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            icon_path = os.path.join(base_path, 'nexus.ico')
            if os.path.exists(icon_path): self.main_window.iconbitmap(icon_path)
        except Exception as e:
            print(f"Попередження: не вдалося завантажити іконку: {e}")
        self.main_window.protocol("WM_DELETE_WINDOW", self.exit)
        self.main_window.resizable(True, True)

        self._configure_styles()
        self._build_main_ui()

        self.input_file_path = None
        self.output_directory_path = self.empty

    def _configure_styles(self):
        style = ttk.Style(self.main_window)
        style.theme_use('clam')
        style.configure("TFrame", background=self.C_BACKGROUND)
        style.configure("Side.TFrame", background=self.C_SIDEBAR)
        style.configure("List.TFrame", background=self.C_SIDEBAR)
        style.configure('Icon.TButton', padding=5, borderwidth=0, relief='flat',
                        background=self.C_BUTTON, foreground=self.C_TEXT, font=self.font)
        style.map('Icon.TButton', background=[('active', self.C_BUTTON_HOVER)],
                  foreground=[('active', self.C_TEXT)])
        style.configure('Remove.TButton', background=self.C_SIDEBAR, foreground="#FF6347",
                        font=("Courier New", 10, "bold"), relief='flat', borderwidth=0)
        style.map('Remove.TButton', background=[('active', "#4a4a4a")])

        style.configure("Toplevel", background=self.C_BACKGROUND)
        style.configure("TCheckbutton", background=self.C_BACKGROUND, foreground=self.C_TEXT, font=self.font,
                        indicatorcolor=self.C_TEXT, selectcolor=self.C_BUTTON_HOVER)
        style.map("TCheckbutton", background=[('active', self.C_BACKGROUND)])

        style.configure("TLabel", background=self.C_BACKGROUND, foreground=self.C_TEXT, font=self.font)
        style.configure("List.TLabel", background=self.C_SIDEBAR, foreground=self.C_TEXT, font=("Courier New", 9))

        style.configure("Dark.TEntry", fieldbackground="#4F4F4F", foreground=self.C_TEXT, insertcolor=self.C_TEXT,
                        bordercolor=self.C_SIDEBAR, font=("Courier New", 9))

        style.configure("TMenubutton", background="#4F4F4F", foreground=self.C_TEXT, font=("Courier New", 9),
                        borderwidth=1, relief='raised', arrowcolor=self.C_TEXT)
        style.map("TMenubutton", background=[('active', "#646464")])

    def load_icons(self):
        self.icons = {}

    def run(self):
        self.main_window.mainloop()

    def exit(self):
        if messagebox.askokcancel("Вихід", "Ви впевнені, що хочете вийти?"):
            self.main_window.destroy()

    def _build_main_ui(self):
        self.main_window.rowconfigure(0, weight=0)
        self.main_window.rowconfigure(1, weight=1)
        self.main_window.columnconfigure(0, weight=1)

        top_container = ttk.Frame(self.main_window)
        top_container.grid(row=0, column=0, sticky="ew", pady=(5, 0))
        top_container.columnconfigure(0, weight=0)
        top_container.columnconfigure(1, weight=1)
        top_container.columnconfigure(2, weight=0)

        left_sidebar = ttk.Frame(top_container, width=50, style="Side.TFrame")
        left_sidebar.grid(row=0, column=0, sticky="ns", padx=(5, 2))

        btn_lightbulb = ttk.Button(left_sidebar, text="i", style='Icon.TButton', command=self.show_info, width=2)
        btn_lightbulb.pack(pady=(5, 5), padx=5, fill='x')
        Tooltip(btn_lightbulb, "Про програму", background=self.C_SIDEBAR, foreground=self.C_TEXT)

        btn_settings = ttk.Button(left_sidebar, text="S", style='Icon.TButton', command=self.open_numeration_settings,
                                  width=2)
        btn_settings.pack(pady=5, padx=5, fill='x')
        Tooltip(btn_settings, "Налаштування нумерації", background=self.C_SIDEBAR, foreground=self.C_TEXT)

        center_frame = ttk.Frame(top_container)
        center_frame.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        self.status_label = ttk.Label(center_frame, anchor="center", font=("Courier New", 14, "bold"),
                                      foreground=self.C_TEXT, relief='flat', padding=(0, 10))
        self.status_label.pack(fill="both", expand=True)
        self._update_status("ДОДАЙТЕ ФАЙЛИ", self.C_STATUS_DEFAULT)

        right_sidebar = ttk.Frame(top_container, width=50, style="Side.TFrame")
        right_sidebar.grid(row=0, column=2, sticky="ns", padx=(2, 5))

        self.btn_open_file = ttk.Button(right_sidebar, text="F", style='Icon.TButton', command=self.add_files_to_list,
                                        width=2)
        self.btn_open_file.pack(pady=(5, 5), padx=5, fill='x')
        Tooltip(self.btn_open_file, "Додати файли", background=self.C_SIDEBAR, foreground=self.C_TEXT)

        self.play_button = ttk.Button(right_sidebar, text="▶", style='Icon.TButton',
                                      command=self.start_convertion, state="disabled", width=2)
        self.play_button.pack(pady=5, padx=5, fill='x')
        Tooltip(self.play_button, "Конвертувати все", background=self.C_SIDEBAR, foreground=self.C_TEXT)

        self.list_container = ttk.Frame(self.main_window, style="List.TFrame")
        self.canvas = tk.Canvas(self.list_container, bg=self.C_SIDEBAR, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.list_container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, style="List.TFrame")

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True, padx=(1, 0), pady=(1, 1))
        scrollbar.pack(side="right", fill="y", padx=(0, 1), pady=(1, 1))
        self.canvas.bind("<Configure>", self._on_canvas_configure)

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _redraw_file_list(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        for i, file_data in enumerate(self.file_list):
            item_frame = ttk.Frame(self.scrollable_frame, style="List.TFrame", padding=(5, 2))
            item_frame.pack(fill='x', expand=True)

            label_text = f"{i + 1}. {file_data['base_name']}"
            if len(label_text) > 40: label_text = label_text[:37] + "..."

            label = ttk.Label(item_frame, text=label_text, style="List.TLabel", anchor='w')
            label.pack(side='left', fill='x', expand=True, padx=(0, 5))

            format_mb = ttk.Menubutton(item_frame, text=file_data['format_var'].get(), style="TMenubutton", width=7)
            format_menu_tk = tk.Menu(format_mb, tearoff=0, bg=self.C_SIDEBAR, fg=self.C_TEXT,
                                     activebackground=self.C_BUTTON_HOVER)
            for fmt_option in self.list_of_formats:
                format_menu_tk.add_radiobutton(label=fmt_option, variable=file_data['format_var'], value=fmt_option,
                                               command=lambda var=file_data['format_var'], button=format_mb,
                                                              val=fmt_option: self._update_menubutton_text(var, button,
                                                                                                           val))
            format_mb['menu'] = format_menu_tk
            format_mb.pack(side='left', padx=3)

            color_mb = ttk.Menubutton(item_frame, text=self.colors_en_ua.get(file_data['color_var'].get(),
                                                                             file_data['color_var'].get()),
                                      style="TMenubutton", width=12)
            color_menu_tk = tk.Menu(color_mb, tearoff=0, bg=self.C_SIDEBAR, fg=self.C_TEXT,
                                    activebackground=self.C_BUTTON_HOVER)
            for color_option in self.color_options:
                disp_name = self.colors_en_ua.get(color_option, color_option)
                color_menu_tk.add_radiobutton(label=disp_name, variable=file_data['color_var'], value=color_option,
                                              command=lambda var=file_data['color_var'], button=color_mb,
                                                             val_en=color_option: self._update_menubutton_text(var,
                                                                                                               button,
                                                                                                               val_en))
            color_mb['menu'] = color_menu_tk
            color_mb.pack(side='left', padx=3)

            remove_btn = ttk.Button(item_frame, text="X", style='Remove.TButton', width=2,
                                    command=lambda fd=file_data: self._remove_file(fd))
            remove_btn.pack(side='left', padx=(3, 0))

        if not self.file_list:
            self._update_status("ДОДАЙТЕ ФАЙЛИ", self.C_STATUS_DEFAULT)
            self.play_button.config(state="disabled")
            if self.list_is_visible:
                self.list_container.grid_forget()
                self.list_is_visible = False
                self.main_window.geometry("450x120")
        else:
            status_text = f"ГОТОВО: {len(self.file_list)} ФАЙЛ(ІВ)"
            if len(self.file_list) == 1:
                status_text = f"ГОТОВО: {len(self.file_list)} ФАЙЛ"
            elif 2 <= len(self.file_list) <= 4:
                status_text = f"ГОТОВО: {len(self.file_list)} ФАЙЛИ"

            self._update_status(status_text, self.C_ACCENT_SUCCESS)
            self.play_button.config(state="normal")

        self.scrollable_frame.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        self.canvas.itemconfig(self.canvas_window, width=self.canvas.winfo_width())

    def _update_menubutton_text(self, var, menubutton, value_english):
        var.set(value_english)
        display_text = self.colors_en_ua.get(value_english, value_english)
        menubutton.config(text=display_text)

    def _remove_file(self, file_to_remove):
        self.file_list.remove(file_to_remove)
        if not self.file_list and self.list_is_visible:
            self.list_container.grid_forget()
            self.list_is_visible = False
            self.main_window.geometry("450x120")
        self._redraw_file_list()

    def add_files_to_list(self):
        file_types = [
            ("Підтримувані файли", " ".join(f"*{ext}" for ext in self.supported_read_formats)),
            ("AlpineQuest файли", ".wpt .set .rte .are .trk .ldk"),
            ("KML/KMZ/KME", ".kml .kmz .kme"),
            ("GPS Exchange", ".gpx"),
            ("Excel", ".xlsx"),
            ("CSV", ".csv"),
            ("SCENE JSON", ".scene"),
            ("Всі файли", "*.*")
        ]
        paths = filedialog.askopenfilenames(filetypes=file_types, title="Виберіть файли для конвертації")

        new_files_added = False
        if paths:
            for path in paths:
                if any(f['full_path'] == path for f in self.file_list):
                    self._update_status(f"Файл вже у списку: {os.path.basename(path)}", warning=True)
                    continue

                if len(self.file_list) >= self.MAX_FILES:
                    messagebox.showwarning("Ліміт файлів",
                                           f"Максимальна кількість файлів у списку ({self.MAX_FILES}) досягнута.")
                    break

                base_name = os.path.basename(path)
                file_ext = os.path.splitext(base_name)[1].lower()

                if file_ext not in self.supported_read_formats:
                    messagebox.showwarning("Формат не підтримується",
                                           f"Програма не може імпортувати дані з файлів формату '{file_ext}'.")
                    continue

                default_export_format = ".kml"
                if file_ext in [".wpt", ".set", ".rte", ".are", ".trk", ".ldk", ".kmz", ".kme"]:
                    default_export_format = ".kml"
                elif file_ext == ".gpx":
                    default_export_format = ".gpx"
                elif file_ext == ".xlsx":
                    default_export_format = ".xlsx"
                elif file_ext == ".csv":
                    default_export_format = ".csv"
                elif file_ext == ".scene":
                    default_export_format = ".geojson"

                file_data = {
                    "full_path": path, "base_name": base_name,
                    "format_var": tk.StringVar(value=default_export_format),
                    "color_var": tk.StringVar(value=self.color_options[0]),
                }
                self.file_list.append(file_data)
                new_files_added = True

            if new_files_added and not self.list_is_visible:
                self.list_container.grid(row=1, column=0, columnspan=3, sticky="nsew", padx=5, pady=(2, 5))
                self.main_window.rowconfigure(1, weight=3)
                self.list_is_visible = True
                self.main_window.geometry("650x400")

            if new_files_added:
                self._redraw_file_list()
            elif not self.file_list:
                self._update_status("ФАЙЛИ НЕ ДОДАНО", self.C_STATUS_DEFAULT)

    def _update_status(self, text, color=None, error=False, warning=False):
        if error:
            final_color = self.C_ACCENT_ERROR
        elif warning:
            final_color = self.C_ACCENT_DONE
        elif color:
            final_color = color
        else:
            final_color = self.C_TEXT

        if not error and not warning and not color and self.status_label.cget("background") == self.C_STATUS_DEFAULT:
            current_bg = self.C_STATUS_DEFAULT
        elif not error and not warning and not color:
            current_bg = self.C_STATUS_DEFAULT if self.status_label.cget(
                "background") == self.C_STATUS_DEFAULT else self.C_BACKGROUND
        else:
            current_bg = final_color

        self.status_label.config(text=text.upper(), background=current_bg, foreground=self.C_TEXT)
        if self.main_window.winfo_exists():
            self.main_window.update_idletasks()

    def _get_chunked_save_path(self, base_save_path, chunk_index):
        """Генерує шлях для збереження для певного чанка."""
        if chunk_index == 0:
            return base_save_path

        directory, filename = os.path.split(base_save_path)
        name_part, ext_part = os.path.splitext(filename)

        name_part = re.sub(r'\s*\(\d+\)$', '', name_part).strip()

        new_filename = f"{name_part}({chunk_index}){ext_part}"
        return os.path.join(directory, new_filename)

    def show_info(self):
        messagebox.showinfo("Про програму",
                            f"Nexus v{self.program_version}\nПрограма для пакетної конвертації та обробки геоданих.\n\nПідтримувані формати для читання:\n{', '.join(self.supported_read_formats)}\n\nПідтримувані формати для запису:\n{', '.join(fmt for fmt in self.list_of_formats if fmt not in ['.nvg', '.pgd'])}")

    def open_numeration_settings(self):
        settings_win = tk.Toplevel(self.main_window)
        settings_win.title("Налаштування нумерації")
        settings_win.configure(background=self.C_BACKGROUND)
        settings_win.transient(self.main_window)
        settings_win.grab_set()
        settings_win.resizable(False, False)

        main_frame = ttk.Frame(settings_win, padding=15)
        main_frame.pack(fill="both", expand=True)

        ttk.Checkbutton(main_frame, text="Увімкнути нумерацію точок", variable=self.names_agree).pack(anchor="w",
                                                                                                      pady=(0, 10))

        numeration_frame = ttk.LabelFrame(main_frame, text="Параметри нумерації", style="TFrame", padding=10)
        numeration_frame.pack(fill="x", expand=True, pady=5)

        ttk.Label(numeration_frame, text="Спосіб нумерації:").pack(anchor="w")
        numeration_combo = ttk.Combobox(numeration_frame, textvariable=self.chosen_numeration,
                                        values=self.numerations, state="readonly", font=("Courier New", 9))
        numeration_combo.pack(fill="x", pady=(2, 8))
        numeration_combo.set(self.chosen_numeration.get())

        ttk.Label(numeration_frame, text="Поворот осі сортування:").pack(anchor="w")
        translation_combo = ttk.Combobox(numeration_frame, textvariable=self.chosen_translation,
                                         values=self.translations, state="readonly", font=("Courier New", 9))
        translation_combo.pack(fill="x", pady=(2, 8))
        translation_combo.set(self.chosen_translation.get())

        ttk.Checkbutton(numeration_frame, text="Виключити номери (30-40, 500-510)",
                        variable=self.exceptions_agree).pack(anchor="w", pady=5)

        ttk.Button(main_frame, text="Закрити", command=settings_win.destroy, style="Icon.TButton").pack(pady=(15, 0))

        settings_win.update_idletasks()
        main_x, main_y = self.main_window.winfo_x(), self.main_window.winfo_y()
        main_w, main_h = self.main_window.winfo_width(), self.main_window.winfo_height()
        win_w, win_h = settings_win.winfo_width(), settings_win.winfo_height()
        x = main_x + (main_w // 2) - (win_w // 2)
        y = main_y + (main_h // 2) - (win_h // 2)
        settings_win.geometry(f"+{x}+{y}")
        settings_win.focus_set()

    def _process_data(self, file_content, color_override_english_name):
        if not file_content:  # <-- ЗАХИСТ ВІД None
            print("DEBUG: file_content is None or empty!", file_content)
            return None
        content = copy.deepcopy(file_content)

        if color_override_english_name != self.color_options[0]:
            for item in content:
                item["color"] = color_override_english_name
                if 'milgeo:meta:color' in item:
                    item['milgeo:meta:color'] = color_override_english_name

        if self.names_agree.get():
            points_to_numerate = [item for item in content if item.get('geometry_type', '').lower() == 'point']
            other_items = [item for item in content if item.get('geometry_type', '').lower() != 'point']

            if points_to_numerate:
                numerated_points = self._apply_selected_numeration(points_to_numerate)
                content = numerated_points + other_items
        return content

    def start_convertion(self):
        if not self.file_list:
            messagebox.showwarning("Увага", "Список файлів для конвертації порожній.")
            return

        readers = {
            ".kml": self.read_kml, ".kme": self.read_kml, ".kmz": self.read_kmz,
            ".gpx": self.read_gpx, ".xlsx": self.read_xlsx, ".csv": self.read_csv,
            ".scene": self.read_scene, ".geojson": self.read_geojson,
            ".wpt": self.read_wpt, ".set": self.read_set, ".rte": self.read_rte,
            ".are": self.read_are, ".trk": self.read_trk, ".ldk": self.read_ldk,
        }
        writers = {
            ".kml": self.create_kml, ".kme": self.create_kml, ".kmz": self.create_kmz,
            ".gpx": self.create_gpx, ".xlsx": self.create_xlsx, ".csv": self.create_csv,
            ".geojson": self.create_geojson, ".scene": self.create_scene,
        }

        total_files = len(self.file_list)
        conversion_successful_count = 0

        for i, file_data in enumerate(self.file_list):
            current_file_basename = file_data['base_name']
            self._update_status(f"ФАЙЛ {i + 1}/{total_files}: {current_file_basename}", self.C_BUTTON_HOVER)
            self.main_window.update_idletasks()

            try:
                input_path = file_data['full_path']
                self.file_name, self.file_ext = os.path.splitext(os.path.basename(input_path))
                self.file_ext = self.file_ext.lower()

                reader_func = readers.get(self.file_ext)
                if not reader_func:
                    self._update_status(f"Непідтримуваний формат для читання: {self.file_ext}", error=True)
                    continue

                file_content = reader_func(input_path)

                if file_content is None or not file_content:
                    messagebox.showwarning("Увага",
                                           f"У файлі {current_file_basename} не знайдено даних або сталася помилка читання. Файл пропущено.")
                    self._update_status(f"ПОМИЛКА ЧИТАННЯ: {current_file_basename}", warning=True)
                    continue

                processed_content = self._process_data(file_content, file_data['color_var'].get())
                if processed_content is None:
                    self._update_status(f"Помилка обробки даних: {current_file_basename}", error=True)
                    continue

                output_format = file_data['format_var'].get().lower()
                writer_func = writers.get(output_format)

                if not writer_func:
                    messagebox.showerror("Формат не підтримується",
                                         f"Конвертація у формат '{output_format}' не підтримується.")
                    continue

                clean_base_name = re.sub(r'\s*\(\d+\)$', '', self.file_name)
                suggested_name = f"new_{clean_base_name}{output_format}"

                if self.output_directory_path == self.empty or not os.path.isdir(self.output_directory_path):
                    self.output_directory_path = os.path.dirname(input_path)

                save_path = filedialog.asksaveasfilename(
                    initialdir=self.output_directory_path,
                    initialfile=suggested_name,
                    defaultextension=output_format,
                    filetypes=[(f"{output_format.upper()} Files", f"*{output_format}"), ("All Files", "*.*")],
                    title=f"Зберегти конвертований файл для {current_file_basename}"
                )
                if not save_path:
                    self._update_status(f"СКАСОВАНО: {current_file_basename}", warning=True)
                    continue

                self.output_directory_path = os.path.dirname(save_path)

                success = writer_func(processed_content, save_path)
                if success:
                    # Specific writers now handle their own success messages for chunking
                    if output_format != '.csv':
                        self._update_status(f"ЗБЕРЕЖЕНО: {os.path.basename(save_path)}", self.C_ACCENT_SUCCESS)
                    conversion_successful_count += 1
                else:
                    pass

            except NotImplementedError as e:
                messagebox.showerror("Не реалізовано", str(e))
                self._update_status(f"ПОМИЛКА ФОРМАТУ: {current_file_basename}", error=True)
            except ValueError as e:
                messagebox.showerror("Помилка даних", f"Проблема з даними у файлі {current_file_basename}: {e}")
                self._update_status(f"ПОМИЛКА ДАНИХ: {current_file_basename}", error=True)
            except Exception as e:
                messagebox.showerror("Критична помилка",
                                     f"Не вдалося конвертувати файл {current_file_basename}:\n\n{type(e).__name__}: {e}\n\nПеревірте консоль для деталей.")
                self._update_status(f"КРИТИЧНА ПОМИЛКА: {current_file_basename}", self.C_ACCENT_ERROR)
                import traceback
                traceback.print_exc()
                continue

        if conversion_successful_count == total_files and total_files > 0:
            final_message = f"УСПІШНО ЗАВЕРШЕНО: {conversion_successful_count}/{total_files} ФАЙЛ(ІВ)"
            final_color = self.C_ACCENT_DONE
        elif conversion_successful_count > 0:
            final_message = f"ЗАВЕРШЕНО З ПОМИЛКАМИ: {conversion_successful_count}/{total_files} УСПІШНО"
            final_color = self.C_ACCENT_DONE
        elif total_files > 0:
            final_message = f"ПОМИЛКА: 0/{total_files} ФАЙЛІВ КОНВЕРТОВАНО"
            final_color = self.C_ACCENT_ERROR
        else:
            final_message = "СПИСОК ПОРОЖНІЙ"
            final_color = self.C_STATUS_DEFAULT

        self._update_status(final_message, final_color)
        if total_files > 0:
            messagebox.showinfo("Завершено",
                                f"Пакетна конвертація завершена.\nУспішно: {conversion_successful_count} з {total_files}.")

    def _normalize_apq_data(self, apq_parsed_data, file_path_for_log=""):
        normalized_content = []
        if not apq_parsed_data or not isinstance(apq_parsed_data, dict) or not apq_parsed_data.get('parse_successful'):
            self._update_status(f"APQ парсер не повернув успішних даних для {file_path_for_log}", warning=True)
            return normalized_content

        apq_type = apq_parsed_data.get('type')
        global_meta = apq_parsed_data.get('meta', {})
        file_basename = os.path.basename(file_path_for_log)

        def _create_point_dict(loc_data, item_meta_data, default_name_prefix="Точка", item_idx=0,
                               apq_source_file_type_for_item=None,
                               source_file_global_meta_for_item=None,
                               apq_icon_id: Optional[int] = None, # ДОДАНО НОВІ ПАРАМЕТРИ
                               apq_color_int: Optional[int] = None): # ДОДАНО НОВІ ПАРАМЕТРИ
            if not loc_data or loc_data.get('lon') is None or loc_data.get('lat') is None:
                self._update_status(f"Увага: Пропущено точку (відсутні координати) у {file_basename}", warning=True)
                return None
            point_lon, point_lat = loc_data['lon'], loc_data['lat']

            effective_meta = source_file_global_meta_for_item.copy() if source_file_global_meta_for_item else {}
            effective_meta.update(item_meta_data)

            final_name = effective_meta.get('name', f"{default_name_prefix}_{item_idx + 1}")
            point_type_val = effective_meta.get('sym', effective_meta.get('icon', 'Landmark'))
            
            # --- DEBUG: Перевірка метаданих для точки ---
            print(f"DEBUG: effective_meta для точки '{final_name}': {effective_meta}")
            print(f"DEBUG: Значення 'color' у effective_meta: {effective_meta.get('color')}")
            print(f"DEBUG: Значення 'icon' у effective_meta: {effective_meta.get('icon')}")
            print(f"DEBUG: Значення 'sym' у effective_meta: {effective_meta.get('sym')}")
            print(f"DEBUG: Значення 'apq_icon_id': {apq_icon_id}")
            print(f"DEBUG: Значення 'apq_color_int': {apq_color_int}")
            # --- Кінець DEBUG ---

            determined_color_value = None

            # 1. Пріоритет: icon_id з ApqFile (якщо є)
            if apq_icon_id is not None and apq_icon_id in self.ICON_ID_COLOR_MAP:
                rgba_tuple = self.ICON_ID_COLOR_MAP[apq_icon_id]
                determined_color_value = f"#{rgba_tuple[0]:02x}{rgba_tuple[1]:02x}{rgba_tuple[2]:02x}"
                print(f"DEBUG: Колір знайдено через ICON_ID ({apq_icon_id}): {determined_color_value}")

            # 2. Пріоритет: color_int з ApqFile (якщо є)
            if determined_color_value is None and apq_color_int is not None:
                r = (apq_color_int >> 16) & 0xFF
                g = (apq_color_int >> 8) & 0xFF
                b = apq_color_int & 0xFF
                determined_color_value = f"#{r:02x}{g:02x}{b:02x}"
                print(f"DEBUG: Колір знайдено через APQ color_int ({apq_color_int}): {determined_color_value}")

            # 3. Пріоритет: поля 'color', 'icon', 'sym' з effective_meta
            if determined_color_value is None: # Якщо колір досі не визначено
                meta_color_val = effective_meta.get('color')
                meta_icon_val = effective_meta.get('icon')
                meta_sym_val = effective_meta.get('sym')

                if meta_color_val is not None:
                    # Важливо: використовуємо self.convert_color, але його потрібно вдосконалити,
                    # щоб він повертав None, якщо колір не розпізнано, а не "White", щоб не перешкоджати.
                    # Поки що, якщо convert_color повертає White, це може бути або дійсний білий, або нерозпізнаний.
                    test_color_hex = self.convert_color(meta_color_val, "hex", True)
                    # Якщо convert_color не розпізнав і дав білий за замовчуванням,
                    # ми не вважаємо це знайденим кольором на цьому етапі.
                    if test_color_hex and test_color_hex != "#ffffff": 
                        determined_color_value = test_color_hex
                        print(f"DEBUG: Колір знайдено через effective_meta['color']: {determined_color_value}")
                
                if determined_color_value is None and meta_icon_val is not None:
                    # Шукаємо назву кольору в назві іконки
                    for color_name_key, color_hex_val in self.colors.items():
                        if color_name_key.lower() in str(meta_icon_val).lower():
                            determined_color_value = color_hex_val
                            print(f"DEBUG: Колір знайдено через effective_meta['icon']: {determined_color_value}")
                            break

                if determined_color_value is None and meta_sym_val is not None:
                    # Шукаємо назву кольору в назві символу
                    for color_name_key, color_hex_val in self.colors.items():
                        if color_name_key.lower() in str(meta_sym_val).lower():
                            determined_color_value = color_hex_val
                            print(f"DEBUG: Колір знайдено через effective_meta['sym']: {determined_color_value}")
                            break
            
            # 4. Назви кольорів у полі 'name'
            if determined_color_value is None:
                name_l = final_name.lower()
                for ru_name, en_name in self.extended_color_names_map.items(): # Використовуємо self.extended_color_names_map
                    if ru_name in name_l:
                        determined_color_value = self.colors.get(en_name, None) # Тепер повертаємо None, якщо не знайшли
                        if determined_color_value: # Перевіряємо, чи знайдено
                            print(f"DEBUG: Колір знайдено через назву у полі 'name' ({ru_name}): {determined_color_value}")
                            break
                if determined_color_value is None: # Якщо не знайдено за українськими/російськими назвами
                    for cname_en in self.colors.keys():
                        if cname_en.lower() in name_l:
                            determined_color_value = self.colors.get(cname_en, None) # Тепер повертаємо None, якщо не знайшли
                            if determined_color_value: # Перевіряємо, чи знайдено
                                print(f"DEBUG: Колір знайдено через англ. назву у полі 'name' ({cname_en}): {determined_color_value}")
                                break

            # Використовуємо знайдений HEX-колір, або білий за замовчуванням
            # convert_color(None, "name", True) поверне "White"
            final_point_color_name = self.convert_color(determined_color_value, "name", True)
            
            description_val = effective_meta.get('comment', effective_meta.get('description', ''))

            extra_desc_parts = []
            if loc_data.get('alt') is not None: extra_desc_parts.append(f"Висота: {loc_data['alt']:.1f}м")
            if loc_data.get('ts') is not None:
                try:
                    dt_obj = datetime.fromtimestamp(loc_data['ts'], timezone.utc)
                    extra_desc_parts.append(f"Час: {dt_obj.strftime('%Y-%m-%d %H:%M:%S %Z')}")
                except:
                    pass
            if loc_data.get('acc') is not None: extra_desc_parts.append(f"Точність: {loc_data['acc']:.1f}м")

            full_description = str(description_val) if description_val else ""
            if extra_desc_parts:
                full_description = (full_description + " | " if full_description else "") + "; ".join(extra_desc_parts)

            entry = {
                "name": final_name, "lat": point_lat, "lon": point_lon,
                "type": point_type_val,
                "description": full_description if full_description else None,
                "geometry_type": "Point",
                "original_location_data": loc_data,
                "apq_original_type": apq_source_file_type_for_item,
                'milgeo:meta:name': final_name,
                'milgeo:meta:color': final_point_color_name, # Оновлено!
                'milgeo:meta:desc': description_val,
                'milgeo:meta:creator': effective_meta.get('creator'),
                'milgeo:meta:creator_url': effective_meta.get('creator_url'),
                'milgeo:meta:sidc': effective_meta.get('sidc', global_meta.get('sidc'))
            }
            return entry

        if apq_type == 'wpt':
            loc = apq_parsed_data.get('location')
            # Отримання icon_id та color_int з apq_parsed_data
            wpt_icon_id = apq_parsed_data.get('icon_id')
            wpt_color_int = apq_parsed_data.get('color_int')

            point = _create_point_dict(loc, global_meta, "Waypoint",
                                       apq_source_file_type_for_item='wpt',
                                       source_file_global_meta_for_item=global_meta,
                                       apq_icon_id=wpt_icon_id,
                                       apq_color_int=wpt_color_int)
            if point: normalized_content.append(point)

        elif apq_type in ['set', 'rte']:
            waypoints_list = apq_parsed_data.get('waypoints', [])
            default_prefix = global_meta.get('name', apq_type.upper())
            for idx, wpt_entry in enumerate(waypoints_list):
                # Для SET/RTE, icon_id та color_int повинні бути в самому wpt_entry,
                # якщо _get_waypoints в apq.py їх зчитав.
                point = _create_point_dict(
                    wpt_entry.get('location'), wpt_entry.get('meta', {}),
                    default_prefix, idx, apq_source_file_type_for_item=apq_type,
                    source_file_global_meta_for_item=global_meta,
                    apq_icon_id=wpt_entry.get('icon_id'), # Передача icon_id з wpt_entry
                    apq_color_int=wpt_entry.get('color_int') # Передача color_int з wpt_entry
                )
                if point: normalized_content.append(point)

        elif apq_type == 'are':
            locations_list = apq_parsed_data.get('locations', [])
            area_name = global_meta.get('name', 'Area')
            area_color_name = self.convert_color(global_meta.get('color', 'Blue'), "name", True)
            area_description = global_meta.get('comment', global_meta.get('description', ''))
            area_points_data_for_polygon = [loc for loc in locations_list if loc and loc.get('lon') is not None]

            if len(area_points_data_for_polygon) >= 3:
                poly_item = {
                    'name': area_name, 'type': 'Area', 'geometry_type': 'Polygon',
                    'points_data': area_points_data_for_polygon,
                    'apq_original_type': 'are',
                    'milgeo:meta:name': area_name, 'milgeo:meta:color': area_color_name,
                    'milgeo:meta:desc': area_description,
                    'milgeo:meta:creator': global_meta.get('creator'),
                    'milgeo:meta:creator_url': global_meta.get('creator_url'),
                    'milgeo:meta:sidc': global_meta.get('sidc')
                }
                normalized_content.append(poly_item)
                # --- DEBUG: Перевірка для полігона ---
                print(f"DEBUG: Poly item '{area_name}' area_color_name: {area_color_name}")
                print(f"DEBUG: Global meta for area: {global_meta}")
                # --- Кінець DEBUG ---

        elif apq_type == 'trk':
            track_default_name = global_meta.get('name', 'Track')
            for idx, poi_entry in enumerate(apq_parsed_data.get('waypoints', [])):
                # Для POI в TRK, icon_id та color_int повинні бути в самому poi_entry,
                # якщо _get_waypoints в apq.py їх зчитав.
                point = _create_point_dict(
                    poi_entry.get('location'), poi_entry.get('meta', {}),
                    f"{track_default_name}_POI", idx,
                    apq_source_file_type_for_item='trk_poi',
                    source_file_global_meta_for_item=global_meta,
                    apq_icon_id=poi_entry.get('icon_id'), # Передача icon_id з poi_entry
                    apq_color_int=poi_entry.get('color_int') # Передача color_int з poi_entry
                )
                if point: normalized_content.append(point)

            segments_data = apq_parsed_data.get('segments', [])
            for seg_idx, segment_item in enumerate(segments_data):
                seg_locs_data = segment_item.get('locations', [])
                seg_meta_data = segment_item.get('meta', {})
                if len(seg_locs_data) < 2: continue

                effective_seg_meta = global_meta.copy()
                effective_seg_meta.update(seg_meta_data)

                segment_name = effective_seg_meta.get('name', f"{track_default_name}_Segment_{seg_idx + 1}")
                segment_color_name = self.convert_color(effective_seg_meta.get('color', 'Red'), "name", True)
                segment_description = effective_seg_meta.get('comment', effective_seg_meta.get('description', ''))
                segment_points_for_line = [loc for loc in seg_locs_data if loc and loc.get('lon') is not None]

                if len(segment_points_for_line) >= 2:
                    line_item = {
                        'name': segment_name, 'geometry_type': 'LineString',
                        'points_data': segment_points_for_line,
                        'apq_original_type': 'trk',
                        'milgeo:meta:name': segment_name, 'milgeo:meta:color': segment_color_name,
                        'milgeo:meta:desc': segment_description,
                        'milgeo:meta:creator': global_meta.get('creator'),
                        'milgeo:meta:creator_url': global_meta.get('creator_url'),
                        'milgeo:meta:sidc': effective_seg_meta.get('sidc', global_meta.get('sidc'))
                    }
                    normalized_content.append(line_item)
                    # --- DEBUG: Перевірка для лінії ---
                    print(f"DEBUG: Line item '{segment_name}' segment_color_name: {segment_color_name}")
                    print(f"DEBUG: Effective segment meta: {effective_seg_meta}")
                    # --- Кінець DEBUG ---

        if not normalized_content and apq_type not in ['ldk', 'bin']:
            self._update_status(f"Увага: Не знайдено даних для нормалізації у {file_basename} (тип {apq_type})",
                                warning=True)
        return normalized_content

    def _read_specific_file(self, file_path_to_read, expected_file_extension):
        self.input_file_path = file_path_to_read
        content_list = []
        file_basename_log = os.path.basename(file_path_to_read)
        try:
            apq_parser_instance = ApqFile(path=file_path_to_read, verbosity=3, gui_logger_func=self._update_status) # Змінено verbosity на 3

            if not apq_parser_instance.parse_successful:
                self._update_status(f"Помилка парсингу APQ файлу: {file_basename_log}", error=True)
                return None

            apq_data_structure = apq_parser_instance.data()
            content_list = self._normalize_apq_data(apq_data_structure, file_path_to_read)

        except ValueError as e:
            self._update_status(f"Помилка ініціалізації парсера для {file_path_to_read}: {e}", error=True)
            return None
        except Exception as e:
            self._update_status(f"Загальна помилка парсингу APQ {file_path_to_read}: {e}", error=True)
            print(f"Загальний виняток APQ ({file_path_to_read}): {type(e).__name__}: {e}")
            return None

        return content_list if content_list else None

    def read_wpt(self, path):
        return self._read_specific_file(path, ".wpt")

    def read_set(self, path):
        return self._read_specific_file(path, ".set")

    def read_rte(self, path):
        return self._read_specific_file(path, ".rte")

    def read_are(self, path):
        return self._read_specific_file(path, ".are")

    def read_trk(self, path):
        return self._read_specific_file(path, ".trk")

    def read_ldk(self, path):
        self._update_status(f"Читання LDK: {os.path.basename(path)}...", self.C_BUTTON_HOVER)
        all_normalized_content = []
        try:
            ldk_apq_file_instance = ApqFile(path=path, verbosity=1, gui_logger_func=self._update_status)
            if not ldk_apq_file_instance.parse_successful:
                messagebox.showerror("Помилка LDK", f"Не вдалося розпарсити LDK файл: {os.path.basename(path)}")
                return None

            parsed_ldk_root_data = ldk_apq_file_instance.data()

            def extract_and_normalize_from_ldk_node(node_data, parent_original_path):
                if not node_data: return

                for ldk_file_entry in node_data.get('files', []):
                    self._update_status(f"Обробка з LDK: {ldk_file_entry['name']}", self.C_BUTTON_HOVER)

                    file_content_bytes = base64.b64decode(ldk_file_entry['data_b64'])
                    inner_file_type = ldk_file_entry['type']
                    inner_file_name = ldk_file_entry['name']

                    if inner_file_type == 'bin':
                        self._update_status(f".bin з LDK: {inner_file_name}, експорт не підтримується.", warning=True)
                        continue

                    try:
                        contained_apq = ApqFile(rawdata=file_content_bytes,
                                                file_type=inner_file_type,
                                                rawname=inner_file_name,
                                                rawts=parsed_ldk_root_data.get('ts', time.time()),
                                                verbosity=0,
                                                gui_logger_func=self._update_status)

                        if contained_apq.parse_successful:
                            normalized_data = self._normalize_apq_data(contained_apq.data(), inner_file_name)
                            if normalized_data:
                                for item_norm in normalized_data:
                                    item_norm['source_file'] = inner_file_name
                                    item_norm['ldk_parent'] = os.path.basename(parent_original_path)
                                all_normalized_content.extend(normalized_data)
                        else:
                            self._update_status(f"Помилка парсингу файлу з LDK: {inner_file_name}", warning=True)
                    except Exception as e:
                        self._update_status(f"Помилка обробки {inner_file_name} з LDK: {e}", error=True)

                for child_node in node_data.get('nodes', []):
                    extract_and_normalize_from_ldk_node(child_node, parent_original_path)

            if parsed_ldk_root_data and parsed_ldk_root_data.get('root'):
                extract_and_normalize_from_ldk_node(parsed_ldk_root_data['root'], path)
            else:
                messagebox.showwarning("Увага LDK",
                                       f"LDK файл {os.path.basename(path)} порожній або має невірну структуру.")
                return None

        except ValueError as e:
            messagebox.showerror("Помилка LDK",
                                 f"Не вдалося ініціалізувати парсер для LDK {os.path.basename(path)}: {e}")
            return None
        except Exception as e:
            messagebox.showerror("Помилка читання LDK",
                                 f"Не вдалося обробити файл {os.path.basename(path)}.\n{type(e).__name__}: {e}")
            import traceback
            traceback.print_exc()
            return None

        return all_normalized_content if all_normalized_content else None

    def _read_kml_from_content(self, kml_content, source_filename="KML"):
        result = []
        try:
            if kml_content.startswith('\ufeff'):
                kml_content = kml_content[1:]
            root = ET.fromstring(kml_content)
            ns_match = re.match(r'\{([^}]+)\}', root.tag)
            ns_uri = ns_match.group(1) if ns_match else 'http://www.opengis.net/kml/2.2'
            ns = {'kml': ns_uri} if root.tag.startswith(f'{{{ns_uri}}}') else {}

            def find_tag(element, tag_name):
                if ns: return element.find(f'.//kml:{tag_name}', namespaces=ns)
                return element.find(f'.//{tag_name}')

            def findall_tags(element, tag_name):
                if ns: return element.findall(f'.//kml:{tag_name}', namespaces=ns)
                return element.findall(f'.//{tag_name}')

            for placemark in findall_tags(root, 'Placemark'):
                name_tag = find_tag(placemark, 'name')
                name = name_tag.text.strip() if name_tag is not None and name_tag.text else f'{source_filename}_Point'

                description_tag = find_tag(placemark, 'description')
                description = description_tag.text.strip() if description_tag is not None and description_tag.text else ""

                color = "White"
                style_url_tag = find_tag(placemark, 'styleUrl')
                inline_style_tag = find_tag(placemark, 'Style')
                parsed_color_hex = None

                if style_url_tag is not None and style_url_tag.text:
                    style_id = style_url_tag.text.lstrip('#')
                    style_node = root.find(f".//{{*}}Style[@id='{style_id}']") or root.find(
                        f".//{{*}}StyleMap[@id='{style_id}']//{{*}}Style")
                    if style_node is not None:
                        for style_type_node in [find_tag(style_node, 'IconStyle'), find_tag(style_node, 'LineStyle'),
                                                find_tag(style_node, 'PolyStyle')]:
                            if style_type_node is not None:
                                color_node = find_tag(style_type_node, 'color')
                                if color_node is not None and color_node.text:
                                    kml_color_str = color_node.text.strip().lower()
                                    if len(kml_color_str) == 8:
                                        parsed_color_hex = f"#{kml_color_str[6:8]}{kml_color_str[4:6]}{kml_color_str[2:4]}"
                                        break

                elif inline_style_tag is not None:
                    for style_type_node in [find_tag(inline_style_tag, 'IconStyle'),
                                            find_tag(inline_style_tag, 'LineStyle'),
                                            find_tag(inline_style_tag, 'PolyStyle')]:
                        if style_type_node is not None:
                            color_node = find_tag(style_type_node, 'color')
                            if color_node is not None and color_node.text:
                                kml_color_str = color_node.text.strip().lower()
                                if len(kml_color_str) == 8:
                                    parsed_color_hex = f"#{kml_color_str[6:8]}{kml_color_str[4:6]}{kml_color_str[2:4]}"
                                    break

                if parsed_color_hex:
                    color = self.convert_color(parsed_color_hex, "name", True)

                item_data = {"name": name, "color": color, "description": description, "source_file": source_filename}

                point_coords_tag = find_tag(placemark, 'Point/coordinates')
                linestring_coords_tag = find_tag(placemark, 'LineString/coordinates')
                polygon_outer_coords_tag = find_tag(placemark, 'Polygon/outerBoundaryIs/LinearRing/coordinates')

                if point_coords_tag is not None and point_coords_tag.text:
                    parts = point_coords_tag.text.strip().split(',')
                    if len(parts) >= 2:
                        try:
                            lon, lat = float(parts[0]), float(parts[1])
                            alt = float(parts[2]) if len(parts) > 2 else 0.0
                            item_data.update({"lat": lat, "lon": lon, "type": "Landmark", "geometry_type": "Point",
                                              "original_location_data": {"alt": alt}})
                            result.append(item_data)
                        except ValueError:
                            pass

                elif linestring_coords_tag is not None and linestring_coords_tag.text:
                    points_data = [{'lon': float(p.split(',')[0]), 'lat': float(p.split(',')[1]),
                                    'alt': float(p.split(',')[2]) if len(p.split(',')) > 2 else 0.0} for p in
                                   linestring_coords_tag.text.strip().split() if len(p.split(',')) >= 2]
                    if len(points_data) >= 2:
                        item_data.update(
                            {"type": "LineString", "geometry_type": "LineString", "points_data": points_data})
                        result.append(item_data)

                elif polygon_outer_coords_tag is not None and polygon_outer_coords_tag.text:
                    points_data = [{'lon': float(p.split(',')[0]), 'lat': float(p.split(',')[1]),
                                    'alt': float(p.split(',')[2]) if len(p.split(',')) > 2 else 0.0} for p in
                                   polygon_outer_coords_tag.text.strip().split() if len(p.split(',')) >= 2]
                    if len(points_data) >= 3:
                        item_data.update({"type": "Area", "geometry_type": "Polygon", "points_data": points_data})
                        result.append(item_data)

        except ET.ParseError as e:
            self._update_status(f"Помилка парсингу KML: {source_filename} ({e})", error=True)
            return None
        except Exception as e:
            self._update_status(f"Загальна помилка читання KML: {source_filename} ({e})", error=True)
            return None

        return result if result else None

    def read_kml(self, path):
        try:
            with open(path, "r", encoding="utf-8-sig") as fo:
                kml_content = fo.read()
            return self._read_kml_from_content(kml_content, os.path.basename(path))
        except Exception as e:
            self._update_status(f"Помилка KML: {os.path.basename(path)}: {e}", error=True)
            return None

    def read_kmz(self, path):
        try:
            with zipfile.ZipFile(path, 'r') as kmz:
                kml_file_name = next((f for f in kmz.namelist() if f.lower().endswith('.kml')), None)
                if kml_file_name:
                    with kmz.open(kml_file_name) as kml_file_stream:
                        kml_content = kml_file_stream.read().decode('utf-8', errors='replace')
                    return self._read_kml_from_content(kml_content, kml_file_name)
                else:
                    self._update_status(f"У файлі KMZ {os.path.basename(path)} не знайдено .kml файл.", warning=True)
                    return None
        except zipfile.BadZipFile:
            self._update_status(f"Файл {os.path.basename(path)} пошкоджений або не є KMZ архівом.", error=True)
            return None
        except Exception as e:
            self._update_status(f"Помилка KMZ: {os.path.basename(path)}: {e}", error=True)
            return None

    def read_gpx(self, path):
        try:
            # Реєструємо поширені простори імен, щоб парсер їх розумів
            namespaces = {
                'gpx': 'http://www.topografix.com/GPX/1/1',
                'gpxx': 'http://www.garmin.com/xmlschemas/GpxExtensions/v3'
            }
            ET.register_namespace('gpxx', namespaces['gpxx'])

            tree = ET.parse(path)
            root = tree.getroot()

            result = []

            # Обробка маршрутних точок (waypoints)
            for wpt in root.findall('gpx:wpt', namespaces):
                name = wpt.find('gpx:name', namespaces).text if wpt.find('gpx:name',
                                                                         namespaces) is not None else 'GPX Waypoint'
                desc = wpt.find('gpx:desc', namespaces).text if wpt.find('gpx:desc', namespaces) is not None else ''

                # Пошук кольору в розширеннях
                color_name = "White"  # Колір за замовчуванням
                try:
                    color_tag = wpt.find('.//gpxx:DisplayColor', namespaces)
                    if color_tag is not None and color_tag.text:
                        # Конвертуємо назву кольору з GPX у нашу палітру
                        color_name = self.convert_color(color_tag.text, "name")
                except Exception:
                    pass  # Залишаємо колір за замовчуванням, якщо щось пішло не так

                try:
                    lat, lon = float(wpt.get('lat')), float(wpt.get('lon'))
                    alt = float(wpt.find('gpx:ele', namespaces).text) if wpt.find('gpx:ele',
                                                                                  namespaces) is not None else 0.0
                    result.append({
                        "name": name, "lat": lat, "lon": lon, "type": "Waypoint",
                        "color": color_name, "description": desc,
                        "geometry_type": "Point", "original_location_data": {"alt": alt}
                    })
                except (ValueError, TypeError):
                    continue

            # Обробка треків
            for trk in root.findall('gpx:trk', namespaces):
                trk_name = trk.find('gpx:name', namespaces).text if trk.find('gpx:name',
                                                                             namespaces) is not None else 'GPX Track'
                for i, trkseg in enumerate(trk.findall('gpx:trkseg', namespaces)):
                    points_data = []
                    for trkpt in trkseg.findall('gpx:trkpt', namespaces):
                        try:
                            lat, lon = float(trkpt.get('lat')), float(trkpt.get('lon'))
                            alt = float(trkpt.find('gpx:ele', namespaces).text) if trkpt.find('gpx:ele',
                                                                                              namespaces) is not None else 0.0
                            points_data.append({'lon': lon, 'lat': lat, 'alt': alt})
                        except (ValueError, TypeError):
                            continue

                    if len(points_data) >= 2:
                        result.append({
                            "name": f"{trk_name} - Сегмент {i + 1}", "type": "TrackSegment",
                            "geometry_type": "LineString", "color": "Red",
                            # Треки зазвичай не мають кольору, залишаємо червоний
                            "description": "", "points_data": points_data
                        })

            return result if result else None
        except Exception as e:
            self._update_status(f"Помилка читання GPX: {os.path.basename(path)}: {e}", error=True)
            return None

    def read_xlsx(self, path):
        try:
            workbook = openpyxl.load_workbook(path, data_only=True)
            result = []
            for sheet in workbook.worksheets:
                if sheet.max_row < 2: continue

                header = [str(cell.value).lower().strip() if cell.value else '' for cell in sheet[1]]

                # Гнучкий пошук колонок
                lat_aliases = ['lat', 'latitude', 'широта', 'y']
                lon_aliases = ['lon', 'long', 'longitude', 'довгота', 'x']
                name_aliases = ['name', 'title', 'назва', 'id']
                color_aliases = ['color', 'колір', 'цвет']  # Додаємо аліаси для кольору

                try:
                    lat_col = next(i for i, h in enumerate(header) if h in lat_aliases)
                    lon_col = next(i for i, h in enumerate(header) if h in lon_aliases)
                    name_col = next((i for i, h in enumerate(header) if h in name_aliases), -1)
                    # Шукаємо колонку кольору, якщо її немає - це не помилка
                    color_col = next((i for i, h in enumerate(header) if h in color_aliases), -1)
                except StopIteration:
                    self._update_status(f"XLSX: пропуск аркуша '{sheet.title}' (немає колонок lat/lon)", warning=True)
                    continue

                for r_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
                    try:
                        lat, lon = float(str(row[lat_col]).replace(',', '.')), float(
                            str(row[lon_col]).replace(',', '.'))
                        if not (-90 <= lat <= 90 and -180 <= lon <= 180): continue

                        name = str(row[name_col]) if name_col != -1 and row[
                            name_col] else f'{sheet.title}_Point_{r_idx - 1}'

                        # Отримуємо колір з колонки, якщо вона є, інакше - білий
                        color_value = str(row[color_col]) if color_col != -1 and row[color_col] else "White"
                        color_name = self.convert_color(color_value, "name")

                        desc_parts = [f"{h.capitalize()}: {v}" for h, v in zip(header, row) if
                                      v is not None and header.index(h) not in [lat_col, lon_col, name_col, color_col]]
                        desc = "; ".join(desc_parts)

                        result.append({"name": name, "lat": lat, "lon": lon, "type": "XLSX Point", "color": color_name,
                                       "description": desc, "geometry_type": "Point"})
                    except (ValueError, TypeError, IndexError):
                        continue
            return result if result else None
        except Exception as e:
            self._update_status(f"Помилка XLSX: {os.path.basename(path)}: {e}", error=True)
            return None

    def read_csv(self, path):
        result = []
        try:
            with open(path, mode='r', encoding='utf-8-sig') as infile:
                try:
                    dialect = csv.Sniffer().sniff(infile.read(2048))
                    infile.seek(0)
                except csv.Error:
                    dialect = 'excel'
                    infile.seek(0)

                reader = csv.DictReader(infile, dialect=dialect)
                if not reader.fieldnames: return None

                h_map = {h.lower().strip(): h for h in reader.fieldnames}

                lat_key = next((h_map[alias] for alias in ['lat', 'latitude', 'широта', 'y'] if alias in h_map), None)
                lon_key = next(
                    (h_map[alias] for alias in ['lon', 'long', 'longitude', 'довгота', 'x'] if alias in h_map), None)
                name_key = next((h_map[alias] for alias in ['name', 'title', 'назва', 'id'] if alias in h_map), None)
                desc_key = next((h_map[alias] for alias in ['desc', 'description', 'опис'] if alias in h_map), None)
                # Шукаємо колонку кольору
                color_key = next((h_map[alias] for alias in ['color', 'колір', 'цвет'] if alias in h_map), None)

                if not lat_key or not lon_key:
                    self._update_status("CSV: Не знайдено колонок для широти/довготи.", error=True)
                    return None

                for i, row in enumerate(reader, 2):
                    try:
                        lat, lon = float(str(row[lat_key]).replace(',', '.')), float(
                            str(row[lon_key]).replace(',', '.'))
                        if not (-90 <= lat <= 90 and -180 <= lon <= 180): continue

                        name = row.get(name_key, f'CSV_Point_{i - 1}')

                        # Отримуємо колір з колонки або білий за замовчуванням
                        color_value = row.get(color_key, "White")
                        color_name = self.convert_color(color_value, "name")

                        other_cols = [k for k in row.keys() if
                                      k not in [lat_key, lon_key, name_key, color_key, desc_key]]
                        desc = row.get(desc_key, "") + "; ".join([f"{k}: {row[k]}" for k in other_cols if row[k]])

                        result.append({"name": name, "lat": lat, "lon": lon, "type": "CSV Point", "color": color_name,
                                       "description": desc, "geometry_type": "Point"})
                    except (ValueError, TypeError, KeyError):
                        continue
            return result
        except Exception as e:
            self._update_status(f"Помилка CSV: {os.path.basename(path)}: {e}", error=True)
            return None

    def read_scene(self, path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            result = []
            for item in data.get("scene", {}).get("items", []):
                pos = item.get("position", {})
                if "lat" in pos and "lon" in pos:
                    try:
                        result.append({"name": str(item.get("name", "SCENE Point")), "lon": float(pos["lon"]),
                                       "lat": float(pos["lat"]), "type": str(item.get("type", "Landmark")),
                                       "color": self.convert_color(str(item.get("color", "White")), "name", True),
                                       "description": str(item.get("description", "")), "geometry_type": "Point"})
                    except (ValueError, TypeError):
                        continue
            return result if result else None
        except Exception as e:
            self._update_status(f"Помилка .scene: {os.path.basename(path)}: {e}", error=True)
            return None

    def read_geojson(self, path):
        result = []
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            features = data.get("features", []) if data.get("type") == "FeatureCollection" else [data] if data.get(
                "type") == "Feature" else []
            for idx, feature in enumerate(features):
                if not isinstance(feature, dict) or feature.get("type") != "Feature": continue
                geom, props = feature.get("geometry", {}), feature.get("properties", {})
                if not geom or not props: continue

                name = props.get("name", props.get("title", f"GeoFeature_{idx + 1}"))
                color = self.convert_color(str(props.get("color", props.get("stroke", "#ffffff"))), "name", True)
                desc = props.get("description", "")
                item_base = {"name": str(name), "color": color, "description": str(desc),
                             "source_file": os.path.basename(path)}

                geom_type, coords = geom.get("type"), geom.get("coordinates")
                if geom_type == "Point" and coords and len(coords) >= 2:
                    try:
                        item_base.update({"lat": float(coords[1]), "lon": float(coords[0]), "type": "Landmark",
                                          "geometry_type": "Point"})
                        result.append(item_base)
                    except (ValueError, TypeError):
                        pass
                elif geom_type == "LineString" and coords and len(coords) >= 2:
                    points_data = [{'lon': c[0], 'lat': c[1]} for c in coords if len(c) >= 2]
                    if len(points_data) >= 2:
                        item_base.update(
                            {"type": "LineString", "geometry_type": "LineString", "points_data": points_data})
                        result.append(item_base)
                elif geom_type == "Polygon" and coords and len(coords[0]) >= 3:
                    points_data = [{'lon': c[0], 'lat': c[1]} for c in coords[0] if len(c) >= 2]
                    if len(points_data) >= 3:
                        item_base.update({"type": "Area", "geometry_type": "Polygon", "points_data": points_data})
                        result.append(item_base)
        except Exception as e:
            self._update_status(f"Помилка GeoJSON: {os.path.basename(path)}: {e}", error=True)
            return None
        return result if result else None

    def create_scene(self, contents_list, save_path):
        if not contents_list: return False
        items_data = []
        for item in contents_list:
            if item.get('geometry_type') == 'Point':
                items_data.append({"color": str(item.get("color", "White")), "creationDate": int(time.time() * 1000),
                                   "name": str(item.get("name", "N/A")),
                                   "position": {"alt": 0.0, "lat": float(item.get("lat", 0.0)),
                                                "lon": float(item.get("lon", 0.0))},
                                   "type": str(item.get("type", "Landmark"))})
        scene_obj = {"scene": {"items": items_data, "name": os.path.splitext(os.path.basename(save_path))[0]},
                     "version": 7}
        try:
            with open(save_path, "w", encoding="UTF-8") as f:
                json.dump(scene_obj, f, ensure_ascii=False, separators=(',', ':')); return True
        except IOError:
            return False

    def create_kml(self, contents_list, save_path):
        if not contents_list: return False
        try:
            with open(save_path, "w", encoding="UTF-8") as f:
                f.write(self._create_kml_string(contents_list,
                                                os.path.splitext(os.path.basename(save_path))[0])); return True
        except IOError:
            return False

    def _create_kml_string(self, contents_list, doc_name):
        kml_doc = ET.Element("kml", xmlns="http://www.opengis.net/kml/2.2")
        document = ET.SubElement(kml_doc, "Document")
        ET.SubElement(document, "name").text = doc_name
        style_map = {}
        for item in contents_list:
            color_name = item.get("color", "White")
            if color_name not in style_map:
                color_hex = self.colors.get(color_name, "#ffffff").lstrip('#')
                kml_color = f"ff{color_hex[4:6]}{color_hex[2:4]}{color_hex[0:2]}"
                style_id = f"style_{color_name}"
                style_map[color_name] = style_id
                style = ET.SubElement(document, "Style", id=style_id)
                ET.SubElement(style, "IconStyle").append(ET.fromstring(
                    f"<color>{kml_color}</color><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon>"))
                ET.SubElement(style, "LineStyle").append(ET.fromstring(f"<color>{kml_color}</color><width>2</width>"))
                ET.SubElement(style, "PolyStyle").append(
                    ET.fromstring(f"<color>7f{color_hex[4:6]}{color_hex[2:4]}{color_hex[0:2]}</color>"))

        for item in contents_list:
            placemark = ET.SubElement(document, "Placemark")
            ET.SubElement(placemark, "name").text = xml_escape(item.get("name", "N/A"))
            if item.get("description"): ET.SubElement(placemark, "description").text = xml_escape(
                item.get("description"))
            ET.SubElement(placemark, "styleUrl").text = f"#{style_map.get(item.get('color', 'White'), 'style_White')}"

            geom_type = item.get("geometry_type")
            if geom_type == "Point":
                ET.SubElement(placemark, "Point").append(
                    ET.fromstring(f"<coordinates>{item.get('lon', 0)},{item.get('lat', 0)},0</coordinates>"))
            elif geom_type in ["LineString", "Polygon"]:
                coords_data = item.get('points_data', [])
                coords_str = " ".join(f"{p['lon']},{p['lat']},0" for p in coords_data)
                if geom_type == "Polygon" and coords_data and coords_data[0] != coords_data[-1]:
                    coords_str += f" {coords_data[0]['lon']},{coords_data[0]['lat']},0"

                # Коректне створення елементів для Polygon KML
                if geom_type == "Polygon":
                    polygon_elem = ET.SubElement(placemark, "Polygon")
                    outer_boundary = ET.SubElement(polygon_elem, "outerBoundaryIs")
                    linear_ring = ET.SubElement(outer_boundary, "LinearRing")
                    ET.SubElement(linear_ring, "coordinates").text = coords_str
                else: # LineString
                    linestring_elem = ET.SubElement(placemark, "LineString")
                    ET.SubElement(linestring_elem, "coordinates").text = coords_str

        ET.indent(kml_doc, space="  ")
        return '<?xml version="1.0" encoding="UTF-8"?>\n' + ET.tostring(kml_doc, encoding='unicode')

    def convert_color(self, input_value: Any, output_type: str = "hex", allow_name_lookup: bool = False) -> Optional[str]:
        """
        Конвертує колір між різними форматами (назва, hex, rgb, int ARGB, icon_id).
        Використовує єдину палітру кольорів.
        - input_value: може бути назвою (str), hex-кодом (str), rgb-кортежем (tuple),
                       цілим числом (int - для ARGB), або icon_id (int).
        - output_type: 'hex', 'name', 'int_rgb', 'str_rgb'.
        - allow_name_lookup: дозволяє пошук за назвою в палітрі.
        """
        if input_value is None:
            return None

        # Створюємо реверсивні словники для швидкого пошуку за HEX
        hex_to_name = {v.lower(): k for k, v in self.colors.items()}

        input_hex = ""

        # 1. Спроба розпізнати колір, якщо input_value - це ціле число (ARGB або icon_id)
        if isinstance(input_value, int):
            # Якщо input_value - це ідентифікатор іконки з ICON_ID_COLOR_MAP
            if input_value in self.ICON_ID_COLOR_MAP:
                rgba_tuple = self.ICON_ID_COLOR_MAP[input_value]
                r, g, b = rgba_tuple[0], rgba_tuple[1], rgba_tuple[2]
                input_hex = f"#{r:02x}{g:02x}{b:02x}"
            # Якщо input_value - це цілочисельний ARGB колір (наприклад, з wp.color)
            else:
                a = (input_value >> 24) & 0xFF
                r = (input_value >> 16) & 0xFF
                g = (input_value >> 8) & 0xFF
                b = (input_value) & 0xFF
                input_hex = f"#{r:02x}{g:02x}{b:02x}"
        
        elif isinstance(input_value, str):
            value_lower = input_value.lower().strip()
            # Пошук за назвою в нашій палітрі (точною назвою)
            if value_lower in {k.lower() for k in self.colors}:
                input_hex = self.colors[[k for k in self.colors if k.lower() == value_lower][0]].lower()
            # Пошук за HEX-кодом
            elif value_lower.startswith("#") and len(value_lower) == 7:
                input_hex = value_lower
            # Розбір RGBA/RGB/HEX з текстового рядка (як у smart_color, використовуючи Regex)
            else:
                # HEX у тексті (наприклад, "text #FF00FF more text")
                m = re.search(r'#?([0-9a-fA-F]{6})', value_lower)
                if m:
                    hx = m.group(1)
                    input_hex = f"#{hx}"
                else:
                    # RGBA у тексті (наприклад, "255,0,255,1.0")
                    m = re.search(r'(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d*\.?\d+)', value_lower)
                    if m:
                        r, g, b = int(m.group(1)), int(m.group(2)), int(m.group(3))
                        input_hex = f"#{r:02x}{g:02x}{b:02x}"
                    else:
                        # RGB у тексті (наприклад, "255,0,255")
                        m = re.search(r'(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})', value_lower)
                        if m:
                            r, g, b = int(m.group(1)), int(m.group(2)), int(m.group(3))
                            input_hex = f"#{r:02x}{g:02x}{b:02x}"
                
                # Якщо після всіх regex все ще не знайшли HEX, спробуємо назву кольору з розширеної мапи
                if not input_hex:
                    for ru_name, en_name in self.extended_color_names_map.items():
                        if ru_name in value_lower:
                            input_hex = self.colors.get(en_name, None) # Тепер повертаємо None, якщо не знайшли
                            if input_hex: # Перевіряємо, чи знайдено
                                break
                    
        elif isinstance(input_value, (list, tuple)) and len(input_value) == 3: # RGB кортеж або список
            try:
                # Конвертуємо RGB в HEX
                input_hex = f"#{input_value[0]:02x}{input_value[1]:02x}{input_value[2]:02x}"
            except (TypeError, ValueError):
                pass

        if not input_hex: # Якщо досі не визначили HEX
            input_hex = "#ffffff" # За замовчуванням - білий

        # Конвертуємо з HEX у потрібний вихідний формат
        if output_type.lower() == 'hex':
            return input_hex
        elif output_type.lower() == 'name':
            found_name = hex_to_name.get(input_hex, None)
            if found_name:
                return found_name
            elif allow_name_lookup:
                # Якщо немає точного співпадіння, але дозволено пошук, повертаємо білий.
                # Якщо потрібно складніше порівняння (наприклад, відстань у RGB),
                # то потрібно додати функцію find_nearest_color_name_from_hex.
                return "White" 
            return "White" # За замовчуванням
        elif output_type.lower() == 'int_rgb':
            h = input_hex.lstrip('#')
            try:
                return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
            except (TypeError, ValueError):
                return (255, 255, 255)
        elif output_type.lower() == 'str_rgb':
            h = input_hex.lstrip('#')
            try:
                rgb = tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
                return f"{rgb[0]},{rgb[1]},{rgb[2]}"
            except (TypeError, ValueError):
                return "255,255,255"
        return None


    def create_kmz(self, contents_list, save_path):
        if not contents_list: return False
        kml_content = self._create_kml_string(contents_list, os.path.splitext(os.path.basename(save_path))[0])
        try:
            with zipfile.ZipFile(save_path, 'w', zipfile.ZIP_DEFLATED) as kmz:
                kmz.writestr('doc.kml', kml_content); return True
        except Exception:
            return False

    def create_gpx(self, contents_list, save_path):
        if not contents_list: return False
        gpx = ET.Element('gpx', version="1.1", creator="Nexus", xmlns="http://www.topografix.com/GPX/1/1")
        for item in contents_list:
            if item.get('geometry_type') == 'Point':
                wpt = ET.SubElement(gpx, 'wpt', lat=str(item.get("lat")), lon=str(item.get("lon")))
                ET.SubElement(wpt, 'name').text = item.get("name")
            elif item.get('geometry_type') == 'LineString':
                trk = ET.SubElement(gpx, 'trk')
                ET.SubElement(trk, 'name').text = item.get("name")
                trkseg = ET.SubElement(trk, 'trkseg')
                for p in item.get('points_data', []): ET.SubElement(trkseg, 'trkpt', lat=str(p['lat']),
                                                                    lon=str(p['lon']))
        tree = ET.ElementTree(gpx)
        ET.indent(tree, space="  ")
        try:
            tree.write(save_path, encoding='utf-8', xml_declaration=True); return True
        except IOError:
            return False

    def create_xlsx(self, contents_list, save_path, split_by_colors=False):
        if not contents_list: return False
        try:
            workbook = xlsxwriter.Workbook(save_path)
        except xlsxwriter.exceptions.FileCreateError:
            self._update_status(f"Помилка XLSX (файл зайнятий?)", error=True); return False

        headers = ["NAME", "LAT", "LON", "TYPE", "COLOR", "DESC", "GEOMETRY_TYPE", "WKT"]
        header_format = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'center'})

        def write_sheet(ws, data):
            ws.write_row(0, 0, headers, header_format)
            for r, item in enumerate(data, 1):
                geom_type = item.get("geometry_type")
                wkt = ""
                if geom_type == "Point":
                    wkt = f"POINT ({item.get('lon')} {item.get('lat')})"
                elif geom_type in ["LineString", "Polygon"]:
                    pts = ", ".join(f"{p['lon']} {p['lat']}" for p in item.get('points_data', []))
                    if geom_type == "Polygon" and item.get('points_data'):
                        p_data = item['points_data']
                        if p_data[0]['lon'] != p_data[-1]['lon'] or p_data[0]['lat'] != p_data[-1]['lat']:
                            pts += f", {p_data[0]['lon']} {p_data[0]['lat']}"
                        wkt = f"POLYGON (({pts}))"
                    else:
                        wkt = f"LINESTRING ({pts})"
                row_data = [item.get(k, '') for k in ['name', 'lat', 'lon', 'type', 'color', 'description']] + [
                    geom_type, wkt]
                ws.write_row(r, 0, row_data)
            ws.autofit()

        if split_by_colors:
            data_by_color = {}
            for item in contents_list: data_by_color.setdefault(item.get('color', 'NoColor'), []).append(item)
            for color, data in data_by_color.items(): write_sheet(workbook.add_worksheet(color[:31]), data)
        else:
            write_sheet(workbook.add_worksheet("Data"), contents_list)

        try:
            workbook.close(); return True
        except xlsxwriter.exceptions.FileCreateError:
            return False

    def create_csv(self, contents_list, save_path):
        source_type = contents_list[0].get('apq_original_type') if contents_list else None
        if source_type in ['wpt', 'set', 'rte', 'are']:
            # Використовуйте саме create_csv_for_points_simple!
            return self.create_csv_for_points_simple(contents_list, save_path)
        else:
            return self.create_csv_original_logic(contents_list, save_path)

        if source_type == 'trk':
            return self.create_csv_for_trk(contents_list, save_path)
        elif source_type == 'set':
            return self.create_csv_for_set(contents_list, save_path)
        elif source_type in ['wpt', 'rte', 'are']:
            # Forcing `generate_rey_names` to False for ARE as it represents an area, not numbered points
            generate_names = True if source_type != 'are' else False
            return self.create_csv_for_points(contents_list, save_path, generate_rey_names=generate_names)
        else:
            # Fallback for data from KML, GPX, XLSX etc.
            return self.create_csv_original_logic(contents_list, save_path)
    
    # !!! Функція get_best_color_for_item була видалена, її логіка інтегрована в convert_color

    def create_csv_for_points_simple(self, contents_list, base_save_path):
        self._update_status(f"Створення CSV (простий) для точок: {os.path.basename(base_save_path)}...",
                            self.C_BUTTON_HOVER)
        headers = [
            "color", "coordinates", "milgeo:meta:color", "milgeo:meta:creator", "milgeo:meta:creator_url",
            "milgeo:meta:desc", "milgeo:meta:name", "name", "observation_datetime", "sidc"
        ]

        if not contents_list:
            self._update_status(f"Немає даних для CSV: {os.path.basename(base_save_path)}", warning=True)
            return False

        try:
            total_items = len(contents_list)
            for chunk_index, i in enumerate(range(0, total_items, self.CSV_CHUNK_SIZE)):
                chunk_contents = contents_list[i:i + self.CSV_CHUNK_SIZE]
                current_save_path = self._get_chunked_save_path(base_save_path, chunk_index)
                with open(current_save_path, "w", encoding="UTF-8", newline='') as f_out:
                    writer = csv.writer(f_out, quoting=csv.QUOTE_ALL)
                    writer.writerow(headers)
                    for item in chunk_contents:
                        if item.get('geometry_type') != 'Point':
                            continue

                        # Використовуємо вже визначений колір з item['color']
                        color_to_export = item.get('color', 'White') 
                        color_str = self.convert_color(color_to_export, 'str_rgb', True) + ',1'
                        wkt_string = f"POINT ({item.get('lon', 0.0)} {item.get('lat', 0.0)})"

                        def meta_json(val):
                            if isinstance(val, list):
                                return json.dumps(val, ensure_ascii=False)
                            elif val is None or val == "":
                                return "[]"
                            else:
                                return json.dumps([val], ensure_ascii=False)

                        meta_color = meta_json(item.get('milgeo:meta:color'))
                        meta_creator = meta_json(item.get('milgeo:meta:creator'))
                        meta_creator_url = meta_json(item.get('milgeo:meta:creator_url'))
                        meta_desc = meta_json(item.get('milgeo:meta:desc'))
                        meta_name = meta_json(item.get('milgeo:meta:name'))
                        name = item.get('name', '')
                        ts = item.get('original_location_data', {}).get('ts')
                        observation_datetime = ""
                        if ts:
                            observation_datetime = datetime.fromtimestamp(ts, timezone.utc).strftime(
                                '%Y-%m-%dT%H:%M:%S')
                        sidc = item.get('milgeo:meta:sidc') or ""

                        row = [
                            color_str, wkt_string, meta_color, meta_creator, meta_creator_url,
                            meta_desc, meta_name, name, observation_datetime, sidc
                        ]
                        writer.writerow(row)
            self._update_status(f"Файли CSV (простий) успішно збережено.", self.C_ACCENT_DONE)
            return True
        except Exception as e:
            self._update_status(f"Помилка CSV: {e}", error=True)
            return False

    def create_csv_for_set(self, contents_list, base_save_path):
        self._update_status(f"Створення CSV для SET: {os.path.basename(base_save_path)}...", self.C_BUTTON_HOVER)
        headers = ["sidc", "id", "quantity", "name", "observation_datetime", "reliability_credibility",
                   "staff_comments", "platform_type", "direction", "speed", "coordinates"]
        DEFAULT_SIDC = "10016600006099000000"

        if not contents_list: return False
        try:
            total_items = len(contents_list)
            for chunk_index, i in enumerate(range(0, total_items, self.CSV_CHUNK_SIZE)):
                chunk_contents = contents_list[i:i + self.CSV_CHUNK_SIZE]
                current_save_path = self._get_chunked_save_path(base_save_path, chunk_index)
                with open(current_save_path, "w", encoding="UTF-8", newline='') as f_out:
                    writer = csv.writer(f_out, quoting=csv.QUOTE_ALL)
                    writer.writerow(headers)
                    for item in chunk_contents:
                        if item.get('geometry_type') != 'Point': continue
                        name = item.get('name', '')
                        sidc = item.get('milgeo:meta:sidc') or DEFAULT_SIDC
                        comments_parts = []
                        if item.get('original_location_data', {}).get('ts'): comments_parts.append(
                            f"Час: {datetime.fromtimestamp(item['original_location_data']['ts'], timezone.utc).strftime('%Y-%m-%dT%H:%M:%S')}")
                        if item.get('milgeo:meta:color'): comments_parts.append(
                            f"Колір: {self.convert_color(item['milgeo:meta:color'], 'str_rgb', True)},1")
                        if item.get('milgeo:meta:desc'): comments_parts.append(item['milgeo:meta:desc'])
                        staff_comments = "; ".join(comments_parts)
                        wkt_string = f"POINT ({item.get('lon', 0.0)} {item.get('lat', 0.0)})"
                        writer.writerow([sidc, "", "", name, "", "", staff_comments, "", "", "", wkt_string])
            self._update_status(f"Файли CSV (SET) успішно збережено.", self.C_ACCENT_DONE)
            return True
        except Exception as e:
            self._update_status(f"Помилка CSV: {e}", error=True); return False

    def create_csv_for_trk(self, contents_list, base_save_path):
        self._update_status(f"Створення CSV для TRK: {os.path.basename(base_save_path)}...", self.C_BUTTON_HOVER)
        headers = ["sidc", "id", "quantity", "name", "observation_datetime", "reliability_credibility",
                   "staff_comments", "platform_type", "direction", "speed", "coordinates"]
        SIDC_BY_COLOR = {"Brown": "10036600001100000000", "Green": "10046600001100000000",
                         "Red": "10066600001100000000", "Yellow": "10016600001100000000"}
        DEFAULT_SIDC = "10066600001100000000"

        if not contents_list: return False
        try:
            total_items = len(contents_list)
            for chunk_index, i in enumerate(range(0, total_items, self.CSV_CHUNK_SIZE)):
                chunk_contents = contents_list[i:i + self.CSV_CHUNK_SIZE]
                current_save_path = self._get_chunked_save_path(base_save_path, chunk_index)
                with open(current_save_path, "w", encoding="UTF-8", newline='') as f_out:
                    writer = csv.writer(f_out, quoting=csv.QUOTE_ALL)
                    writer.writerow(headers)
                    for item in chunk_contents:
                        if item.get('geometry_type') != 'LineString': continue
                        track_name = item.get('name', '')
                        if 'Segment' in track_name or '_Трек_' in track_name or '_Track_' in track_name: track_name = ''
                        staff_comments = item.get('milgeo:meta:desc', '')
                        color_en = item.get('milgeo:meta:color', '')
                        other_meta_parts = [f"Колір: {self.colors_en_ua.get(color_en, color_en)}"] if color_en else []
                        if item.get('milgeo:meta:creator', ''): other_meta_parts.append(
                            f"Автор: {item['milgeo:meta:creator']}")
                        reliability_credibility = "; ".join(other_meta_parts)
                        sidc = item.get('milgeo:meta:sidc') or SIDC_BY_COLOR.get(color_en, DEFAULT_SIDC)
                        item_id = str(uuid.uuid4())
                        points_data = item.get('points_data', [])
                        obs_datetime = datetime.fromtimestamp(points_data[0]['ts'], timezone.utc).strftime(
                            '%Y-%m-%dT%H:%M:%S') if points_data and points_data[0].get('ts') else ""
                        wkt_string = f"LINESTRING ({', '.join(f'{p.get("lon", 0.0)} {p.get("lat", 0.0)}' for p in points_data)})" if points_data else ""
                        writer.writerow(
                            [sidc, item_id, "", track_name, obs_datetime, reliability_credibility, staff_comments, "",
                             "", "", wkt_string])
            self._update_status(f"Файли CSV (TRK) успішно збережено.", self.C_ACCENT_DONE)
            return True
        except Exception as e:
            self._update_status(f"Помилка CSV: {e}", error=True); return False

    def create_csv_original_logic(self, contents_list, base_save_path):
        self._update_status(f"Створення CSV: {os.path.basename(base_save_path)}...", self.C_BUTTON_HOVER)
        headers = ['coordinates', 'milgeo:meta:color', 'milgeo:meta:creator', 'milgeo:meta:creator_url',
                   'milgeo:meta:desc', 'milgeo:meta:name', 'name', 'sidc']

        if not contents_list: return False
        try:
            total_items = len(contents_list)
            for chunk_index, i in enumerate(range(0, total_items, self.CSV_CHUNK_SIZE)):
                chunk_contents = contents_list[i:i + self.CSV_CHUNK_SIZE]
                current_save_path = self._get_chunked_save_path(base_save_path, chunk_index)
                with open(current_save_path, "w", encoding="UTF-8", newline='') as f_out:
                    writer = csv.writer(f_out, quoting=csv.QUOTE_MINIMAL)
                    writer.writerow(headers)
                    for item in chunk_contents:
                        geom_type = item.get('geometry_type')
                        if geom_type not in ['LineString', 'Polygon']: continue
                        points_data = item.get('points_data', [])
                        wkt_string = ""
                        if points_data:
                            coords_parts = [f"{p.get('lon', 0.0)} {p.get('lat', 0.0)}" for p in points_data]
                            if geom_type == 'LineString':
                                wkt_string = f"LINESTRING ({', '.join(coords_parts)})"
                            elif geom_type == 'Polygon':
                                if coords_parts and coords_parts[0] != coords_parts[-1]: coords_parts.append(
                                    coords_parts[0])
                                wkt_string = f"POLYGON (({', '.join(coords_parts)}))"

                        def format_meta(v):
                            return json.dumps(v if isinstance(v, list) else [v],
                                              ensure_ascii=False) if v is not None else "[]"

                        row = [wkt_string, item.get('milgeo:meta:color', '')] + [
                            format_meta(item.get(f)) if format_meta(item.get(f)) != '[]' else '' for f in
                            ['milgeo:meta:creator', 'milgeo:meta:creator_url', 'milgeo:meta:desc',
                             'milgeo:meta:name']] + [item.get('name', ''), item.get('milgeo:meta:sidc', '')]
                        writer.writerow(row)
            self._update_status(f"Файли CSV успішно збережено.", self.C_ACCENT_DONE)
            return True
        except Exception as e:
            self._update_status(f"Помилка CSV: {e}", error=True); return False

    def create_geojson(self, contents_list, save_path):
        if not contents_list: return False
        features = []
        for item in contents_list:
            geom_type = item.get('geometry_type')
            props = {k.replace("milgeo:meta:", ""): v for k, v in item.items() if
                     k.startswith("milgeo:meta:") and v is not None}
            props.update({k: v for k, v in item.items() if
                          not k.startswith("milgeo:meta:") and k not in ['points_data', 'geometry_type',
                                                                         'original_location_data'] and v is not None})
            props['color'] = self.convert_color(item.get("color", "White"), "hex", True)
            geometry = None
            if geom_type == 'Point':
                geometry = {"type": "Point", "coordinates": [item.get("lon", 0.0), item.get("lat", 0.0),
                                                             item.get("original_location_data", {}).get("alt",
                                                                                                        0.0) or 0.0]}
            elif geom_type in ['LineString', 'Polygon']:
                coords = [[p.get('lon', 0.0), p.get('lat', 0.0)] for p in item.get('points_data', [])]
                if len(coords) >= (2 if geom_type == 'LineString' else 3):
                    if geom_type == 'Polygon' and coords[0] != coords[-1]: coords.append(coords[0])
                    geometry = {"type": geom_type, "coordinates": [coords] if geom_type == 'Polygon' else coords}
            if geometry: features.append({"type": "Feature", "properties": props, "geometry": geometry})
        if not features: return False
        try:
            with open(save_path, "w", encoding="UTF-8") as f:
                json.dump({"type": "FeatureCollection", "features": features}, f, indent=2,
                          ensure_ascii=False); return True
        except IOError:
            return False

    def _apply_selected_numeration(self, point_list):
        if not point_list: return []
        if len(point_list) == 1: point_list[0]["name"] = self.generate_free_numbers_list(1)[0]; return point_list
        method = self.numerations.index(self.chosen_numeration.get())
        if method == 0: return self.apply_neighbor_numeration(point_list)
        if method == 1: return self.apply_snake_numeration(point_list)
        if method == 2: return self.apply_two_axis_numeration(point_list)
        if method == 3: return self.apply_one_axis_numeration(point_list)
        if method == 4: return self.apply_random_numeration(point_list)
        return point_list

    def generate_free_numbers_list(self, count):
        if count <= 0: return []
        numbers, num, exceptions_on = [], 1, self.exceptions_agree.get()
        while len(numbers) < count:
            if not (exceptions_on and (30 <= num <= 40 or 500 <= num <= 510)):
                numbers.append(str(num))
            num += 1
        return numbers

    def apply_random_numeration(self, content_list):
        if not content_list: return []
        result = copy.deepcopy(content_list)
        free_numbers = self.generate_free_numbers_list(len(result))
        random.shuffle(free_numbers)
        for i, item in enumerate(result): item["name"] = free_numbers[i]
        return result

    def calculate_distance(self, p1, p2):
        R = 6371e3
        lat1, lon1, lat2, lon2 = map(math.radians, [p1[0], p1[1], p2[0], p2[1]])
        dlon, dlat = lon2 - lon1, lat2 - lat1
        a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
        return R * 2 * atan2(sqrt(a), sqrt(1 - a))

    def apply_snake_numeration(self, content_list):
        if not content_list: return content_list
        points = [p for p in content_list if 'lat' in p and 'lon' in p]
        if not points: return content_list

        min_lon, max_lon = min(p['lon'] for p in points), max(p['lon'] for p in points)
        min_lat, max_lat = min(p['lat'] for p in points), max(p['lat'] for p in points)

        # Simple grid logic for snake pattern
        points.sort(key=lambda p: (int((p['lat'] - min_lat) / (max_lat - min_lat) * 10),
                                   p['lon'] if int((p['lat'] - min_lat) / (max_lat - min_lat) * 10) % 2 == 0 else -p[
                                       'lon']))

        free_numbers = self.generate_free_numbers_list(len(points))
        for i, item in enumerate(points):
            item['name'] = free_numbers[i]
        return points

    def apply_neighbor_numeration(self, content_list):
        if not content_list: return []
        points = [p for p in content_list if 'lat' in p and 'lon' in p]
        if not points: return content_list

        unvisited = points[:]
        start_point = min(unvisited, key=lambda p: (p['lat'], p['lon']))
        ordered_points = [start_point]
        unvisited.remove(start_point)

        current_point = start_point
        while unvisited:
            next_point = min(unvisited,
                             key=lambda p: self.calculate_distance((current_point['lat'], current_point['lon']),
                                                                   (p['lat'], p['lon'])))
            ordered_points.append(next_point)
            unvisited.remove(next_point)
            current_point = next_point

        free_numbers = self.generate_free_numbers_list(len(ordered_points))
        for i, item in enumerate(ordered_points):
            item['name'] = free_numbers[i]
        return ordered_points

    def apply_one_axis_numeration(self, content_list):
        if not content_list: return []
        points = [p for p in content_list if 'lat' in p and 'lon' in p]
        if not points: return content_list

        trans = self.chosen_translation.get()
        if trans == "На 90 градусів" or trans == "На 270 градусів":
            key, reverse = 'lat', trans == "На 270 градусів"
        else:
            key, reverse = 'lon', trans == "На 180 градусів"

        points.sort(key=lambda p: p[key], reverse=reverse)
        free_numbers = self.generate_free_numbers_list(len(points))
        for i, item in enumerate(points): item['name'] = free_numbers[i]
        return points

    def apply_two_axis_numeration(self, content_list):
        if not content_list: return []
        points = [p for p in content_list if 'lat' in p and 'lon' in p]
        if not points: return content_list

        min_lat, max_lat = min(p['lat'] for p in points), max(p['lat'] for p in points)
        min_lon, max_lon = min(p['lon'] for p in points), max(p['lon'] for p in points)

        trans = self.chosen_translation.get()
        if trans == "Не повертати":
            corner = (min_lat, min_lon)
        elif trans == "На 90 градусів":
            corner = (min_lat, max_lon)
        elif trans == "На 180 градусів":
            corner = (max_lat, max_lon)
        else:
            corner = (max_lat, min_lon)

        points.sort(key=lambda p: self.calculate_distance(corner, (p['lat'], p['lon'])))
        free_numbers = self.generate_free_numbers_list(len(points))
        for i, item in enumerate(points): item['name'] = free_numbers[i]
        return points


if __name__ == "__main__":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (ImportError, AttributeError, OSError):
        pass
    app = Main()
    app.run()
