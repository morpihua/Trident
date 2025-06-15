# nexus_app.py
# -*- coding: utf-8 -*-

import sys
import struct
import json
from dataclasses import dataclass, field
from typing import List, Dict, Any, BinaryIO

# Використовуємо PyQt6 для сучасного інтерфейсу
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTableWidget, QTableWidgetItem, QTextEdit, QFileDialog,
    QMessageBox, QHeaderView, QLabel
)
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import Qt

# --- СТИЛІЗАЦІЯ ІНТЕРФЕЙСУ (ВІЙСЬКОВА ТЕМАТИКА) ---
MILITARY_STYLESHEET = """
    QWidget {
        background-color: #212121; /* Темний вугільний фон */
        color: #E0E0E0; /* Світло-сірий текст */
        font-family: 'Consolas', 'Lucida Console', 'Courier New', monospace;
        font-size: 14px;
    }
    QMainWindow {
        background-color: #1a1a1a;
    }
    QLabel#title {
        font-size: 20px;
        font-weight: bold;
        color: #00A7E1; /* Яскравий блакитний для заголовків */
    }
    QPushButton {
        background-color: #3C3C3C;
        border: 1px solid #00A7E1;
        padding: 8px;
        min-width: 120px;
    }
    QPushButton:hover {
        background-color: #4A4A4A;
    }
    QPushButton:pressed {
        background-color: #2A2A2A;
    }
    QTextEdit {
        background-color: #1A1A1A;
        border: 1px solid #444;
        font-size: 12px;
        color: #32CD32; /* Зелений, як на старих моніторах */
    }
    QTableWidget {
        background-color: #2C2C2C;
        gridline-color: #444;
    }
    QHeaderView::section {
        background-color: #3C3C3C;
        border: 1px solid #444;
        padding: 4px;
    }
"""

# --- БЕКЕНД: ОНОВЛЕНИЙ МОДУЛЬ ПАРСЕРІВ ---

class Primitives:
    @staticmethod
    def read_byte(file: BinaryIO) -> int:
        return struct.unpack('>b', file.read(1))[0]
    
    @staticmethod
    def read_int(file: BinaryIO) -> int:
        return struct.unpack('>i', file.read(4))[0]

    @staticmethod
    def read_long(file: BinaryIO) -> int:
        return struct.unpack('>q', file.read(8))[0]

    @staticmethod
    def read_string(file: BinaryIO) -> str:
        str_len = Primitives.read_int(file)
        return file.read(str_len).decode('utf-8', errors='ignore') if str_len > 0 else ""

    @staticmethod
    def read_coordinate(file: BinaryIO) -> float:
        return Primitives.read_int(file) / 1e7

@dataclass
class Metadata:
    entries: Dict[str, Any] = field(default_factory=dict)

@dataclass
class Location:
    lon: float
    lat: float
    properties: Dict[str, Any] = field(default_factory=dict)

@dataclass
class Waypoint:
    location: Location
    metadata: Metadata

class Parser:
    @staticmethod
    def parse_metadata_content(file: BinaryIO) -> Dict[str, Any]:
        """
        [ВИПРАВЛЕНО] Повністю переписаний парсер метаданих згідно зі специфікацією.
        Based on psyberia_landmarks_files_specs.pdf, section 1.7.12
        """
        entries = {}
        num_entries = Primitives.read_int(file)
        if num_entries == -1 or num_entries > 1000: # Захист від некоректних файлів
            return {}
        
        for _ in range(num_entries):
            name = Primitives.read_string(file)
            entry_type = Primitives.read_int(file)
            
            data = None
            if entry_type >= 0: # String
                # ***КЛЮЧОВЕ ВИПРАВЛЕННЯ***: довжина рядка - це сам entry_type
                data = file.read(entry_type).decode('utf-8', 'ignore')
            elif entry_type == -1: # Boolean
                data = bool(Primitives.read_byte(file))
            elif entry_type == -2: # Long
                data = Primitives.read_long(file)
            elif entry_type == -3: # Double
                # У специфікації немає double, але додамо про всяк випадок
                data = struct.unpack('>d', file.read(8))[0]
            elif entry_type == -4: # Raw data
                data_size = Primitives.read_int(file)
                data = file.read(data_size)

            if name:
                entries[name] = data
        
        # У сучасних форматах після метаданих йде їхня версія
        if num_entries != -1:
            try:
                # Намагаємося прочитати версію, але не падаємо, якщо її немає (для сумісності)
                _ = Primitives.read_int(file)
            except struct.error:
                pass # Кінець файлу, версії не було
                
        return entries

    @staticmethod
    def parse_location(file: BinaryIO) -> Location:
        """Парсер структури {Location}."""
        # Based on psyberia_landmarks_files_specs.pdf, section 1.7.1
        try:
            struct_size = Primitives.read_int(file)
            lon = Primitives.read_coordinate(file)
            lat = Primitives.read_coordinate(file)
            bytes_read = 8 # lon (4) + lat (4)
            properties = {}
            
            # Читаємо додаткові значення, поки не досягнемо кінця структури
            while bytes_read < struct_size:
                value_type = Primitives.read_byte(file)
                bytes_read += 1
                
                # Based on psyberia_landmarks_files_specs.pdf, section 1.7.2.1
                if value_type == 0x65: # Elevation
                    properties['elevation'] = Primitives.read_int(file) / 1e3
                    bytes_read += 4
                elif value_type == 0x74: # Time
                    properties['time_utc_ms'] = Primitives.read_long(file)
                    bytes_read += 8
                # TODO: Додати інші типи (0x61, 0x70 і т.д.) за потреби
                
            return Location(lon=lon, lat=lat, properties=properties)
        except struct.error:
            return None # Не вдалося прочитати дані

    @staticmethod
    def parse_wpt(file_path: str) -> Waypoint:
        """
        [ВИПРАВЛЕНО] Більш надійний парсер для .wpt, що викликає оновлені функції.
        """
        try:
            with open(file_path, 'rb') as f:
                # Based on psyberia_landmarks_files_specs.pdf, section 1.1
                magic_and_version = Primitives.read_int(f)
                header_size = Primitives.read_int(f)

                # Структура {Waypoint} = {Metadata} + {Location}
                metadata_entries = Parser.parse_metadata_content(f)
                location = Parser.parse_location(f)

                if location:
                    return Waypoint(
                        metadata=Metadata(entries=metadata_entries),
                        location=location
                    )
                return None
        except Exception as e:
            print(f"Error parsing WPT file: {e}")
            return None

class Converter:
    @staticmethod
    def to_geojson_feature(waypoint: Waypoint) -> dict:
        """Перетворює об'єкт Waypoint у словник формату GeoJSON Feature."""
        properties = waypoint.metadata.entries.copy()
        properties.update(waypoint.location.properties)
        
        coordinates = [waypoint.location.lon, waypoint.location.lat]
        if 'elevation' in properties:
            coordinates.append(properties['elevation'])

        return {
            "type": "Feature",
            "geometry": { "type": "Point", "coordinates": coordinates },
            "properties": properties
        }

class NexusApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Nexus Geospatial Converter")
        self.setGeometry(100, 100, 1200, 800)
        
        self.parser = Parser()
        self.converter = Converter()
        self.loaded_data = []

        self.init_ui()

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)

        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_panel.setFixedWidth(300)

        title_label = QLabel("NEXUS")
        title_label.setObjectName("title")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        btn_import = QPushButton("Імпорт файлу")
        btn_import.clicked.connect(self.import_file)
        
        self.btn_export = QPushButton("Експорт в GeoJSON")
        self.btn_export.clicked.connect(self.export_to_geojson)
        self.btn_export.setEnabled(False)

        self.log_widget = QTextEdit()
        self.log_widget.setReadOnly(True)
        self.log_widget.setPlaceholderText("Лог операцій...")

        left_layout.addWidget(title_label)
        left_layout.addWidget(btn_import)
        left_layout.addWidget(self.btn_export)
        left_layout.addWidget(self.log_widget)

        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)

        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(4)
        self.table_widget.setHorizontalHeaderLabels(["Тип", "Назва", "Довгота", "Широта"])
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        
        right_layout.addWidget(self.table_widget)

        main_layout.addWidget(left_panel)
        main_layout.addWidget(right_panel)

        self.log("Програму Nexus запущено.")

    def log(self, message):
        self.log_widget.append(f">> {message}")

    def import_file(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Виберіть файл для імпорту", "", "Waypoint Files (*.wpt);;All Files (*)")
        
        if filepath:
            self.log(f"Імпортування файлу: {filepath}")
            if filepath.lower().endswith(".wpt"):
                data = self.parser.parse_wpt(filepath)
                if data:
                    self.loaded_data = [data]
                    self.display_data()
                    self.log("Файл .wpt успішно розібрано.")
                    self.btn_export.setEnabled(True)
                else:
                    self.log("Помилка: не вдалося розібрати файл.")
                    QMessageBox.critical(self, "Помилка парсингу", "Не вдалося прочитати структуру файлу. Можливо, файл пошкоджений або має непідтримувану версію.")
            else:
                QMessageBox.warning(self, "Помилка", "Цей тип файлу ще не підтримується.")
                self.log(f"Помилка: формат файлу не підтримується.")

    def display_data(self):
        self.table_widget.setRowCount(0)
        for item in self.loaded_data:
            if isinstance(item, Waypoint):
                row_position = self.table_widget.rowCount()
                self.table_widget.insertRow(row_position)
                
                name = item.metadata.entries.get("name", "Без назви")
                lon = item.location.lon
                lat = item.location.lat

                self.table_widget.setItem(row_position, 0, QTableWidgetItem("Waypoint"))
                self.table_widget.setItem(row_position, 1, QTableWidgetItem(str(name)))
                self.table_widget.setItem(row_position, 2, QTableWidgetItem(f"{lon:.6f}"))
                self.table_widget.setItem(row_position, 3, QTableWidgetItem(f"{lat:.6f}"))

    def export_to_geojson(self):
        if not self.loaded_data:
            self.log("Немає даних для експорту.")
            return
        
        filepath, _ = QFileDialog.getSaveFileName(self, "Зберегти як GeoJSON", "", "GeoJSON Files (*.geojson);;All Files (*)")

        if filepath:
            self.log(f"Експортування в: {filepath}")
            features = [self.converter.to_geojson_feature(item) for item in self.loaded_data if isinstance(item, Waypoint)]
            
            geojson_data = {"type": "FeatureCollection", "features": features}
            
            try:
                with open(filepath, 'w', encoding='utf-8') as f:
                    json.dump(geojson_data, f, ensure_ascii=False, indent=2)
                self.log("Експорт успішно завершено.")
            except Exception as e:
                self.log(f"Помилка експорту: {e}")
                QMessageBox.critical(self, "Помилка", f"Не вдалося зберегти файл:\n{e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(MILITARY_STYLESHEET)
    window = NexusApp()
    window.show()
    sys.exit(app.exec())