import os
import struct
import time
import base64
import sys
from typing import Optional, Callable

# Додаємо кольори для консолі (мінімальний варіант)
COLORS_CONSOLE = {
    'E': '\033[91m',  # Червоний
    'W': '\033[93m',  # Жовтий
    'P': '\033[0m',   # Звичайний
    'D': '\033[90m',  # Сірий
    'T': '\033[95m',  # Фіолетовий
    'o': '\033[0m',   # Скидання
}

class Base:
    def __init__(self, verbosity: int = 0, gui_logger_func: Optional[Callable] = None):
        self.verbosity = verbosity
        self.gui_logger_func = gui_logger_func

    def _log(self, level_char, message, *args):
        level_map_verbosity = {'E': -2, 'W': -1, 'P': 0, 'D': 1, 'T': 2}
        if self.verbosity < level_map_verbosity.get(level_char, 0):
            return

        prefix_map = {'E': 'ПОМИЛКА APQ: ', 'W': 'УВАГА APQ: ', 'P': 'APQ: ', 'D': 'НАЛАГОДЖЕННЯ APQ: ', 'T': 'TRACE APQ: '}
        log_message = prefix_map.get(level_char, '?') + (message % args if args else message)

        if self.gui_logger_func:
            is_error = level_char == 'E'
            is_warning = level_char == 'W'
            try:
                self.gui_logger_func(log_message, error=is_error, warning=is_warning)
            except TypeError:
                self.gui_logger_func(log_message)

        print(COLORS_CONSOLE.get(level_char, '') + log_message + COLORS_CONSOLE.get('o', ''), file=sys.stderr)

    def error(self, message, *args):
        self._log('E', message, *args)

    def warning(self, message, *args):
        self._log('W', message, *args)

    def print(self, message, *args):
        self._log('P', message, *args)

    def debug(self, message, *args):
        self._log('D', message, *args)

    def trace(self, message, *args):
        self._log('T', message, *args)

    def trace_hexdump(self, data_bytes):
        if self.verbosity >= 2:
            data_size = len(data_bytes)
            for offs in range(0, data_size, 16):
                s_bytes = data_bytes[offs:offs + 16]
                hex_parts = []
                for i in range(16):
                    if (i % 8) == 0 and i > 0: hex_parts.append(' ')
                    if i < len(s_bytes):
                        hex_parts.append(f'{s_bytes[i]:02x}')
                    else:
                        hex_parts.append('  ')
                hex_str = ' '.join(hex_parts)
                ascii_str = "".join(chr(b) if 32 <= b <= 126 else '.' for b in s_bytes)
                self.trace(f'0x{offs:08x} {hex_str:<47} |{ascii_str:<16}|')

    def _load_raw(self, path_to_load):
        if not os.path.isfile(path_to_load) or not os.access(path_to_load, os.R_OK):
            self.warning("Неможливо прочитати '%s'!", path_to_load)
            return None
        try:
            with open(path_to_load, 'rb') as f_in:
                raw_data_content = f_in.read()
            self.trace("Прочитано '%s': %d байт.", path_to_load, len(raw_data_content))
            return raw_data_content
        except IOError as e_io:
            self.warning("Помилка читання '%s': %s", path_to_load, e_io)
            return None

class ApqFile(Base):
    # ... (залишаємо все як було, додаємо type hints)
    
    class ApqFile(Base):
    MAX_REASONABLE_STRING_LEN = 65536 * 2
    MAX_REASONABLE_ENTRIES = 100000

    def __init__(self, path=None, rawdata=None, file_type=None, rawname=None, rawts=None, verbosity=0,
                 gui_logger_func=None):
        super().__init__(verbosity, gui_logger_func)
        self.path = path
        self.rawdata = rawdata
        self._file_type = file_type.lower() if file_type else None  # Ensure lowercase
        self.rawname = rawname
        self.rawts = rawts if rawts is not None else time.time()

        self.data_parsed = {}
        self.version = 0
        self.rawoffs = 0
        self.parse_successful = False
        load_success = False

        if self.path:
            self.trace('new(path => %s)', self.path)
            file_name_local = os.path.basename(self.path)  # Renamed to avoid conflict with class member
            ext_match = file_name_local.lower().split('.')[-1] if '.' in file_name_local else ''
            aq_types = ["wpt", "set", "rte", "are", "trk", "ldk"]
            if ext_match in aq_types:
                self._file_type = ext_match
                self.rawdata = self._load_raw(self.path)
                if self.rawdata is None:
                    self.error(f"Не вдалося завантажити дані для {self.path}")
                    raise ValueError(f"Не вдалося завантажити {self.path}")
                try:
                    self.rawts = os.path.getmtime(self.path)
                except OSError:
                    self.rawts = time.time()
                load_success = True
            else:
                self.error("Невідомий тип файлу для шляху: %s!", self.path)
                raise ValueError(f"Unknown file type for {self.path}")
        elif self.rawdata is not None and self._file_type and self.rawname:
            valid_raw_types = ["wpt", "set", "rte", "are", "trk", "bin", "ldk"]
            if self._file_type not in valid_raw_types:
                self.error("Невідомий тип файлу: %s!", self._file_type)
                raise ValueError(f"Unknown raw type: {self._file_type}")
            self.path = self.rawname
            if self.rawts is None: self.rawts = time.time()
            load_success = True
        else:
            self.error("Неправильні параметри ApqFile!")
            raise ValueError("Illegal ApqFile params")

        if not load_success or self.rawdata is None:
            self.error("Дані не завантажено або відсутні для ApqFile.")
            return

        self.rawsize = len(self.rawdata)
        if self.verbosity >= 3: self.trace_hexdump(self.rawdata)

        parser_method_name = f"_parse_{self._file_type}"

        if hasattr(self, parser_method_name) and callable(getattr(self, parser_method_name)):
            try:
                self.parse_successful = getattr(self, parser_method_name)()
            except Exception as e_parse:  # Catch exceptions during parsing
                self.error(f"Виняток під час парсингу {self._file_type} ({self.path or self.rawname}): {e_parse}")
                self.parse_successful = False
        else:
            self.warning(f"Парсер для типу не знайдено: {self._file_type}")
            # Handle 'bin' specifically for raw data passthrough
            if self._file_type == "bin":
                self.data_parsed['raw_content_b64'] = base64.b64encode(self.rawdata).decode('ascii')
                self.parse_successful = True

        aq_types_for_check = ["wpt", "set", "rte", "are", "trk", "ldk"]
        if self.parse_successful and self.rawoffs != self.rawsize and self._file_type in aq_types_for_check:
            remaining_bytes = self.rawsize - self.rawoffs
            # Adjusted condition for logging remaining bytes
            if remaining_bytes > 0 and (
                    self.rawsize < 32 or remaining_bytes > 8 or remaining_bytes >= self.rawsize * 0.01):
                self.debug('Залишились невикористані дані: %d байт з %d (0x%04x з 0x%04x).', remaining_bytes,
                           self.rawsize, self.rawoffs, self.rawsize)

        if not self.parse_successful and self._file_type in aq_types_for_check:
            self.error("Помилка парсингу даних для %s (тип: %s)!", self.path or self.rawname or "невідомий файл",
                       self._file_type or "невідомий тип")

    def type(self):
        return self._file_type

    def data(self):
        return self.get_parsed_data()

    def _tell(self):
        return self.rawoffs

    def _seek(self, offset):
        self.rawoffs = offset;
        return self.rawoffs

    def _size(self):
        return self.rawsize

    def _getval(self, val_type, arg=None):
        original_offset = self.rawoffs
        value = None
        raw_bytes_read = b''

        type_map_struct = {
            'int': ('>i', 4), 'bool': ('>?', 1), 'byte': ('>b', 1), 'ubyte': ('>B', 1),
            'long': ('>q', 8), 'pointer': ('>Q', 8), 'double': ('>d', 8),
            'short': ('>h', 2), 'ushort': ('>H', 2)
        }

        if val_type in type_map_struct:
            struct_format, num_bytes = type_map_struct[val_type]
            if self.rawoffs + num_bytes > self.rawsize:
                self.debug(
                    f"Недостатньо даних для '{val_type}' на 0x{self.rawoffs:X} (потрібно {num_bytes}, є {self.rawsize - self.rawoffs})")
                return None
            raw_bytes_read = self.rawdata[self.rawoffs: self.rawoffs + num_bytes]
            try:
                value = struct.unpack(struct_format, raw_bytes_read)[0]
            except struct.error as e:
                self.error(f"Помилка розпаковки '{val_type}' на 0x{original_offset:X}: {e}");
                return None
            self.rawoffs += num_bytes
            if val_type == 'bool': value = bool(value)
        elif val_type == 'int+raw':
            size_val = self._getval('int')
            if size_val is None or size_val < 0 or size_val > self.MAX_REASONABLE_STRING_LEN * 10:
                self.error(f"Некоректний або завеликий розмір ({size_val}) для int+raw на 0x{original_offset:X}");
                return None
            if self.rawoffs + size_val > self.rawsize:
                self.debug(f"Недостатньо даних для int+raw (розмір {size_val})")
                return None
            raw_bytes_read = self.rawdata[self.rawoffs: self.rawoffs + size_val]
            value = base64.b64encode(raw_bytes_read).decode('ascii')
            self.rawoffs += size_val
        elif val_type == 'raw' or val_type == 'bin':
            size = arg
            if size is None or size < 0 or size > self.MAX_REASONABLE_STRING_LEN * 100:
                self.error(f"Некоректний або завеликий розмір ({size}) для '{val_type}'");
                return None
            if self.rawoffs + size > self.rawsize:
                self.debug(f"Недостатньо даних для '{val_type}' (розмір {size})")
                return None
            raw_bytes_read = self.rawdata[self.rawoffs: self.rawoffs + size]
            value = base64.b64encode(raw_bytes_read).decode('ascii') if val_type == 'raw' else raw_bytes_read
            self.rawoffs += size
        elif val_type == 'string':
            size = arg
            if size is None or size < 0 or size > self.MAX_REASONABLE_STRING_LEN:
                self.error(f"Некоректний або завеликий розмір ({size}) для string на 0x{original_offset:X}");
                return None
            if self.rawoffs + size > self.rawsize:
                self.debug(f"Недостатньо даних для string (size {size})")
                return None
            raw_bytes_read = self.rawdata[self.rawoffs: self.rawoffs + size]
            try:
                value = raw_bytes_read.decode('utf-8')
            except UnicodeDecodeError:
                self.warning(f"Помилка UTF-8 для рядка на 0x{original_offset:X}, вик. replace");
                value = raw_bytes_read.decode('utf-8', errors='replace')
            self.rawoffs += size
        elif val_type == 'coords':  # Renamed from 'coordinates' for clarity in parser
            int_val = self._getval('int');
            value = int_val * 1e-7 if int_val is not None else None
        elif val_type == 'height':
            int_val = self._getval('int');
            value = (None if int_val == -999999999 else int_val * 1e-3) if int_val is not None else None
        elif val_type == 'timestamp':
            long_val = self._getval('long');
            value = (None if long_val == 0 else long_val * 1e-3) if long_val is not None else None
        elif val_type == 'accuracy':
            int_val = self._getval('int');
            value = (None if int_val == 0 else int_val) if int_val is not None else None
        elif val_type == 'accuracy2':  # Used for some accuracy fields
            int_val = self._getval('int');
            value = (None if int_val == 0 else int_val * 1e-2) if int_val is not None else None
        elif val_type == 'pressure':
            int_val = self._getval('int');
            value = (None if int_val == 999999999 else int_val * 1e-3) if int_val is not None else None
        else:
            self.warning("Невідомий тип '%s' для _getval!", val_type);
            return None

        if self.verbosity >= 2 and value is not None:
            display_val = value
            if val_type in ['raw', 'bin', 'int+raw']:
                # Handle bytes for length display, decode for b64 string
                actual_bytes_len = len(value if isinstance(value, bytes) else base64.b64decode(value))
                display_val = f"<bytes len={actual_bytes_len}>"
            elif isinstance(value, str) and len(value) > 40:
                display_val = value[:37] + "..."
            # Make hex_bytes_str more robust
            hex_bytes_str = ' '.join(f'{b:02x}' for b in raw_bytes_read) if raw_bytes_read else "''"
            self.trace('%-11s at 0x%05x [%02d] %-23s = %s', val_type, original_offset, len(raw_bytes_read),
                       hex_bytes_str, display_val)
        return value

    def _getvalmulti(self, **kwargs_types):
        data_dict = {'_order': list(kwargs_types.keys())}
        first_val_offset = self._tell()
        all_none = True  # Track if all attempts to read return None
        for key, type_info in kwargs_types.items():
            val = None
            arg_for_getval = None
            type_name_for_getval = type_info

            if isinstance(type_info, tuple):
                type_name_for_getval = type_info[0]
                if len(type_info) > 1:
                    arg_for_getval = type_info[1]

            val = self._getval(type_name_for_getval, arg_for_getval)
            data_dict[key] = val

            if val is not None:
                all_none = False  # At least one value was read

            # Adjusted warning condition: warn if a critical field is None
            critical_fields = ['magic', 'offset', 'uid', 'size', 'metaOffset', 'rootOffset', 'nTotal', 'nChild',
                               'nData']
            if val is None and key in critical_fields:
                self.warning(
                    f"_getvalmulti отримав None для критичного поля '{key}' (тип '{type_name_for_getval}') на зсуві 0x{first_val_offset:X}")

        # If all_none, it might indicate an issue with the data stream or struct definition
        if all_none and first_val_offset < self._size() - 8:  # Don't log if at the very end of file
            self.debug(f"_getvalmulti: Усі поля повернули None, починаючи з 0x{first_val_offset:X}")

        if self.verbosity >= 1: self.debug('MultiRead: %s', ', '.join(
            [f"{k}={repr(data_dict.get(k, '<FAIL>'))}" for k in data_dict['_order']]))
        return data_dict

    def _check_header(self, *expected_file_versions):
        file_version = self._getval('int')
        if file_version is None: self.error("Не вдалося прочитати версію файлу."); return None
        if (file_version & 0x50500000) == 0x50500000: file_version = (file_version & 0xff) + 100
        header_size = self._getval('int')
        if header_size is None: self.error("Не вдалося прочитати розмір заголовка."); return None
        if header_size < 0 or header_size > self.rawsize or header_size > 1024:  # Added sanity check for header_size
            self.error(
                f"Некоректний розмір заголовка: {header_size} (0x{header_size:X}). Загальний розмір файлу: {self.rawsize}");
            return None
        self.debug('fileVersion=%s headerSize=0x%X (%d)', file_version, header_size, header_size)
        if expected_file_versions and file_version not in expected_file_versions:
            self.warning('Неочікувана версія файлу %s (очікувалось %s).', file_version,
                         ' або '.join(map(str, expected_file_versions)) if expected_file_versions else "будь-яка")
        self.version = file_version
        return header_size

    def _get_metadata(self):
        metadata_version = 1
        if self.version > 100:
            metadata_version = 3
        elif self._file_type == 'trk' and self.version >= 3:  # Corrected check for TRK v3+
            metadata_version = 2
        elif self._file_type != 'trk' and self.version == 2:  # For WPT, SET, RTE v2
            metadata_version = 2
        # else metadata_version remains 1 (default for older versions)

        n_meta_entries = self._getval('int')
        if n_meta_entries is None: self.error("Не вдалося прочитати nMetaEntries."); return None

        # Handle n_meta_entries == 0 explicitly as valid (no entries)
        if n_meta_entries == 0:
            self.debug('nMetaEntries=0, metadataVersion=%s. Немає записів метаданих.', metadata_version)
            meta = {'_order': [], '_types': {}}
        elif n_meta_entries < -1 or n_meta_entries > self.MAX_REASONABLE_ENTRIES:
            self.error(f"Некоректна кількість записів метаданих: {n_meta_entries} на 0x{self._tell() - 4:X}.")
            if self.verbosity >= 2: self.trace_hexdump(
                self.rawdata[max(0, self._tell() - 12):min(self.rawsize, self._tell() + 8)])
            return None
        else:  # n_meta_entries > 0 or n_meta_entries == -1
            self.debug('nMetaEntries=%d metadataVersion=%s', n_meta_entries, metadata_version)
            meta = {'_order': [], '_types': {}}
            if n_meta_entries != -1:  # If -1, it means an empty list of meta, not an error
                for i in range(n_meta_entries):
                    name_len = self._getval('int')
                    if name_len is None or name_len < 0 or name_len > self.MAX_REASONABLE_STRING_LEN:
                        self.error(f"Некоректна довжина імені ({name_len}) у мета, запис {i}.");
                        return None
                    name_str = self._getval('string', name_len)
                    if name_str is None: self.error(f"Не вдалося прочитати ім'я мета, запис {i}."); return None

                    data_len_or_type = self._getval('int')
                    if data_len_or_type is None: self.error(
                        f"Не вдалося прочитати тип/довжину для '{name_str}'."); return None

                    data_value, data_type_str = None, None
                    type_mapping = {-1: 'bool', -2: 'long', -3: 'double', -4: 'int+raw'}

                    if data_len_or_type in type_mapping:
                        data_type_str = type_mapping[data_len_or_type];
                        data_value = self._getval(data_type_str)
                    elif data_len_or_type >= 0:  # String type
                        if data_len_or_type > self.MAX_REASONABLE_STRING_LEN:
                            self.error(
                                f"Завелика довжина рядка ({data_len_or_type}) для мета '{name_str}' на 0x{self._tell() - 4:X}.")
                            # Additional check for known numeric fields misinterpreted as strings
                            if name_str.lower() in ["lat", "lon", "latitude", "longitude", "altitude", "ele", "east",
                                                    "north", "dte"]:
                                self.error(
                                    f"Поле '{name_str}' (ймовірно числове) не повинно мати такий великий строковий тип/довжину ({data_len_or_type}). Можливо, дані пошкоджено або невірний тип у файлі.")
                            return None  # Critical error if string length is unreasonable
                        data_type_str = 'string';
                        data_value = self._getval(data_type_str, data_len_or_type)
                    else:
                        self.warning('Невідомий тип/довжина мета %d (%s) для "%s" на 0x%X.', data_len_or_type,
                                     hex(data_len_or_type), name_str, self._tell() - 4);
                        return None
                    meta[name_str] = data_value;
                    meta['_order'].append(name_str);
                    meta['_types'][name_str] = data_type_str

        if metadata_version == 3 and n_meta_entries >= 0:  # Read trailing int for v3 meta
            _ = self._getval('int')

        if metadata_version >= 2:
            n_meta_ext = self._getval('int')
            if n_meta_ext is None: self.error("Не вдалося прочитати nMetaExt."); return None
            self.debug('nMetaExt=%d', n_meta_ext)
            if n_meta_ext > 0:
                self.warning("Розширені метадані (%d) не реалізовано. Парсинг метаданих зупинено.", n_meta_ext);
                # Potentially skip these bytes if format is known, for now, we stop.
                return meta  # Return what we have so far, as the main meta might be valid
            elif n_meta_ext < -1:  # -1 is valid (no ext meta), other negative are errors
                self.error(f"Некоректна nMetaExt: {n_meta_ext}");
                return None

        if self.verbosity >= 1 and meta.get('_order'):  # Ensure meta['_order'] exists
            self.debug("--- Metadata ---");
            for ix, k_meta in enumerate(meta['_order']):
                self.debug(' %2d: %-20s (%-7s) = %s', ix + 1, k_meta, meta['_types'].get(k_meta, 'N/A'),
                           repr(meta.get(k_meta)))
            self.debug("--- Кінець Meta ---")
        return meta

    def _get_location(self):
        location_version = 2 if self.version > 100 else 1
        loc = {'lat': None, 'lon': None, 'alt': None, 'ts': None, 'acc': None, 'bar': None, 'batt': None, 'acc_v': None,
               'cell': {'gen': None, 'prot': None, 'sig': None},
               'numsv': {'tot': 0, 'unkn': None, 'G': None, 'S': None, 'R': None, 'J': None, 'C': None, 'E': None,
                         'I': None}}
        loc_start_offs = self._tell()

        # Peek at struct_size first
        if self.rawoffs + 4 > self.rawsize:
            self.error(f"Недостатньо даних для читання struct_size на 0x{loc_start_offs:X}");
            return None
        struct_size_bytes = self.rawdata[self.rawoffs: self.rawoffs + 4]
        struct_size = struct.unpack('>i', struct_size_bytes)[0]

        # Now validate struct_size before reading it with _getval (which advances offset)
        if struct_size < 8 or struct_size > 256:  # Basic sanity check for Location struct size
            self.error(f"Некоректний або завеликий struct_size ({struct_size}) для Location на 0x{loc_start_offs:X}.")
            return None

        # Check if the full struct would exceed file bounds
        if loc_start_offs + 4 + struct_size > self.rawsize:
            self.error(
                f"struct_size ({struct_size}) виходить за межі файлу (0x{loc_start_offs + 4 + struct_size:X} > 0x{self.rawsize:X}) на 0x{loc_start_offs:X}.")
            return None

        struct_size_val_check = self._getval('int')  # This reads and advances offset
        if struct_size_val_check != struct_size:  # Should match the peeked value
            self.warning(
                f"Прочитаний struct_size ({struct_size_val_check}) не збігається з попередньо переглянутим ({struct_size}) на 0x{loc_start_offs:X}")
            # Potentially adjust or error out, for now, proceed with struct_size_val_check if it seems more plausible
            if not (8 <= struct_size_val_check <= 256):  # If the read value is also bad
                self.error("Обидва значення struct_size (попередньо переглянуте та прочитане) невалідні.")
                return None
            struct_size = struct_size_val_check  # Trust the read value more if it's within plausible range

        loc['lon'] = self._getval('coords');
        loc['lat'] = self._getval('coords')
        if loc['lon'] is None or loc['lat'] is None:
            self.error("Не вдалося прочитати lon/lat у Location.");
            return None

        expected_payload_end_pos = loc_start_offs + 4 + struct_size

        if location_version == 1:
            if self._tell() + 4 <= expected_payload_end_pos: loc['alt'] = self._getval('height')
            if self._tell() + 8 <= expected_payload_end_pos: loc['ts'] = self._getval('timestamp')
            if self._tell() + 4 <= expected_payload_end_pos: loc['acc'] = self._getval('accuracy')
            if self._tell() + 4 <= expected_payload_end_pos: loc['bar'] = self._getval('pressure')
        else:  # Location version 2 (fields identified by type byte)
            while self._tell() < expected_payload_end_pos:
                if self.rawoffs + 1 > self.rawsize: self.debug(
                    "Location v2: Кінець даних при читанні field_type."); break
                field_type_byte_val = self.rawdata[self.rawoffs: self.rawoffs + 1]

                self.rawoffs += 1
                field_type = struct.unpack('>b', field_type_byte_val)[0]

                # Check remaining bytes before attempting to read value
                bytes_needed = 0
                if field_type == 0x61:
                    bytes_needed = 4  # accuracy2 (int)
                elif field_type == 0x65:
                    bytes_needed = 4  # height (int)
                elif field_type == 0x70:
                    bytes_needed = 4  # pressure (int)
                elif field_type == 0x74:
                    bytes_needed = 8  # timestamp (long)
                elif field_type == 0x62:
                    bytes_needed = 1  # batt (byte)
                elif field_type == 0x6e:
                    bytes_needed = 2  # cell gen_prot (byte) + sig (byte)
                elif field_type == 0x73:
                    bytes_needed = 8  # numsv (8 bytes)
                elif field_type == 0x76:
                    bytes_needed = 4  # acc_v (int)
                else:
                    self.warning(
                        f"Невідомий тип поля 0x{field_type:02X} у Location v2 на 0x{self._tell() - 1:X}. Пропускаємо решту Location.")
                    break  # Stop parsing this Location on unknown type

                if self._tell() + bytes_needed > expected_payload_end_pos:
                    self.debug(
                        f"Location v2: Недостатньо даних для поля 0x{field_type:02X}. Очікувалось {bytes_needed}, залишилось {expected_payload_end_pos - self._tell()}.")
                    break

                if field_type == 0x61:
                    loc['acc'] = self._getval('accuracy2');
                elif field_type == 0x65:
                    loc['alt'] = self._getval('height');
                elif field_type == 0x70:
                    loc['bar'] = self._getval('pressure');
                elif field_type == 0x74:
                    loc['ts'] = self._getval('timestamp');
                elif field_type == 0x62:
                    loc['batt'] = self._getval('byte');
                elif field_type == 0x6e:
                    gen_prot = self._getval('byte');
                    loc['cell']['sig'] = self._getval('byte');
                    if gen_prot is not None:
                        # Ensure gen_prot is treated as integer for divmod
                        gen_val, prot_val = divmod(int(gen_prot), 10)
                        loc['cell']['gen'], loc['cell']['prot'] = gen_val, prot_val
                elif field_type == 0x73:
                    sats_k = ['unkn', 'G', 'S', 'R', 'J', 'C', 'E', 'I'];
                    total_s = 0;
                    valid_s = False
                    for sk in sats_k:
                        v_s = self._getval('byte')
                        if v_s is None: valid_s = False; break
                        loc['numsv'][sk] = v_s
                        if isinstance(v_s, (int, float)): total_s += v_s  # Ensure it's a number
                        valid_s = True
                    if valid_s: loc['numsv']['tot'] = total_s
                elif field_type == 0x76:
                    loc['acc_v'] = self._getval('accuracy2');

        if self._tell() != expected_payload_end_pos:
            self.debug(
                f"Location: зсув після читання (0x{self._tell():X}) не збігається з очікуваним кінцем (0x{expected_payload_end_pos:X}) для struct_size={struct_size}. Коригування.")
            self._seek(expected_payload_end_pos)

        self.debug('Loc: lon=%.6f, lat=%.6f, alt=%s, ts=%s', loc.get('lon', 0.0), loc.get('lat', 0.0),
                   loc.get('alt', '-'), loc.get('ts', '-'))
        return loc

    def _get_waypoints(self):
        wp_list = []
        n_wp = self._getval('int')
        if n_wp is None or n_wp < 0 or n_wp > self.MAX_REASONABLE_ENTRIES:
            self.error(f"Некоректна або занадто велика к-ть Waypoints: {n_wp} на 0x{self._tell() - 4:X}.")
            if self.verbosity >= 2 and n_wp is not None:
                self.trace_hexdump(self.rawdata[max(0, self._tell() - 12):min(self.rawsize, self._tell() + 8)])
            return None
        self.debug('nWaypoints=%s', n_wp)
        for i in range(n_wp):
            # Pass the global meta of the current file (e.g., SET or RTE file's global meta)
            # to be combined with individual waypoint's meta.
            # For this, ApqFile's global meta needs to be available or passed down.
            # Assuming self.data_parsed['meta'] is not yet populated here, but self.version and self._file_type are.
            # We will rely on _create_point_dict inside _normalize_apq_data to handle meta merging correctly.
            m = self._get_metadata()
            l = self._get_location()
            if m is None or l is None: self.error(f"Помилка парсингу waypoint {i + 1}."); return None
            wp_list.append({'meta': m, 'location': l})
        return wp_list

    def _get_locations(self):  # For ARE files
        loc_list = []
        n_loc = self._getval('int')
        if n_loc is None or n_loc < 0 or n_loc > self.MAX_REASONABLE_ENTRIES * 10:  # Higher limit for ARE
            self.error(f"Некоректна або занадто велика к-ть Locations (ARE): {n_loc}.");
            return None
        self.debug('nLocations=%s (ARE)', n_loc)
        for i in range(n_loc):
            l = self._get_location()
            if l is None: self.error(f"Помилка парсингу location {i + 1} (ARE)."); return None
            loc_list.append(l)
        return loc_list

    def _get_segment(self):  # For TRK files
        seg_ver = 2 if self._file_type == 'trk' and self.version >= 3 else 1
        seg_meta = self._get_metadata() if seg_ver == 2 else {}
        if seg_meta is None and seg_ver == 2:  # Metadata could be legitimately empty
            self.debug("Метадані для Segment v2 не прочитано (можливо, порожні).");
            seg_meta = {}

        n_loc_seg = self._getval('int')
        if n_loc_seg is None or n_loc_seg < 0 or n_loc_seg > self.MAX_REASONABLE_ENTRIES * 100:  # Very high limit for track segments
            self.error(f"Некоректна або занадто велика к-ть locations у Segment: {n_loc_seg}.");
            return None
        self.debug('nLocations in segment=%s, segVer=%s', n_loc_seg, seg_ver)
        locs_in_seg = []
        for i in range(n_loc_seg):
            l = self._get_location()
            if l is None: self.error(f"Помилка парсингу location {i + 1} у Segment."); return None
            locs_in_seg.append(l)
        return {'meta': seg_meta, 'locations': locs_in_seg}

    def _get_segments(self):  # For TRK files
        seg_list = []
        n_seg = self._getval('int')
        if n_seg is None or n_seg < 0 or n_seg > self.MAX_REASONABLE_ENTRIES:  # Max segments for a track
            self.error(f"Некоректна або занадто велика к-ть Segments: {n_seg} на 0x{self._tell() - 4:X}.")
            if self.verbosity >= 2 and n_seg is not None:
                self.trace_hexdump(self.rawdata[max(0, self._tell() - 12):min(self.rawsize, self._tell() + 8)])
            return None
        self.debug('nSegments=%s', n_seg)
        for i in range(n_seg):
            seg_data = self._get_segment()
            if seg_data is None: self.error(f"Помилка парсингу segment {i + 1}."); return None
            seg_list.append(seg_data)
        return seg_list

    def _parse_wpt(self):
        self.debug(f"Розбір WPT: {self.path or self.rawname}")
        h_size = self._check_header(2, 101);  # WPT v2 or v101
        if h_size is None: return False
        # Header in WPT v101 seems to contain main metadata, then another metadata block for the point itself.
        # In v2, header is simpler.
        if self.version > 100:  # v101
            self.data_parsed['meta'] = self._get_metadata()  # глобальні метадані
            self.data_parsed['location'] = self._get_location()  # location одразу після meta
            return bool(self.data_parsed.get('meta') is not None and self.data_parsed.get('location') is not None)

    def _parse_set(self):
        self.debug(f"Розбір SET: {self.path or self.rawname}")
        h_size = self._check_header(2, 101)  # SET v2 or v101
        if h_size is None: return False
        if self.version < 100:  # v2
            if h_size > 0: self._seek(self._tell() + h_size)
        else:  # v101
            _ = self._get_metadata()  # File-level/creator meta
            # header_size for v101 should point after this first meta block
            # No explicit seek based on h_size needed here as it's handled by _get_metadata logic for v3 meta
        self.data_parsed['meta'] = self._get_metadata()  # This is the SET's own metadata
        self.data_parsed['waypoints'] = self._get_waypoints()
        return bool(self.data_parsed.get('meta') is not None and self.data_parsed.get('waypoints') is not None)

    def _parse_rte(self):
        self.debug(f"Розбір RTE: {self.path or self.rawname}")
        h_size = self._check_header(2, 101)  # RTE v2 or v101
        if h_size is None: return False
        if self.version < 100:  # v2
            if h_size > 0: self._seek(self._tell() + h_size)
        else:  # v101
            _ = self._get_metadata()  # File-level/creator meta
        self.data_parsed['meta'] = self._get_metadata()  # RTE's own metadata
        self.data_parsed['waypoints'] = self._get_waypoints()
        return bool(self.data_parsed.get('meta') is not None and self.data_parsed.get('waypoints') is not None)

    def _parse_are(self):
        self.debug(f"Синтаксичний аналіз ARE: {self.path or self.rawname}")  # Corrected "Синтаксичний аналіз:"
        h_size = self._check_header(2)  # ARE only has v2
        if h_size is None: return False
        if h_size > 0: self._seek(self._tell() + h_size)
        self.data_parsed['meta'] = self._get_metadata()
        self.data_parsed['locations'] = self._get_locations()
        return bool(self.data_parsed.get('meta') is not None and self.data_parsed.get('locations') is not None)

    def _parse_trk(self):
        self.debug(f"Розбір TRK: {self.path or self.rawname}")
        h_size = self._check_header(2, 3, 101)  # TRK v2, v3 or v101
        if h_size is None: return False
        if self.version < 100:  # v2 or v3
            if h_size > 0: self._seek(self._tell() + h_size)
        else:  # v101
            _ = self._get_metadata()  # File-level/creator meta
            # header_size for v101 should point after this first meta block
        self.data_parsed['meta'] = self._get_metadata()  # Track's own metadata
        if self.data_parsed.get('meta') is None:
            self.error("Не вдалося розпарсити головні метадані треку.");
            return False
        self.data_parsed['waypoints'] = self._get_waypoints()  # POIs associated with the track
        if self.data_parsed.get('waypoints') is None:
            self.error("Не вдалося розпарсити шляхові точки треку.");
            return False
        self.data_parsed['segments'] = self._get_segments()
        if self.data_parsed.get('segments') is None:
            self.error("Не вдалося розпарсити сегменти треку.");
            return False
        return True

    def _parse_ldk(self):
        self.debug(f"Parsing LDK: {self.path or self.rawname}")
        hdr = self._getvalmulti(magic='int', archVersion='int', rootOffset='pointer',
                                res1='long', res2='long', res3='long', res4='long')
        if None in [hdr.get('magic'), hdr.get('archVersion'), hdr.get('rootOffset')]:
            self.error("Не вдалося прочитати заголовок LDK (обов'язкові поля).");
            return False
        if hdr.get('magic') != 0x4c444b3a:  # "LDK:"
            self.warning('Невідомий LDK magic 0x%08x.', hdr.get('magic'));
            return False
        if hdr.get('archVersion') != 1:
            self.warning('Невідома версія архіву LDK %d.', hdr.get('archVersion'));
            return False

        root_offset_val = hdr.get('rootOffset')
        if root_offset_val == 0 or root_offset_val >= self.rawsize:  # Check for invalid offset
            self.error(f"Некоректний rootOffset LDK: {root_offset_val}");
            return False
        self.data_parsed['root'] = self._get_node(root_offset_val)
        return self.data_parsed.get('root') is not None

    def _get_node_data(self, initial_offset):
        self._seek(initial_offset)
        hdr = self._getvalmulti(magic='int', flags='int', totalSize='long',
                                size='long', addOffset='pointer')
        if None in [hdr.get('magic'), hdr.get('size'), hdr.get('addOffset')]:
            self.error("Не вдалося прочитати заголовок даних вузла LDK (обов'язкові поля).");
            return None
        if hdr.get('magic') != 0x00105555:
            self.warning('Неправильний LDK data magic 0x%08x.', hdr.get('magic'));
            return None

        main_data_size_val = hdr.get('size')
        if main_data_size_val < 0 or main_data_size_val > self.rawsize:  # Sanity check
            self.error(f"Некоректний розмір основного блоку даних LDK: {main_data_size_val}");
            return None
        data_chunks = []
        main_data_block = self._getval('bin', main_data_size_val)
        if main_data_block is None: self.error("Не вдалося прочитати основний блок даних LDK."); return None
        data_chunks.append(main_data_block)

        current_add_offset_val = hdr.get('addOffset')
        safety_counter = 0  # Prevent infinite loops from corrupted addOffset
        while current_add_offset_val != 0 and current_add_offset_val is not None and safety_counter < 100:
            safety_counter += 1
            if current_add_offset_val >= self.rawsize:  # Offset out of bounds
                self.error(f"Некоректний addOffset LDK: {current_add_offset_val}");
                return None
            self._seek(current_add_offset_val)
            add_hdr = self._getvalmulti(magic='int', size='long', addOffset='pointer')
            if None in [add_hdr.get('magic'), add_hdr.get('size'), add_hdr.get('addOffset')]:
                self.error("Не вдалося прочитати заголовок дод. блоку LDK (обов'язкові поля).");
                return None
            if add_hdr.get('magic') != 0x00205555:
                self.warning('Неправильний LDK additional data magic 0x%08x.', add_hdr.get('magic'));
                return None

            additional_data_size_val = add_hdr.get('size')
            if additional_data_size_val < 0 or self._tell() + additional_data_size_val > self.rawsize:
                self.error(f"Некоректний розмір дод. блоку даних LDK: {additional_data_size_val}");
                return None

            additional_data_block = self._getval('bin', additional_data_size_val)
            if additional_data_block is None: self.error("Не вдалося прочитати дод. блок даних LDK."); return None
            data_chunks.append(additional_data_block)
            current_add_offset_val = add_hdr.get('addOffset')
        if safety_counter >= 100: self.warning("LDK: Досягнуто ліміту обробки додаткових блоків даних.")
        return b"".join(data_chunks)

    def _get_node(self, offset, current_path_prefix="/", uid_for_path=None):
        if offset >= self.rawsize:  # Check offset validity
            self.error(f"Некоректний offset вузла LDK: {offset}");
            return None
        self.debug('LDK Node at 0x%04x', offset)
        self._seek(offset)
        hdr = self._getvalmulti(magic='int', flags='int', metaOffset='pointer', res1='long')
        if None in [hdr.get('magic'), hdr.get('metaOffset')]:
            self.error("Не вдалося прочитати заголовок вузла LDK (обов'язкові поля).");
            return None
        if hdr.get('magic') != 0x00015555:
            self.warning('Невідомий LDK node magic 0x%08x.', hdr.get('magic'));
            return None

        meta_offset_val = hdr.get('metaOffset')
        if meta_offset_val == 0 or meta_offset_val + 0x20 >= self.rawsize:  # Meta offset check
            self.error(f"Некоректний metaOffset LDK: {meta_offset_val}");
            return None

        prev_offs = self._tell()
        self._seek(meta_offset_val + 0x20)  # Meta data starts after a 0x20 byte header in the meta block
        node_meta = self._get_metadata()
        self._seek(prev_offs)  # Restore offset to continue reading node structure

        node_path = current_path_prefix
        if uid_for_path is not None:
            node_name_from_meta = node_meta.get('name') if node_meta else None
            # Sanitize node_name_from_meta for path component
            safe_node_name = re.sub(r'[\\/*?:"<>|]', '_', node_name_from_meta) if node_name_from_meta else None
            node_path += f"{safe_node_name}/" if safe_node_name else f"UID{uid_for_path:08X}/"

        node_entries_magic = self._getval('int')
        if node_entries_magic is None: self.error("Не вдалося прочитати magic для записів вузла LDK."); return None
        self.debug('LDK node path=%s, nodeEntriesMagic=0x%08x', node_path, node_entries_magic)

        node_obj = {'path': node_path, 'nodes': [], 'files': [], 'meta': node_meta if node_meta else {}}
        n_child, n_data, n_empty = 0, 0, 0

        if node_entries_magic == 0x00025555:  # List type node
            list_hdr = self._getvalmulti(nTotal='int', nChild='int', nData='int', addOffset='pointer')
            if None in [list_hdr.get('nTotal'), list_hdr.get('nChild'), list_hdr.get('nData')]: return None
            n_child, n_data = list_hdr.get('nChild', 0), list_hdr.get('nData', 0)
            n_empty = list_hdr.get('nTotal', 0) - n_child - n_data
        elif node_entries_magic == 0x00045555:  # Table type node (hash table for entries)
            # The structure for table nodes is more complex and might involve hash lookups.
            # This simplified parsing assumes entries are still somewhat linear for now.
            # A more accurate parsing would require understanding the hash table structure.
            # For now, we read nChild and nData assuming they are directly available.
            # This part might need significant refinement if LDK uses complex hash tables.
            self.warning("LDK: Обробка вузла типу 'таблиця' (0x00045555) може бути неповною.")
            # Attempt to read nChild and nData, assuming a simple structure first
            table_hdr_simple = self._getvalmulti(nChild='int', nData='int')
            if table_hdr_simple.get('nChild') is not None and table_hdr_simple.get('nData') is not None:
                n_child, n_data = table_hdr_simple.get('nChild', 0), table_hdr_simple.get('nData', 0)
            else:  # Fallback if the simple read fails, might need to parse hash table structure
                self.error("LDK: Не вдалося прочитати nChild/nData для вузла-таблиці. Структура невідома.");
                return None
        else:
            self.warning('Неправильний LDK node entries magic 0x%08x.', node_entries_magic);
            return None

        entry_size = 12  # Each entry (offset + uid + 4 reserved bytes) is 12 bytes (Q + i = 8 + 4)
        child_defs, data_defs = [], []

        for i in range(n_child):
            # Each child_def is 12 bytes: 8 for offset (pointer), 4 for uid (int)
            d = self._getvalmulti(offset='pointer', uid='int')
            if None in [d.get('offset'), d.get('uid')]:
                self.error(f"Помилка читання child_def {i}");
                return None
            d['_ix'] = i;
            child_defs.append(d)
            self.trace('LDK childDef[%d]: off=0x%x uid=0x%x', i, d['offset'], d['uid'])

        if n_empty < 0: self.warning(f"Негативна кількість порожніх записів ({n_empty}) у вузлі LDK."); n_empty = 0
        bytes_to_skip = n_empty * entry_size
        if self._tell() + bytes_to_skip > self.rawsize:
            self.error(f"LDK: Спроба пропустити порожні записи виходить за межі файлу.");
            return None
        self._seek(self._tell() + bytes_to_skip)  # Skip empty entries

        for i in range(n_data):
            # Each data_def is 12 bytes: 8 for offset (pointer), 4 for uid (int)
            d = self._getvalmulti(offset='pointer', uid='int')
            if None in [d.get('offset'), d.get('uid')]:
                self.error(f"Помилка читання data_def {i}");
                return None
            d['_ix'] = i;
            data_defs.append(d)
            self.trace('LDK dataDef[%d]: off=0x%x uid=0x%x', i, d['offset'], d['uid'])

        for entry_def in sorted(child_defs, key=lambda x: x['_ix']):
            if entry_def['offset'] == 0:  # Skip null offsets
                self.warning(f"LDK: Нульовий offset для дочірнього вузла UID {entry_def['uid']}. Пропускається.");
                continue
            child_node = self._get_node(entry_def['offset'], node_path, entry_def['uid'])
            if child_node is None:
                self.error(f"Помилка парсингу дочірнього вузла LDK (offset {entry_def['offset']}).");
                # Continue parsing other children if one fails, rather than failing the whole node
                continue
            child_node['order'] = entry_def['_ix'];
            node_obj['nodes'].append(child_node)

        type_map_ldk = {0x65: 'wpt', 0x66: 'set', 0x67: 'rte', 0x68: 'trk', 0x69: 'are'}
        ldk_original_filename = self.path or self.rawname or "unknown.ldk"
        ldk_base_fn_for_contained = os.path.splitext(os.path.basename(ldk_original_filename))[0]

        for entry_def in sorted(data_defs, key=lambda x: x['_ix']):
            if entry_def['offset'] == 0:  # Skip null offsets
                self.warning(f"LDK: Нульовий offset для файлу UID {entry_def['uid']}. Пропускається.");
                continue
            file_bytes = self._get_node_data(entry_def['offset'])
            if file_bytes is None or not file_bytes:
                self.warning(f"Пропущено порожній/пошкоджений файл у LDK (UID {entry_def.get('uid', 'N/A')})")
                continue

            if not file_bytes:  # Should be caught by above, but double check
                self.warning(f"LDK: Нульовий вміст файлу для UID {entry_def['uid']}. Пропускається.");
                continue

            file_type_val = file_bytes[0]  # First byte is the type
            actual_data_bytes = file_bytes[1:]  # Rest is the actual file data

            if not actual_data_bytes:  # If only type byte was present
                self.warning(f"LDK: Файл UID {entry_def['uid']} містить тільки байт типу. Пропускається.");
                continue

            data_b64_str = base64.b64encode(actual_data_bytes).decode('ascii')
            type_str_from_map = type_map_ldk.get(file_type_val, 'bin')

            path_part_for_name = node_obj['path'].strip('/').replace('/', '_')
            if path_part_for_name: path_part_for_name = "_" + path_part_for_name

            contained_file_unique_name = f"{ldk_base_fn_for_contained}{path_part_for_name}_UID{entry_def.get('uid', 0):08X}.{type_str_from_map}"

            node_obj['files'].append({
                'name': contained_file_unique_name, 'data_b64': data_b64_str, 'type': type_str_from_map,
                'size': len(actual_data_bytes), 'order': entry_def['_ix']
            })
            self.debug('LDK file extracted: %s (type: %s, size: %d bytes)', contained_file_unique_name,
                       type_str_from_map, len(actual_data_bytes))
        return node_obj

    def get_parsed_data(self):
        output_data = {
            'ts': self.rawts, 'type': self._file_type,
            'path': self.path or self.rawname,
            'file': os.path.basename(self.path or self.rawname or "unknown_file"),  # Keep original filename
            'parse_successful': self.parse_successful
        }
        if self.parse_successful:
            # data_parsed should contain the direct output of the _parse_xxx method
            if self._file_type == 'wpt':
                output_data.update({'meta': self.data_parsed.get('meta'), 'location': self.data_parsed.get('location')})
            elif self._file_type in ['set', 'rte']:  # Grouped SET and RTE
                output_data.update(
                    {'meta': self.data_parsed.get('meta'), 'waypoints': self.data_parsed.get('waypoints')})
            elif self._file_type == 'are':
                output_data.update(
                    {'meta': self.data_parsed.get('meta'), 'locations': self.data_parsed.get('locations')})
            elif self._file_type == 'trk':
                output_data.update(
                    {'meta': self.data_parsed.get('meta'),
                     'waypoints': self.data_parsed.get('waypoints'),  # These are POIs for TRK
                     'segments': self.data_parsed.get('segments')})
            elif self._file_type == 'ldk':
                output_data['root'] = self.data_parsed.get('root')  # The hierarchical structure
            elif self._file_type == 'bin':
                output_data['raw_content_b64'] = self.data_parsed.get('raw_content_b64')  # Store as b64
        else:  # If parsing failed, still provide basic info
            self.error(f"Парсинг файлу {output_data['file']} (тип: {output_data['type']}) не вдався.")
        return output_data
