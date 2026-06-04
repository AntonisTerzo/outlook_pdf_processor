import re
from pypdf import PdfReader


def extract_versand_for_tianjin(pdf_path):
    """
    Extract Versand value specifically for TIANJIN city.
    Searches for one of: China TI ASS, China TI IG 4, China TI IG, China TI ED, China TI PROD
    Returns the found value or None.
    """
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text()

            # Find the section after "Versand:"
            versand_match = re.search(
                r'Versand:(.*?)(?=\n\n|Warenempfänger:|Besteller|$)', text, re.DOTALL)

            if versand_match:
                versand_section = versand_match.group(1)

                # List of possible Versand values for TIANJIN (ordered by specificity - longest first)
                tianjin_versand_options = [
                    "China TI IG 4",
                    "China TI ASS",
                    "China TI PROD",
                    "China TI ED",
                    "China TI IG"
                ]

                # Search for each option in order
                for option in tianjin_versand_options:
                    if re.search(r'\b' + re.escape(option) + r'\b', versand_section, re.IGNORECASE):
                        return option

            return None
    except Exception as e:
        print(f"Error extracting Versand for TIANJIN: {e}")
        return None


def check_motors_in_warenempfanger(pdf_path):
    """
    Check if "Motors" appears in Warenempfänger section for SUZHOU.
    Returns True if Motors found, False otherwise.
    """
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text()

            # Find the section after "Warenempfänger:"
            warenempfanger_match = re.search(
                r'Warenempfänger:(.*?)(?=Besteller|Lieferkondition:|$)', text, re.DOTALL)

            if warenempfanger_match:
                warenempfanger_section = warenempfanger_match.group(1)
                
                # Check for "Motors"
                if re.search(r'\bMotors\b', warenempfanger_section, re.IGNORECASE):
                    return True

            return False
    except Exception as e:
        print(f"Error checking Motors for SUZHOU: {e}")
        return False


def extract_info_from_pdf_task1(pdf_path, cities_list):
    """
    Extract city name and document number from PDF for Task 1.
    City: Search for any city from the cities_list after 'Warenempfänger:' (takes LAST match)
    Number: 8-digit number after 'Pack- und Gewichtsliste'
    """
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text()

            # Find the section after "Warenempfänger:"
            warenempfanger_match = re.search(
                r'Warenempfänger:(.*?)(?=Besteller|Lieferkondition:|$)', text, re.DOTALL)

            city = None
            if warenempfanger_match:
                warenempfanger_section = warenempfanger_match.group(1)

                # Find ALL cities in the section, then take the LAST one
                found_cities = []
                for city_name in cities_list:
                    # Find all matches for this city
                    for match in re.finditer(r'\b' + re.escape(city_name) + r'\b',
                                             warenempfanger_section, re.IGNORECASE):
                        found_cities.append((match.start(), city_name))

                # Sort by position and take the last one
                if found_cities:
                    found_cities.sort(key=lambda x: x[0])
                    city = found_cities[-1][1]

            # Find 8-digit number after "Pack- und Gewichtsliste"
            number_pattern = r'Pack- und Gewichtsliste\s+Nr\.\s+(\d{8})'
            number_match = re.search(number_pattern, text)

            if city and number_match:
                doc_number = number_match.group(1)
                return city, doc_number

            return None, None
    except Exception as e:
        print(f"Error extracting info from PDF: {e}")
        return None, None


def check_city_inside_pdf(pdf_path, cities_list):
    """
    Check inside PDF for city names (same method as Task 1).
    Searches in the Warenempfänger section and returns the LAST matching city.
    Used for Brasilien subcities and manual review fallback.
    Returns the city name if found, None otherwise.
    """
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text()

            # Find the section after "Warenempfänger:"
            # Stop at "Besteller" (sender section) OR "Pack- und Gewichtsliste"
            warenempfanger_match = re.search(
                r'Warenempfänger:(.*?)(?=Besteller|Pack- und Gewichtsliste|Lieferkondition:|$)', text, re.DOTALL)

            city = None
            if warenempfanger_match:
                warenempfanger_section = warenempfanger_match.group(1)

                # Find ALL cities in the section, then take the LAST one
                found_cities = []
                for city_name in cities_list:
                    # Find all matches for this city
                    for match in re.finditer(r'\b' + re.escape(city_name) + r'\b',
                                             warenempfanger_section, re.IGNORECASE):
                        found_cities.append((match.start(), city_name))

                # Sort by position and take the last one
                if found_cities:
                    found_cities.sort(key=lambda x: x[0])
                    city = found_cities[-1][1]

            return city

    except Exception as e:
        print(f"Error checking city inside PDF: {e}")
        return None


def extract_city_from_filename_task2(filename, cities_list):
    """
    Extract city name from PDF filename for Task 2.
    Handles cities with multiple words that may have text in between.
    Example: "Japan XYZ Kyoto" should match "JAPAN KYOTO"
    Returns the city name if found, None otherwise.
    """
    try:
        # Remove .pdf extension
        name_without_ext = filename.replace('.pdf', '').replace('.PDF', '')

        # Sort cities by length (longest first) to match specific names before generic ones
        sorted_cities = sorted(cities_list, key=len, reverse=True)

        # Search for each city in the filename
        for city_name in sorted_cities:
            # Split city name into words
            city_words = city_name.split()

            if len(city_words) == 1:
                # Single word city - simple match
                if re.search(r'\b' + re.escape(city_name) + r'\b', name_without_ext, re.IGNORECASE):
                    return city_name
            else:
                # Multi-word city - allow any text between words
                # Example: "JAPAN KYOTO" becomes pattern "JAPAN.*KYOTO"
                pattern = r'\b' + r'.*'.join(re.escape(word)
                                             for word in city_words) + r'\b'
                if re.search(pattern, name_without_ext, re.IGNORECASE):
                    return city_name

        return None
    except Exception as e:
        print(f"Error extracting city from filename: {e}")
        return None


def round_dimension(value_mm):
    """
    Convert mm to centimeters and round based on .5 threshold.
    Examples:
    - 680 -> 68.0 -> 68 cm
    - 555 -> 55.5 -> 56 cm (>= .5)
    - 554 -> 55.4 -> 55 cm (< .5)
    - 10 -> 1.0 -> 1 cm
    - 25 -> 2.5 -> 3 cm (>= .5)
    """
    centimeters = value_mm / 10

    # Get the decimal part
    decimal_part = centimeters - int(centimeters)

    if decimal_part >= 0.5:
        return int(centimeters) + 1
    else:
        return int(centimeters)


def extract_dimensions_from_pdf(pdf_path):
    """
    Extract dimensions from PDF ONLY from the "Abmessung(MM)" column.
    Handles cases where dimensions span multiple pages.
    Returns tuple: (list of dimension strings, warning message or None)
    - Dimensions in format "LxWxH" (in centimeters, rounded)
    - Warning if any Abmessung section had no dimensions
    Example: (["68x36x47", "55x30x40"], None) or ([], "WARNING: Found 2 Abmessung sections with no dimensions")
    """
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text()

            dimensions = []

            # Split text into lines
            lines = text.split('\n')

            # Track Abmessung sections and whether they have dimensions
            in_abmessung_section = False
            current_section_has_dimensions = False
            abmessung_sections_without_dimensions = 0

            for i, line in enumerate(lines):
                # Check if this line contains "Abmessung(MM)" header
                if 'Abmessung' in line and 'MM' in line:
                    # If we were already in a section, check if it had dimensions
                    if in_abmessung_section and not current_section_has_dimensions:
                        abmessung_sections_without_dimensions += 1

                    # Start new section
                    in_abmessung_section = True
                    current_section_has_dimensions = False
                    continue

                if in_abmessung_section:
                    # Check if we've hit Bestellung section - stop tracking this Abmessung
                    if 'Bestellung' in line:
                        # Before closing section, check if it had dimensions
                        if not current_section_has_dimensions:
                            abmessung_sections_without_dimensions += 1

                        in_abmessung_section = False
                        continue

                    # Check for "variofix" - use fixed dimensions
                    if 'variofix' in line.lower():
                        dimensions.append("36x23x19")
                        current_section_has_dimensions = True
                        continue

                    # Look for dimension pattern - no line limit, just until next section
                    dimension_pattern = r'(\d{2,4})[xX](\d{2,4})[xX](\d{2,4})'
                    matches = re.finditer(dimension_pattern, line)

                    for match in matches:
                        length_mm = int(match.group(1))
                        width_mm = int(match.group(2))
                        height_mm = int(match.group(3))

                        # Convert to centimeters with rounding
                        length_cm = round_dimension(length_mm)
                        width_cm = round_dimension(width_mm)
                        height_cm = round_dimension(height_mm)

                        # Format as "LxWxH"
                        dimension_str = f"{length_cm}x{width_cm}x{height_cm}"
                        dimensions.append(dimension_str)
                        current_section_has_dimensions = True

            # Check the last section if we ended while still in one
            if in_abmessung_section and not current_section_has_dimensions:
                abmessung_sections_without_dimensions += 1

            # Create warning message if needed
            warning = None
            if abmessung_sections_without_dimensions > 0:
                warning = f"WARNING: Found {abmessung_sections_without_dimensions} Abmessung section(s) with no dimensions"

            return dimensions, warning

    except Exception as e:
        print(f"Error extracting dimensions from PDF: {e}")
        return [], None




# ============================================================
# Task 3 helpers  (pdfplumber-based)
#
# Task 3 PDFs are digital-born documents with selectable text and tables
# drawn with visible gridlines. pypdf flattens the table reading order, so
# Task 3 uses pdfplumber instead, which reconstructs real table cells from
# the ruling lines. Tasks 1 & 2 are unaffected and still use pypdf above.
#
# These readers accept raw PDF bytes (Task 3 never saves the files to disk).
# Only two file types matter: PACKING_LIST and CUSTOMS_CODE. The invoice
# file is ignored entirely.
# ============================================================

import io
import pdfplumber

# pdfplumber settings tuned for tables with visible borders.
_LINES_TABLE_SETTINGS = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
}


def _open_plumber(pdf_bytes):
    """Open a pdfplumber document from bytes. Caller must close it."""
    return pdfplumber.open(io.BytesIO(pdf_bytes))


def _read_pdf_text_from_bytes(pdf_bytes):
    """Read full plain text from PDF bytes via pdfplumber (for type detection
    and regex fallbacks). Returns string (empty on error)."""
    try:
        parts = []
        with _open_plumber(pdf_bytes) as pdf:
            for page in pdf.pages:
                parts.append(page.extract_text() or "")
        return "\n".join(parts)
    except Exception as e:
        print(f"Error reading PDF text from bytes: {e}")
        return ""


def _all_tables(pdf_bytes):
    """
    Return every table on every page as a list of tables, each a list of rows,
    each row a list of cell strings (None cells become '').
    """
    tables = []
    try:
        with _open_plumber(pdf_bytes) as pdf:
            for page in pdf.pages:
                for raw in page.extract_tables(_LINES_TABLE_SETTINGS):
                    cleaned = [
                        [(cell or "").strip().replace("\n", " ") for cell in row]
                        for row in raw
                    ]
                    tables.append(cleaned)
    except Exception as e:
        print(f"Error extracting tables: {e}")
    return tables


def detect_file_type_task3(pdf_bytes, file_types_map):
    """
    Detect which Task 3 file type a PDF is, from raw bytes.
    Returns the type key ("PACKING_LIST"/"CUSTOMS_CODE") or None if neither
    (e.g. the invoice file, which we deliberately do not match).
    """
    text = _read_pdf_text_from_bytes(pdf_bytes)
    if not text:
        return None
    text_lower = text.lower()
    for type_key in ("PACKING_LIST", "CUSTOMS_CODE"):
        if type_key not in file_types_map:
            continue
        for identifier in file_types_map[type_key]:
            if identifier.lower() in text_lower:
                return type_key
    return None


def extract_combo_id_from_filename(filename):
    """
    Extract the combo identifier (a number string) from a filename.
    The combo is whatever the longest run of digits in the filename is.
    Returns the digit string or None.
    """
    name = filename.replace('.pdf', '').replace('.PDF', '')
    numbers = re.findall(r'\d+', name)
    if not numbers:
        return None
    numbers.sort(key=len, reverse=True)
    return numbers[0]


def _to_cm(value_str):
    """
    Convert a dimension string (meters, may use comma or dot decimal) to an
    integer number of centimetres. '1,20'->120, '1.20'->120, '0.8'->80.
    """
    s = value_str.strip()
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    val = float(s)
    return int(round(val * 100))


def _norm(s):
    """Lowercase and collapse whitespace for header matching."""
    return re.sub(r'\s+', ' ', (s or "").strip().lower())


def _find_header_column(table, *keywords):
    """
    Search a table's rows for a header row containing a cell that matches ALL
    given keywords (case-insensitive substring). Returns (header_row_index,
    column_index) or (None, None) if not found.
    """
    for r_idx, row in enumerate(table):
        for c_idx, cell in enumerate(row):
            cell_n = _norm(cell)
            if all(kw in cell_n for kw in keywords):
                return r_idx, c_idx
    return None, None


def _extract_descriptions_from_tables(tables):
    """
    Find the cargo descriptions: the column headed 'Descripción', which sits
    before the 'Cod. Modelo / Model id' column. Returns a list of unique,
    in-order description strings.
    """
    descriptions = []
    seen = set()

    for table in tables:
        # Find a header row that has BOTH a "descripcion" cell and a
        # "model id" / "modelo" cell (confirms it's the cargo table).
        desc_col = None
        header_row = None
        for r_idx, row in enumerate(table):
            row_norm = [_norm(c) for c in row]
            d_col = next((i for i, c in enumerate(row_norm)
                          if "descripci" in c), None)
            has_model = any(("modelo" in c) or ("model id" in c)
                            for c in row_norm)
            if d_col is not None and has_model:
                desc_col = d_col
                header_row = r_idx
                break

        if desc_col is None:
            continue

        # Collect description cells below the header
        for row in table[header_row + 1:]:
            if desc_col >= len(row):
                continue
            val = (row[desc_col] or "").strip()
            if not val:
                continue
            # Skip rows that are clearly totals/labels leaking into the table
            if re.search(r'(resumen|summary|total|peso\s*bruto|volumen)',
                         val, re.IGNORECASE):
                continue
            if val not in seen and re.search(r'[A-Za-zÀ-ÿ]', val):
                seen.add(val)
                descriptions.append(val)

    return descriptions


def _extract_dims_from_tables(tables):
    """
    Find the dimensions table (headers Long./Length, Anchura/Width,
    Altura/Height) and collect EVERY dimension row.

    Identical dimensions are counted, and each distinct dimension is prefixed
    with its count as "count-LxWxH". Multiple distinct dimensions are joined
    by a single space, e.g. "2-80x50x15 1-43x35x31". All values in cm, no
    internal whitespace within a dimension token. Returns the string or None.
    """
    from collections import Counter

    ordered = []          # preserves first-seen order of distinct dims
    counts = Counter()

    for table in tables:
        len_col = wid_col = hgt_col = None
        header_row = None
        for r_idx, row in enumerate(table):
            row_norm = [_norm(c) for c in row]
            for c_idx, c in enumerate(row_norm):
                if "length" in c or "long." in c or "longitud" in c:
                    len_col = c_idx
                if "width" in c or "anchura" in c:
                    wid_col = c_idx
                if "height" in c or "altura" in c:
                    hgt_col = c_idx
            if None not in (len_col, wid_col, hgt_col):
                header_row = r_idx
                break

        if header_row is None:
            continue

        # Read ALL data rows under the header (not just the first).
        for row in table[header_row + 1:]:
            try:
                l_raw = row[len_col].strip()
                w_raw = row[wid_col].strip()
                h_raw = row[hgt_col].strip()
            except (IndexError, AttributeError):
                continue
            if not (l_raw and w_raw and h_raw):
                continue
            if not all(re.search(r'\d', x) for x in (l_raw, w_raw, h_raw)):
                continue
            try:
                dim = f"{_to_cm(l_raw)}x{_to_cm(w_raw)}x{_to_cm(h_raw)}"
            except Exception as e:
                print(f"Error parsing dims from table: {e}")
                continue
            if dim not in counts:
                ordered.append(dim)
            counts[dim] += 1

    if not ordered:
        return None

    # Format: "count-LxWxH", joined by single spaces, fully trimmed.
    parts = [f"{counts[dim]}-{dim}" for dim in ordered]
    return " ".join(parts)


def _number_after_label(cell_text, label_keywords):
    """
    Within a single cell, return the numeric value that appears AFTER the
    label text, ignoring digits that are part of unit markers (m3, m2, kgs).
    Returns the number string or None.
    """
    text = cell_text or ""
    # Find where the label ends: use the last keyword's position.
    last_kw = label_keywords[-1]
    pos = text.lower().rfind(last_kw.lower())
    search_from = pos + len(last_kw) if pos != -1 else 0
    tail = text[search_from:]
    # Remove unit markers so their digits aren't picked up (m3, m2, kgs, etc.)
    tail = re.sub(r'm\s*[23]\b', ' ', tail, flags=re.IGNORECASE)
    tail = re.sub(r'\bkgs?\b', ' ', tail, flags=re.IGNORECASE)
    m = re.search(r'(\d[\d\.,]*)', tail)
    return m.group(1) if m else None


def _summary_value(tables, full_text, *label_keywords):
    """
    Find a summary value (PCS/KG/M3). First look through tables for a row whose
    label cell matches all keywords. The number may sit in a later cell of the
    same row, or after the label inside the same cell. Falls back to a text
    regex if not found in tables.
    Digits embedded in unit markers (the '3' in 'm3', '2' in 'm2', etc.) are
    ignored so they aren't mistaken for the value.
    """
    # Table-based lookup
    for table in tables:
        for row in table:
            row_norm = [_norm(c) for c in row]
            label_idx = next(
                (i for i, c in enumerate(row_norm)
                 if all(kw in c for kw in label_keywords)),
                None
            )
            if label_idx is None:
                continue
            # Prefer a number in a later cell of the same row.
            for later in row[label_idx + 1:]:
                cleaned = re.sub(r'm\s*[23]\b', ' ', later or "", flags=re.IGNORECASE)
                cleaned = re.sub(r'\bkgs?\b', ' ', cleaned, flags=re.IGNORECASE)
                m = re.search(r'(\d[\d\.,]*)', cleaned)
                if m:
                    return m.group(1)
            # Otherwise take the number after the label within the same cell.
            same = _number_after_label(row[label_idx], list(label_keywords))
            if same:
                return same

    # Text fallback: keyword phrase, then the first number after it,
    # skipping unit-embedded digits.
    pattern = r'.*?'.join(re.escape(k) for k in label_keywords)
    m = re.search(pattern, full_text, re.IGNORECASE)
    if m:
        tail = full_text[m.end():]
        tail = re.sub(r'm\s*[23]\b', ' ', tail, flags=re.IGNORECASE)
        tail = re.sub(r'\bkgs?\b', ' ', tail, flags=re.IGNORECASE)
        m2 = re.search(r'(\d[\d\.,]*)', tail)
        if m2:
            return m2.group(1)
    return None


def extract_packing_list_data(pdf_bytes):
    """
    Extract all needed fields from a PACKING_LIST PDF (given as bytes), using
    pdfplumber table parsing with text-based fallbacks.

    Returns dict: descriptions (list[str]), iv, srn, pcs, kg, m3, dims
    """
    result = {
        "descriptions": [], "iv": None, "srn": None,
        "pcs": None, "kg": None, "m3": None, "dims": None,
    }

    text = _read_pdf_text_from_bytes(pdf_bytes)
    tables = _all_tables(pdf_bytes)
    if not text and not tables:
        return result

    # IV: value after "Factura / Invoice"
    iv_match = re.search(
        r'Factura\s*/\s*Invoice\s*[:\-]?\s*([A-Z0-9][A-Z0-9\-_/]*)',
        text, re.IGNORECASE)
    if iv_match:
        result["iv"] = iv_match.group(1).strip()

    # SRN: value after "Envío / Shipment Nr"
    srn_match = re.search(
        r'Env[ií]o\s*/\s*Shipment\s*Nr\s*[\.:]*\s*([A-Z0-9][A-Z0-9\-_/]*)',
        text, re.IGNORECASE)
    if srn_match:
        result["srn"] = srn_match.group(1).strip()

    # Descriptions: from the cargo table
    result["descriptions"] = _extract_descriptions_from_tables(tables)

    # Summary fields
    result["pcs"] = _summary_value(tables, text, "parcels")
    if not result["pcs"]:
        result["pcs"] = _summary_value(tables, text, "bultos")
    result["kg"] = _summary_value(tables, text, "gross", "weight")
    if not result["kg"]:
        result["kg"] = _summary_value(tables, text, "peso", "bruto")
    result["m3"] = _summary_value(tables, text, "total", "volume")
    if not result["m3"]:
        result["m3"] = _summary_value(tables, text, "volumen")

    # Dimensions
    result["dims"] = _extract_dims_from_tables(tables)

    return result


def extract_hs_codes(pdf_bytes):
    """
    Extract HS codes from a CUSTOMS_CODE PDF (bytes). Looks for the column
    headed 'Cod. Arancel / Customs code', reads the codes beneath it, and
    strips the dots. Returns a list of unique codes in order.
    Falls back to a text scan if the table header isn't found.
    """
    hs_codes = []
    seen = set()

    tables = _all_tables(pdf_bytes)
    for table in tables:
        # Find a column whose header mentions arancel / customs code
        code_col = None
        header_row = None
        for r_idx, row in enumerate(table):
            row_norm = [_norm(c) for c in row]
            c_idx = next((i for i, c in enumerate(row_norm)
                          if ("arancel" in c) or ("customs code" in c)
                          or ("customs" in c and "code" in c)), None)
            if c_idx is not None:
                code_col = c_idx
                header_row = r_idx
                break
        if code_col is None:
            continue

        for row in table[header_row + 1:]:
            if code_col >= len(row):
                continue
            raw = (row[code_col] or "").strip()
            # A code may look like 6403.99.96 or 64039996
            m = re.search(r'\d{2,5}(?:\.\d{1,5}){0,4}', raw)
            if not m:
                continue
            digits = m.group(0).replace('.', '')
            if 6 <= len(digits) <= 12 and digits not in seen:
                seen.add(digits)
                hs_codes.append(digits)

    if hs_codes:
        return hs_codes

    # Fallback: text scan after the header phrase
    text = _read_pdf_text_from_bytes(pdf_bytes)
    header_match = re.search(
        r'Cod\.?\s*Arancel\s*/\s*Customs\s*code(.*)$',
        text, re.IGNORECASE | re.DOTALL)
    section = header_match.group(1) if header_match else text
    for cand in re.findall(r'\b\d{2,5}(?:\.\d{1,5}){0,4}\b', section):
        digits = cand.replace('.', '')
        if 6 <= len(digits) <= 12 and digits not in seen:
            seen.add(digits)
            hs_codes.append(digits)
    return hs_codes


def extract_cargo_description_from_subject(subject):
    """
    Return all the text between the two dashes that come immediately before
    CHINA in the subject. Example:
        "shipment - world wide cargo - CHINA" -> "world wide cargo"
    Falls back to the single word before CHINA if no dash pattern is present.
    """
    if not subject:
        return None
    m = re.search(r'-\s*([^-]+?)\s*-\s*CHINA\b', subject, re.IGNORECASE)
    if m:
        inside = m.group(1).strip()
        if inside:
            return inside
    m2 = re.search(r'(\S+)\s+CHINA\b', subject, re.IGNORECASE)
    if m2:
        return m2.group(1)
    return None
