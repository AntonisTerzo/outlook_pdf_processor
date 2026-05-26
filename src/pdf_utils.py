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
# Task 3 helpers
# ============================================================

def _read_pdf_text(pdf_path):
    """Read full text from PDF. Returns string (empty on error)."""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PdfReader(file)
            return "\n".join(page.extract_text() or "" for page in pdf_reader.pages)
    except Exception as e:
        print(f"Error reading PDF text: {e}")
        return ""


def detect_file_type_task3(pdf_path, file_types_map):
    """
    Detect which of the 3 Task 3 file types a PDF is.
    file_types_map: dict like {"INVOICE": [...identifiers...], ...}
    Returns the type key (e.g. "INVOICE") or None if no match.
    Note: PACKING_LIST and INVOICE share the substring "FACTURA / INVOICE"
    on the packing-list header, so check PACKING_LIST first.
    """
    text = _read_pdf_text(pdf_path)
    if not text:
        return None

    # Check in priority order: PACKING_LIST and CUSTOMS_CODE have more specific
    # headers; INVOICE check comes last because its identifier also appears
    # on the packing list page.
    priority = ["PACKING_LIST", "CUSTOMS_CODE", "INVOICE"]
    for type_key in priority:
        if type_key not in file_types_map:
            continue
        for identifier in file_types_map[type_key]:
            if identifier.lower() in text.lower():
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
    # Longest digit run is the combo identifier
    numbers.sort(key=len, reverse=True)
    return numbers[0]


def extract_packing_list_data(pdf_path):
    """
    Extract all data needed from a PACKING_LIST PDF.
    Returns dict with keys:
        descriptions (list[str]), iv, srn, pcs, kg, m3, dims
    Missing fields are None (or empty list for descriptions).
    """
    text = _read_pdf_text(pdf_path)
    result = {
        "descriptions": [],
        "iv": None,
        "srn": None,
        "pcs": None,
        "kg": None,
        "m3": None,
        "dims": None,
    }
    if not text:
        return result

    # --- IV: after "Factura/Invoice" near the top of the file ---
    iv_match = re.search(
        r'Factura\s*/\s*Invoice[^\w\d]*([A-Z0-9\-_/]+)',
        text, re.IGNORECASE
    )
    if iv_match:
        result["iv"] = iv_match.group(1).strip()

    # --- SRN: after "Nº Envío/Shipment Nr" ---
    srn_match = re.search(
        r'N[ºo°]?\s*Env[ií]o\s*/\s*Shipment\s*Nr[\.:]?\s*([A-Z0-9\-_/]+)',
        text, re.IGNORECASE
    )
    if srn_match:
        result["srn"] = srn_match.group(1).strip()

    # --- Descriptions: all values under "Descripción" column ---
    # Strategy: find each line containing the descripcion header, then read
    # subsequent rows until we hit RESUMEN/SUMMARY or the end.
    descripcion_section = re.search(
        r'Descripci[oó]n(.*?)(?=RESUMEN\s*/\s*SUMMARY|$)',
        text, re.IGNORECASE | re.DOTALL
    )
    if descripcion_section:
        section_text = descripcion_section.group(1)
        # Each non-empty line that has some letters is a candidate description.
        # Skip lines that are purely numeric / dimension rows.
        candidates = []
        for raw_line in section_text.split('\n'):
            line = raw_line.strip()
            if not line:
                continue
            # Skip if it's a pure number/dimension row
            if re.fullmatch(r'[\d\s\.\,xX]+', line):
                continue
            # Skip header-like leftovers
            if re.search(r'(long|length|anchura|width|altura|height|peso|weight|volumen|volume|bultos|parcels)',
                         line, re.IGNORECASE):
                continue
            # Likely description text
            if re.search(r'[A-Za-zÀ-ÿ]{3,}', line):
                candidates.append(line)
        # Deduplicate while preserving order
        seen = set()
        for c in candidates:
            if c not in seen:
                seen.add(c)
                result["descriptions"].append(c)

    # --- Summary section ---
    summary_match = re.search(
        r'RESUMEN\s*/\s*SUMMARY(.*)$',
        text, re.IGNORECASE | re.DOTALL
    )
    summary_text = summary_match.group(1) if summary_match else text

    # PCS: after "Nº Total de Bultos/Total Nª Parcels"
    pcs_match = re.search(
        r'N[ºo°]?\s*Total\s*de\s*Bultos\s*/\s*Total\s*N[ºoª°]?\s*Parcels[^\d]*([\d\.,]+)',
        summary_text, re.IGNORECASE
    )
    if pcs_match:
        result["pcs"] = pcs_match.group(1).strip()

    # KG: after "Peso Bruto Total/Total Gross Weight: (Kgs)"
    kg_match = re.search(
        r'Peso\s*Bruto\s*Total\s*/\s*Total\s*Gross\s*Weight[^\d]*([\d\.,]+)',
        summary_text, re.IGNORECASE
    )
    if kg_match:
        result["kg"] = kg_match.group(1).strip()

    # M3: after "Volumen Total/Total Volume: (m3)"
    m3_match = re.search(
        r'Volumen\s*Total\s*/\s*Total\s*Volume[^\d]*([\d\.,]+)',
        summary_text, re.IGNORECASE
    )
    if m3_match:
        result["m3"] = m3_match.group(1).strip()

    # --- DIMS: from Long./Length (m) Anchura/Width (m) Altura/Height (m) ---
    # We grab 3 consecutive numeric values right after that header.
    dims_match = re.search(
        r'Long\.?\s*/\s*Length.*?Anchura\s*/\s*Width.*?Altura\s*/\s*Height[^\d]*'
        r'([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)',
        text, re.IGNORECASE | re.DOTALL
    )
    if dims_match:
        try:
            # Values are in meters - convert to cm.
            def to_cm(s):
                # Replace comma decimal with dot, then convert to float
                val = float(s.replace('.', '').replace(',', '.')) if ',' in s else float(s)
                # If value seems already in cm (>10), assume cm; otherwise m->cm
                if val < 20:
                    val = val * 100
                return int(round(val))

            l = to_cm(dims_match.group(1))
            w = to_cm(dims_match.group(2))
            h = to_cm(dims_match.group(3))
            result["dims"] = f"{l}x{w}x{h}"
        except Exception as e:
            print(f"Error parsing dims: {e}")

    return result


def extract_hs_codes(pdf_path):
    """
    Extract HS codes from a CUSTOMS_CODE PDF.
    Looks for values under "Cod. Arancel / Customs code" and strips dots.
    Returns a list of unique HS code strings (in order).
    """
    text = _read_pdf_text(pdf_path)
    if not text:
        return []

    # Find the section after the customs-code header
    header_match = re.search(
        r'Cod\.?\s*Arancel\s*/\s*Customs\s*code(.*)$',
        text, re.IGNORECASE | re.DOTALL
    )
    section = header_match.group(1) if header_match else text

    # HS codes look like "6403.99.96" or "64039996" - 6-10 digits possibly
    # interleaved with dots. We accept patterns of digits+dots that resolve
    # to >=6 digits.
    candidates = re.findall(r'\b\d{2,5}(?:\.\d{1,5}){0,4}\b', section)
    hs_codes = []
    seen = set()
    for cand in candidates:
        digits = cand.replace('.', '')
        # HS codes are typically 6-10 digits
        if 6 <= len(digits) <= 12 and digits not in seen:
            seen.add(digits)
            hs_codes.append(digits)
    return hs_codes


def extract_cargo_description_from_subject(subject):
    """
    Return the word immediately before "CHINA" in the email subject.
    Case-insensitive. Returns None if not found.
    """
    if not subject:
        return None
    match = re.search(r'(\S+)\s+CHINA\b', subject, re.IGNORECASE)
    if match:
        return match.group(1)
    return None
