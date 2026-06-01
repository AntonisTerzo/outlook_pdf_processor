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
#
# Task 3 does NOT download/save the PDFs. The Outlook attachment is read
# straight into memory as bytes, so every reader below accepts raw bytes
# rather than a filesystem path.
#
# Only two file types matter: PACKING_LIST and CUSTOMS_CODE. The invoice
# file is ignored entirely. Everything we need (including the IV value) is
# read from the packing list and customs-code documents.
# ============================================================

import io


def _read_pdf_text_from_bytes(pdf_bytes):
    """Read full text from PDF bytes. Returns string (empty on error)."""
    try:
        pdf_reader = PdfReader(io.BytesIO(pdf_bytes))
        return "\n".join(page.extract_text() or "" for page in pdf_reader.pages)
    except Exception as e:
        print(f"Error reading PDF text from bytes: {e}")
        return ""


def detect_file_type_task3(pdf_bytes, file_types_map):
    """
    Detect which Task 3 file type a PDF is, from raw bytes.
    file_types_map: dict like {"PACKING_LIST": [...identifiers...], ...}
    Returns the type key or None if it is neither (e.g. the invoice file,
    which we deliberately do not match).
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
    integer number of centimetres. Examples: '1,20' -> 120, '1.20' -> 120,
    '0.8' -> 80.
    """
    s = value_str.strip()
    # Normalise decimal separator: if both '.' and ',' present, assume '.' is
    # thousands and ',' is decimal. Otherwise treat ',' as decimal.
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    val = float(s)
    return int(round(val * 100))


def extract_packing_list_data(pdf_bytes):
    """
    Extract all needed fields from a PACKING_LIST PDF (given as bytes).

    Returns dict with keys:
        descriptions (list[str]), iv, srn, pcs, kg, m3, dims

    Parsing approach: the PDF text is split into lines. The cargo descriptions
    live in a table column headed "Descripción", and that header always sits
    before "Cod. Modelo / Model id" on the same header row. We capture the
    description column values between those anchors.
    """
    text = _read_pdf_text_from_bytes(pdf_bytes)
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

    lines = [ln.rstrip() for ln in text.split('\n')]

    # --- IV: the value after "PO Customer" ---
    # Look on the same line first, then fall back to the next non-empty line.
    for i, line in enumerate(lines):
        m = re.search(r'PO\s*Customer[^\w\d]*([A-Z0-9\-_/]+)', line, re.IGNORECASE)
        if m and m.group(1).strip():
            result["iv"] = m.group(1).strip()
            break
        if re.search(r'PO\s*Customer', line, re.IGNORECASE):
            # Value may be on the following line(s)
            for j in range(i + 1, min(i + 4, len(lines))):
                nxt = lines[j].strip()
                if nxt:
                    token = re.match(r'([A-Z0-9\-_/]+)', nxt, re.IGNORECASE)
                    if token:
                        result["iv"] = token.group(1).strip()
                    break
            break

    # --- SRN: after "Nº Envío / Shipment Nr" ---
    srn_match = re.search(
        r'Env[ií]o\s*/\s*Shipment\s*Nr\s*[\.:]*\s*([A-Z0-9][A-Z0-9\-_/]*)',
        text, re.IGNORECASE
    )
    if srn_match:
        result["srn"] = srn_match.group(1).strip()

    # --- Descriptions: column under "Descripción", which precedes
    #     "Cod. Modelo / Model id" on the header row. ---
    result["descriptions"] = _extract_descriptions(lines)

    # --- Summary section (RESUMEN / SUMMARY) for PCS / KG / M3 ---
    summary_match = re.search(
        r'RESUMEN\s*/\s*SUMMARY(.*)$', text, re.IGNORECASE | re.DOTALL
    )
    summary_text = summary_match.group(1) if summary_match else text

    pcs_match = re.search(
        r'Total\s*N[ºoª°]?\s*Parcels\s*[:.]?\s*(\d[\d\.,]*)',
        summary_text, re.IGNORECASE
    )
    if not pcs_match:
        # Fallback: Spanish label "Bultos"
        pcs_match = re.search(
            r'Total\s*de\s*Bultos[^\n]*?(\d[\d\.,]*)',
            summary_text, re.IGNORECASE
        )
    if pcs_match:
        result["pcs"] = pcs_match.group(1).strip()

    kg_match = re.search(
        r'Gross\s*Weight\s*:?\s*\(?\s*Kgs?\s*\)?\s*[:.]?\s*(\d[\d\.,]*)',
        summary_text, re.IGNORECASE
    )
    if not kg_match:
        kg_match = re.search(
            r'Peso\s*Bruto\s*Total[^\n]*?(\d[\d\.,]*)',
            summary_text, re.IGNORECASE
        )
    if kg_match:
        result["kg"] = kg_match.group(1).strip()

    m3_match = re.search(
        r'Total\s*Volume\s*:?\s*\(?\s*m3\s*\)?\s*[:.]?\s*(\d[\d\.,]*)',
        summary_text, re.IGNORECASE
    )
    if not m3_match:
        m3_match = re.search(
            r'Volumen\s*Total[^\n]*?(\d[\d\.,]*)',
            summary_text, re.IGNORECASE
        )
    if m3_match:
        result["m3"] = m3_match.group(1).strip()

    # --- DIMS: from Long./Length (m), Anchura/Width (m), Altura/Height (m) ---
    dims_match = re.search(
        r'Long\.?\s*/\s*Length.*?Anchura\s*/\s*Width.*?Altura\s*/\s*Height[^\d]*'
        r'([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)',
        text, re.IGNORECASE | re.DOTALL
    )
    if dims_match:
        try:
            l = _to_cm(dims_match.group(1))
            w = _to_cm(dims_match.group(2))
            h = _to_cm(dims_match.group(3))
            result["dims"] = f"{l}x{w}x{h}"
        except Exception as e:
            print(f"Error parsing dims: {e}")

    return result


def _extract_descriptions(lines):
    """
    Find the descriptions column in the packing-list table.

    The table header contains "Descripción" immediately before
    "Cod. Modelo / Model id". We locate the header line, then read each
    subsequent data row until the table ends (RESUMEN / SUMMARY, or a
    clearly non-table line). On each data row we take the leading text up to
    the model-id token as the description.

    Returns a list of unique descriptions, comma-joining handled by caller.
    """
    descriptions = []
    seen = set()

    # Locate the header row: "Descripción ... Cod. Modelo / Model id"
    header_idx = None
    for i, line in enumerate(lines):
        if re.search(r'Descripci[oó]n', line, re.IGNORECASE) and \
           re.search(r'Cod\.?\s*Modelo\s*/\s*Model\s*id', line, re.IGNORECASE):
            header_idx = i
            break
    # Fallback: header where "Descripción" appears and "Model id" is on the
    # same or the very next line.
    if header_idx is None:
        for i, line in enumerate(lines):
            if re.search(r'Descripci[oó]n', line, re.IGNORECASE):
                joined = line + " " + (lines[i + 1] if i + 1 < len(lines) else "")
                if re.search(r'Cod\.?\s*Modelo\s*/\s*Model\s*id', joined, re.IGNORECASE):
                    header_idx = i
                    break
    if header_idx is None:
        return descriptions

    # Read data rows after the header
    for line in lines[header_idx + 1:]:
        stripped = line.strip()
        if not stripped:
            continue
        # Stop at the summary / end of table
        if re.search(r'RESUMEN\s*/\s*SUMMARY', stripped, re.IGNORECASE):
            break
        # Skip pure-number / dimension rows
        if re.fullmatch(r'[\d\s\.\,xX/-]+', stripped):
            continue
        # Skip obvious non-description footer/labels
        if re.search(r'(Long\.?\s*/\s*Length|Anchura|Altura|Peso\s*Bruto|'
                     r'Volumen|Bultos|Parcels|N[ºo°]?\s*Env[ií]o|PO\s*Customer)',
                     stripped, re.IGNORECASE):
            continue

        # The description is the leading text of the row, sitting in the
        # column before "Cod. Modelo / Model id". Model ids contain a digit
        # (e.g. MOD-1234, AB12, 998877). Cut the row at the first token that
        # looks like a code/quantity (contains a digit), keeping the leading
        # all-text portion as the description. Purely alphabetic multi-word
        # descriptions (e.g. "STEEL BOLTS") are preserved in full.
        desc = stripped
        cut = re.search(r'\s+\S*\d\S*', stripped)
        if cut:
            candidate = stripped[:cut.start()].strip()
            if len(candidate) >= 3 and re.search(r'[A-Za-zÀ-ÿ]', candidate):
                desc = candidate

        if re.search(r'[A-Za-zÀ-ÿ]{3,}', desc) and desc not in seen:
            seen.add(desc)
            descriptions.append(desc)

    return descriptions


def extract_hs_codes(pdf_bytes):
    """
    Extract HS codes from a CUSTOMS_CODE PDF (given as bytes).
    Looks for values under "Cod. Arancel / Customs code" and strips dots.
    Returns a list of unique HS code strings, in order of appearance.
    """
    text = _read_pdf_text_from_bytes(pdf_bytes)
    if not text:
        return []

    header_match = re.search(
        r'Cod\.?\s*Arancel\s*/\s*Customs\s*code(.*)$',
        text, re.IGNORECASE | re.DOTALL
    )
    section = header_match.group(1) if header_match else text

    candidates = re.findall(r'\b\d{2,5}(?:\.\d{1,5}){0,4}\b', section)
    hs_codes = []
    seen = set()
    for cand in candidates:
        digits = cand.replace('.', '')
        if 6 <= len(digits) <= 12 and digits not in seen:
            seen.add(digits)
            hs_codes.append(digits)
    return hs_codes


def extract_cargo_description_from_subject(subject):
    """
    Return the word that sits between dashes and immediately before CHINA in
    the subject. Example: "cheese going to the - world - CHINA" -> "world".

    Strategy: find a "- <text> -" segment that is immediately followed by
    CHINA, and return the last word of that segment. Falls back to the word
    directly before CHINA if no dash pattern is found.
    """
    if not subject:
        return None

    # Primary: "- something - CHINA"  -> last word inside the dashes
    m = re.search(r'-\s*([^-]+?)\s*-\s*CHINA\b', subject, re.IGNORECASE)
    if m:
        inside = m.group(1).strip()
        if inside:
            return inside.split()[-1]

    # Fallback: plain word before CHINA
    m2 = re.search(r'(\S+)\s+CHINA\b', subject, re.IGNORECASE)
    if m2:
        return m2.group(1)

    return None
