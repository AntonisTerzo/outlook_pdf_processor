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
                    # Check if we've hit another section header - stop tracking this Abmessung
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
