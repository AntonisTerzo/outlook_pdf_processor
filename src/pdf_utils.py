import re
from pypdf import PdfReader


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
                r'Warenempfänger:(.*?)(?=Lieferkondition:|$)', text, re.DOTALL)

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


def extract_city_from_filename_task2(filename, cities_list):
    """
    Extract city name from PDF filename for Task 2.
    Searches the filename for any city from the cities_list.
    IMPORTANT: Checks longer names FIRST to avoid partial matches.
    Example: "BRASILIEN JOINVILLE" must be checked before "BRASILIEN"
    Returns the city name if found, None otherwise.
    """
    try:
        # Remove .pdf extension
        name_without_ext = filename.replace('.pdf', '').replace('.PDF', '')

        # Sort cities by length (longest first) to match specific names before generic ones
        sorted_cities = sorted(cities_list, key=len, reverse=True)

        # Search for each city in the filename (case-insensitive)
        for city_name in sorted_cities:
            if re.search(re.escape(city_name), name_without_ext, re.IGNORECASE):
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
    Extract dimensions from PDF under "Abmessung(MM)" column.
    Returns list of dimension strings in format "LxWxH" (in centimeters, rounded).
    Example: ["68x36x47", "55x30x40", "1x2x3"]
    """
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text()

            dimensions = []

            # Pattern to find dimensions in format: 680x360x470 or 10x20x30 (2-4 digits each)
            dimension_pattern = r'(\d{2,4})[xX](\d{2,4})[xX](\d{2,4})'

            for match in re.finditer(dimension_pattern, text):
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

            return dimensions

    except Exception as e:
        print(f"Error extracting dimensions from PDF: {e}")
        return []
