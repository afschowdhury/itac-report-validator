from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
import re
from typing import Iterable, List, Union, Optional, Dict, Any
import json

DOCX_PATH = "/Users/afschowdhury/Code Local/itac-report-validator/docs/LS2502 - Final Draft R2.docx"

# ---------- Low-level helpers ----------

def iter_block_items(doc: Document) -> Iterable[Union[Paragraph, Table]]:
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)

def escape_html(text: str) -> str:
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
    )

def para_alignment_name(p: Paragraph) -> str:
    if p.alignment == 1:
        return "center"
    if p.alignment == 2:
        return "right"
    return "left"

# ---------- Renderers: HTML ----------

def paragraph_to_html(p: Paragraph) -> str:
    if not p.runs:
        return "<p></p>"
    parts = []
    for r in p.runs:
        t = escape_html(r.text)
        if not t:
            continue
        if r.bold:
            t = f"<b>{t}</b>"
        if r.italic:
            t = f"<i>{t}</i>"
        parts.append(t)
    align = para_alignment_name(p)
    style = f' style="text-align:{align}"' if align != "left" else ""
    return f"<p{style}>" + "".join(parts) + "</p>"

def table_to_html(tbl: Table) -> str:
    rows_html = []
    for row in tbl.rows:
        cells_html = []
        for cell in row.cells:
            cell_html = "".join(paragraph_to_html(p) for p in cell.paragraphs)
            cells_html.append(f"<td>{cell_html}</td>")
        rows_html.append("<tr>" + "".join(cells_html) + "</tr>")
    return "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;width:100%'>" + "".join(rows_html) + "</table>"

def blocks_to_html(blocks: List[Union[Paragraph, Table]]) -> str:
    html_parts = []
    for b in blocks:
        if isinstance(b, Paragraph):
            html_parts.append(paragraph_to_html(b))
        elif isinstance(b, Table):
            html_parts.append(table_to_html(b))
    return "\n".join(html_parts)

# ---------- Renderers: JSON ----------

def paragraph_to_json(p: Paragraph) -> Dict[str, Any]:
    runs = []
    for r in p.runs:
        if r.text:
            runs.append({
                "text": r.text,
                "bold": bool(r.bold),
                "italic": bool(r.italic),
            })
    return {
        "type": "paragraph",
        "alignment": para_alignment_name(p),
        "runs": runs
    }

def table_to_json(tbl: Table) -> Dict[str, Any]:
    grid = []
    for row in tbl.rows:
        row_cells = []
        for cell in row.cells:
            row_cells.append({
                "paragraphs": [paragraph_to_json(p) for p in cell.paragraphs]
            })
        grid.append(row_cells)
    return {
        "type": "table",
        "rows": grid
    }

def blocks_to_json(blocks: List[Union[Paragraph, Table]]) -> List[Dict[str, Any]]:
    out = []
    for b in blocks:
        if isinstance(b, Paragraph):
            out.append(paragraph_to_json(b))
        elif isinstance(b, Table):
            out.append(table_to_json(b))
    return out

# ---------- Finders for sections and tables ----------

def normalize(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()

def is_title(p: Paragraph, pattern: re.Pattern) -> bool:
    return bool(pattern.match(normalize(p.text)))

def slice_blocks_between(doc_blocks: List[Union[Paragraph, Table]],
                         start_idx: int,
                         end_idx: Optional[int]) -> List[Union[Paragraph, Table]]:
    return doc_blocks[start_idx:end_idx] if end_idx is not None else doc_blocks[start_idx:]

def find_section_index(doc_blocks: List[Union[Paragraph, Table]], title_regex: str) -> Optional[int]:
    pat = re.compile(title_regex, flags=re.IGNORECASE)
    for i, b in enumerate(doc_blocks):
        if isinstance(b, Paragraph) and is_title(b, pat):
            return i
    return None

def find_next_section_start(doc_blocks: List[Union[Paragraph, Table]], from_idx: int, stop_regex: str) -> Optional[int]:
    pat = re.compile(stop_regex, flags=re.IGNORECASE)
    for i in range(from_idx + 1, len(doc_blocks)):
        b = doc_blocks[i]
        if isinstance(b, Paragraph) and is_title(b, pat):
            return i
    return None

def extract_section_by_title(doc_blocks: List[Union[Paragraph, Table]],
                             title_regex: str,
                             next_titles_regex: List[str]) -> List[Union[Paragraph, Table]]:
    start = find_section_index(doc_blocks, title_regex)
    if start is None:
        return []
    end_candidates = []
    for nxt in next_titles_regex:
        idx = find_next_section_start(doc_blocks, start, nxt)
        if idx is not None:
            end_candidates.append(idx)
    end = min(end_candidates) if end_candidates else None
    return slice_blocks_between(doc_blocks, start, end)

def find_table_by_caption(doc_blocks: List[Union[Paragraph, Table]],
                          caption_patterns: List[str]) -> Optional[Table]:
    pats = [re.compile(p, flags=re.IGNORECASE) for p in caption_patterns]
    matches = []
    for i, b in enumerate(doc_blocks):
        if isinstance(b, Paragraph):
            text = normalize(b.text)
            if any(pat.match(text) for pat in pats):
                for j in range(i + 1, len(doc_blocks)):
                    if isinstance(doc_blocks[j], Table):
                        matches.append(doc_blocks[j])
                        break
    # Return the last match (likely the actual table, not from table of contents)
    return matches[-1] if matches else None

def extract_ars(doc_blocks: List[Union[Paragraph, Table]]) -> List[List[Union[Paragraph, Table]]]:
    # Updated pattern to match actual format: "AR No. 1 – ..." (without section numbers)
    ar_title_pat = re.compile(r"^\s*AR\s+No\.\s*\d+\b", flags=re.IGNORECASE)
    ar_starts: List[int] = []
    for i, b in enumerate(doc_blocks):
        if isinstance(b, Paragraph) and ar_title_pat.match(normalize(b.text)):
            ar_starts.append(i)

    results: List[List[Union[Paragraph, Table]]] = []
    for k, start in enumerate(ar_starts):
        next_ar = ar_starts[k + 1] if k + 1 < len(ar_starts) else None
        # Look for next major section after ARs (could be various patterns)
        next_major = find_next_section_start(doc_blocks, start, r"^\s*(5(\.|$)|INDUSTRIAL\s+CONTROL|CONCLUSIONS?|REFERENCES?|APPENDIX)")
        end_candidates = [idx for idx in [next_ar, next_major] if idx is not None]
        end = min(end_candidates) if end_candidates else None
        results.append(slice_blocks_between(doc_blocks, start, end))
    return results

# ---------- Main extraction with output switch ----------

def build_outputs(blocks: List[Union[Paragraph, Table]], output: str) -> Dict[str, Any]:
    # Sections - Updated patterns to match actual document structure
    sec_11 = extract_section_by_title(
        blocks,
        r"^\s*General\s+Information\b",
        [r"^\s*Annual\s+Energy\s+Usages\s+and\s+Costs\b", r"^\s*Carbon\s+Footprint\b", r"^\s*Summary\s+of\s+Best\s+Practices"]
    )
    sec_12 = extract_section_by_title(
        blocks,
        r"^\s*Annual\s+Energy\s+Usages\s+and\s+Costs\b",
        [r"^\s*Carbon\s+Footprint\b", r"^\s*Summary\s+of\s+Best\s+Practices"]
    )
    sec_13 = extract_section_by_title(
        blocks,
        r"^\s*Carbon\s+Footprint\b",
        [r"^\s*Summary\s+of\s+Best\s+Practices", r"^\s*COMPANY\s+BACKGROUND"]
    )

    # Table 1.3/1-3 caption (allow minor caption text variation)
    rec_tbl = find_table_by_caption(
        blocks,
        [
            r"^\s*Table\s*1[.-]3\b.*Recommendation Summary Table",
            r"^\s*Table\s*1[.-]3\b.*Assessment Recommendation Summary Table",
        ],
    )

    # ARs
    ar_blocks_list = extract_ars(blocks)

    if output == "json":
        return {
            "general_information": blocks_to_json(sec_11),
            "annual_energy_usages_and_costs": blocks_to_json(sec_12),
            "carbon_footprint": blocks_to_json(sec_13),
            "recommendation_summary_table": (table_to_json(rec_tbl) if rec_tbl else None),
            "assessment_recommendations": [blocks_to_json(b) for b in ar_blocks_list],
        }
    # default: HTML
    return {
        "general_information": blocks_to_html(sec_11),
        "annual_energy_usages_and_costs": blocks_to_html(sec_12),
        "carbon_footprint": blocks_to_html(sec_13),
        "recommendation_summary_table": (table_to_html(rec_tbl) if rec_tbl else ""),
        "assessment_recommendations": [blocks_to_html(b) for b in ar_blocks_list],
    }

def write_artifacts(payload: Dict[str, Any], output: str) -> None:
    """
    Save files to disk for inspection.
    """
    if output == "json":
        with open("extracted_sections.json", "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    else:
        def write(name: str, content: str):
            with open(name, "w", encoding="utf-8") as f:
                f.write(content)

        write("general_information.html", payload["general_information"])
        write("annual_energy_usages_and_costs.html", payload["annual_energy_usages_and_costs"])
        write("carbon_footprint.html", payload["carbon_footprint"])
        write("recommendation_summary_table.html", payload["recommendation_summary_table"])
        for i, html in enumerate(payload["assessment_recommendations"], start=1):
            write(f"AR_{i:02d}.html", html)

def extract_itac_report(docx_path: str = DOCX_PATH, output: str = "html", save_files: bool = True) -> Dict[str, Any]:
    """
    output: "html" or "json"
    save_files: write artifacts to disk if True
    """
    assert output in {"html", "json"}, "output must be 'html' or 'json'"
    doc = Document(docx_path)
    blocks = list(iter_block_items(doc))
    data = build_outputs(blocks, output=output)
    if save_files:
        write_artifacts(data, output=output)
    return data

def extract_general_info_fields(general_info_html: str) -> Dict[str, Union[str, float]]:
    """
    Extract specific fields from the general information HTML table.
    
    Args:
        general_info_html: HTML string containing the general information table
        
    Returns:
        Dictionary with extracted field names as keys and their values as values.
        All values are converted to float except principal_product which remains as string.
    """
    import re
    from bs4 import BeautifulSoup
    
    def extract_numeric_value(value_str: str) -> float:
        """Extract numeric value from a string, handling millions/billions and removing units, currency symbols, and commas."""
        # Convert to lowercase for easier matching
        value_lower = value_str.lower()
        
        # Remove common currency symbols and commas
        clean_str = re.sub(r'[$,]', '', value_str)
        
        # Find all numbers (including decimals) in the string
        numbers = re.findall(r'\d+\.?\d*', clean_str)
        if not numbers:
            return 0.0
            
        base_number = float(numbers[0])
        
        # Check for scale multipliers
        if 'billion' in value_lower or ('b' in value_lower and not 'mb' in value_lower):
            return base_number * 1_000_000_000
        elif 'million' in value_lower or ('m' in value_lower and not 'mb' in value_lower):
            return base_number * 1_000_000
        elif 'thousand' in value_lower or 'k' in value_lower:
            return base_number * 1_000
        else:
            return base_number
    
    # Parse the HTML
    soup = BeautifulSoup(general_info_html, 'html.parser')
    
    # Initialize the result dictionary
    extracted_fields = {}
    
    # Find the table containing the general information
    table = soup.find('table')
    if not table:
        return extracted_fields
    
    # Extract data from table rows
    for row in table.find_all('tr'):
        cells = row.find_all('td')
        for cell in cells:
            cell_text = cell.get_text(strip=True)
            if ':' in cell_text:
                # Split on the first colon to separate field name and value
                parts = cell_text.split(':', 1)
                if len(parts) == 2:
                    field_name = parts[0].strip()
                    field_value = parts[1].strip()
                    
                    # Normalize field names to consistent keys
                    field_mapping = {
                        'SIC. No.': 'sic_no',
                        'SIC No.': 'sic_no',
                        'SIC No': 'sic_no',
                        'NAICS Code': 'naics_code',
                        'Principal Product': 'principal_product',
                        'Principal Products': 'principal_products',
                        'No. of Employees': 'no_of_employees',
                        'Number of Employees': 'no_of_employees',
                        'Total Facility Area': 'total_facility_area',
                        'Operating Hours': 'operating_hours',
                        'Annual Production': 'annual_production',
                        'Annual Sales': 'annual_sales',
                        'Value per Finished Product': 'value_per_finished_product',
                        'Total Energy Usage': 'total_energy_usage',
                        'Total Utility Cost': 'total_utility_cost',
                        'No. of Assessment Recommendations': 'no_of_assessment_recommendations'
                    }
                    
                    # Map to standardized key or use original field name
                    standardized_key = field_mapping.get(field_name, field_name.lower().replace(' ', '_').replace('.', ''))
                    
                    # Convert to appropriate type
                    if standardized_key in ['principal_product', 'principal_products']:
                        extracted_fields[standardized_key] = field_value
                    else:
                        extracted_fields[standardized_key] = extract_numeric_value(field_value)
    
    return extracted_fields


def extract_energy_usage(annual_energy_html: str) -> Dict[str, Any]:
    """
    Extract energy usage data from the annual energy usages and costs HTML.
    
    Args:
        annual_energy_html: HTML string containing the annual energy usages and costs table
        
    Returns:
        Dictionary with period information and energy usage data
    """
    import re
    from bs4 import BeautifulSoup
    
    def extract_period_from_text(text: str) -> Dict[str, str]:
        """Extract start and end period from descriptive text."""
        # Look for patterns like "between September 2023 and August 2024"
        period_pattern = r'between\s+(\w+\s+\d{4})\s+and\s+(\w+\s+\d{4})'
        match = re.search(period_pattern, text, re.IGNORECASE)
        if match:
            return {"start": match.group(1), "end": match.group(2)}
        
        # Alternative pattern: "from X to Y" or "X - Y"
        alt_pattern = r'(?:from\s+)?(\w+\s+\d{4})(?:\s+(?:to|-)\s+)(\w+\s+\d{4})'
        match = re.search(alt_pattern, text, re.IGNORECASE)
        if match:
            return {"start": match.group(1), "end": match.group(2)}
            
        return {"start": "", "end": ""}
    
    def parse_usage_cell(usage_text: str) -> Dict[str, float]:
        """Parse usage cell that may contain multiple values with different units."""
        usage_dict = {}
        
        # Find all patterns like "649,680 kWh/yr" or "(2,217 MMBTU/yr)"
        patterns = re.findall(r'[\(]?([0-9,]+\.?[0-9]*)\s+([A-Za-z/]+)[\)]?', usage_text)
        
        for value_str, unit in patterns:
            # Clean up the value string and convert to float
            clean_value = re.sub(r'[,\s]', '', value_str)
            try:
                value = float(clean_value)
                usage_dict[unit] = value
            except ValueError:
                continue
                
        return usage_dict
    
    def parse_cost_cell(cost_text: str) -> float:
        """Parse cost cell and return numeric value."""
        # Remove currency symbols, commas, and /yr
        clean_cost = re.sub(r'[\$,/yr\s]', '', cost_text)
        # Extract numeric value
        numbers = re.findall(r'\d+\.?\d*', clean_cost)
        if numbers:
            return float(numbers[0])
        return 0.0
    
    def parse_unit_cost_cell(unit_cost_text: str) -> Optional[Dict[str, Union[float, str]]]:
        """Parse unit cost cell like '$0.102/kWh' or '$4.522/kW'."""
        if unit_cost_text.strip() in ['-', '']:
            return None
            
        # Pattern for $X.XX/unit
        pattern = r'\$([0-9,]+\.?[0-9]*)/([A-Za-z]+)'
        match = re.search(pattern, unit_cost_text)
        if match:
            amount = float(re.sub(r'[,]', '', match.group(1)))
            unit = match.group(2)
            return {"amount": amount, "unit": unit}
        return None
    
    # Parse the HTML
    soup = BeautifulSoup(annual_energy_html, 'html.parser')
    
    # Initialize result structure
    result = {
        "period": {"start": "", "end": ""},
        "data": []
    }
    
    # Extract period information from paragraph text
    paragraphs = soup.find_all('p')
    for p in paragraphs:
        text = p.get_text()
        period_info = extract_period_from_text(text)
        if period_info["start"] and period_info["end"]:
            result["period"] = period_info
            break
    
    # Find the energy usage table
    table = soup.find('table')
    if not table:
        return result
    
    # Process table rows (skip header)
    rows = table.find_all('tr')
    if len(rows) <= 1:  # No data rows
        return result
    
    for row in rows[1:]:  # Skip header row
        cells = row.find_all('td')
        if len(cells) < 4:  # Should have Type, Usage, Cost, Unit Cost
            continue
            
        # Extract data from each cell
        energy_type_raw = cells[0].get_text(strip=True).replace('**', '').strip()
        usage_text = cells[1].get_text()
        cost_text = cells[2].get_text()
        unit_cost_text = cells[3].get_text()
        
        # Map energy types to programming-oriented field names
        type_mapping = {
            'Electrical Energy': 'electrical_energy',
            'Electrical Demand': 'electrical_demand',
            'Electric Energy': 'electrical_energy',
            'Electric Demand': 'electrical_demand',
            'Electricity': 'electrical_energy',
            'Demand Charge': 'electrical_demand', # TODO: Verify this is correct
            'Demand': 'electrical_demand',
            'Natural Gas': 'natural_gas',
            'Propane': 'propane_gas',
            'Propane Gas': 'propane_gas',
            'Steam': 'steam',
            'Water': 'water',
            'Compressed Air': 'compressed_air',
            'Total Utility': 'total_utility',
            'TotalUtility': 'total_utility',
            'Total': 'total_utility',
            'Fuel Oil': 'fuel_oil',
            'Heating Oil': 'heating_oil',
            'Diesel': 'diesel',
            'Gasoline': 'gasoline',
            'Coal': 'coal',
            'Biomass': 'biomass',
            'Solar': 'solar',
            'Wind': 'wind',
            'Geothermal': 'geothermal',
            'Chilled Water': 'chilled_water',
            'Hot Water': 'hot_water'
        }
        
        # Get standardized type name or create one from raw name
        energy_type = type_mapping.get(energy_type_raw, 
                                     energy_type_raw.lower()
                                     .replace(' ', '_')
                                     .replace('-', '_')
                                     .replace('&', 'and')
                                     .replace('/', '_'))
        
        # Parse the data
        usage_data = parse_usage_cell(usage_text)
        cost_value = parse_cost_cell(cost_text)
        unit_cost_data = parse_unit_cost_cell(unit_cost_text)
        
        # Create entry
        entry = {
            "type": energy_type,
            "usage": usage_data,
            "cost": cost_value,
            "unit_cost": unit_cost_data
        }
        
        result["data"].append(entry)
    
    return result


if __name__ == "__main__":
    # HTML run
    html_out = extract_itac_report(DOCX_PATH, output="html", save_files=True)
    print("HTML extraction complete.",
          len(html_out["assessment_recommendations"]), "AR sections")

    # JSON run
    json_out = extract_itac_report(DOCX_PATH, output="json", save_files=True)
    print("JSON extraction complete.",
          len(json_out["assessment_recommendations"]), "AR sections")
    
    # Extract general information fields
    general_info_fields = extract_general_info_fields(html_out["general_information"])
    print("\nExtracted General Information Fields:")
    for key, value in general_info_fields.items():
        print(f"  {key}: {value}")
    
    # Extract energy usage data
    energy_usage_data = extract_energy_usage(html_out["annual_energy_usages_and_costs"])
    print(f"\nExtracted Energy Usage Data:")
    print(f"  Period: {energy_usage_data['period']['start']} to {energy_usage_data['period']['end']}")
    print(f"  Number of energy types: {len(energy_usage_data['data'])}")
    for item in energy_usage_data['data']:
        print(f"    - {item['type']}: {item['usage']} (Cost: ${item['cost']:.2f})")
