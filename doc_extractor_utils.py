import re
from typing import Any, Dict, Union

from bs4 import BeautifulSoup


def get_recommended_summary_table_json(
    recommendation_summary_table_html: str,
) -> Dict[str, Any]:
    """
    Convert the HTML recommendation summary table to structured JSON format.

    Args:
        recommendation_summary_table_html: HTML string containing the recommendation summary table

    Returns:
        Dictionary containing structured recommendation data with headers and rows
    """
    if (
        not recommendation_summary_table_html
        or not recommendation_summary_table_html.strip()
    ):
        return {"headers": [], "recommendations": [], "totals": {}}

    # Parse the HTML
    soup = BeautifulSoup(recommendation_summary_table_html, "html.parser")

    # Find the table
    table = soup.find("table")
    if not table:
        return {"headers": [], "recommendations": [], "totals": {}}

    # Get all rows
    rows = table.find_all("tr")
    if len(rows) < 2:  # Need at least header + 1 data row
        return {"headers": [], "recommendations": [], "totals": {}}

    # Extract headers from first row
    header_row = rows[0]
    headers = []
    header_cells = header_row.find_all("td")

    for cell in header_cells:
        # Get text and clean it up
        header_text = cell.get_text().strip()
        # Remove extra whitespace and newlines
        header_text = re.sub(r"\s+", " ", header_text)
        headers.append(header_text)

    # Create standardized field mapping
    field_mapping = {
        "AR No.": "ar_number",
        "Category": "category",
        "Description": "description",
        "Electricity Savings (kWh/yr)": "electricity_savings_kwh_per_year",
        "Energy Cost Savings ($/yr)": "energy_cost_savings_per_year",
        "Demand Savings (kW/yr)": "demand_savings_kw_per_year",
        "Demand Savings(kW/yr)": "demand_savings_kw_per_year",  # No space variant
        "Demand Cost Savings ($/yr)": "demand_cost_savings_per_year",
        "Demand Cost Savings($/yr)": "demand_cost_savings_per_year",  # No space variant
        "Admin Cost Savings ($/yr)": "admin_cost_savings_per_year",
        "Propane Savings (mmbtu/yr)": "propane_savings_mmbtu_per_year",
        "Propane Cost Saving ($/yr)": "propane_cost_savings_per_year",
        "Total Cost Savings ($/yr)": "total_cost_savings_per_year",
        "CO2 Reduction (Tons/yr)": "co2_reduction_tons_per_year",
        "Impl. Cost ($)": "implementation_cost",
        "Impl.Cost ($)": "implementation_cost",  # No space variant
        "Payback Period (yrs)": "payback_period_years",
        "Payback Period(yrs)": "payback_period_years",  # No space variant
    }

    # Map headers to standardized field names
    standardized_headers = []
    for header in headers:
        # Clean up header text for matching
        clean_header = re.sub(r"\s+", " ", header.strip())
        standardized_field = field_mapping.get(
            clean_header,
            clean_header.lower()
            .replace(" ", "_")
            .replace("(", "")
            .replace(")", "")
            .replace("/", "_per_")
            .replace("$", "dollar")
            .replace(".", ""),
        )
        standardized_headers.append(standardized_field)

    def parse_numeric_value(value_str: str) -> Union[float, int, str]:
        """Parse a value that could be numeric or text."""
        if not value_str or value_str.strip() in {"-", "", "0"}:
            return 0

        # Remove commas and extra spaces
        clean_str = re.sub(r"[,\s]", "", value_str.strip())

        # Try to convert to number
        try:
            # Check if it's an integer
            if "." not in clean_str and clean_str.lstrip("-").isdigit():
                return int(clean_str)
            else:
                return float(clean_str)
        except ValueError:
            # Return as string if not numeric
            return value_str.strip()

    # Process data rows (skip header and totals row)
    recommendations = []
    totals = {}

    for row in rows[1:]:  # Skip header row
        cells = row.find_all("td")
        if len(cells) != len(standardized_headers):
            continue

        row_data = {}
        is_totals_row = False

        for j, cell in enumerate(cells):
            cell_text = cell.get_text().strip()
            field_name = standardized_headers[j]

            # Check if this is the totals row
            if "TOTALS" in cell_text.upper():
                is_totals_row = True
                continue

            # Parse the value
            if field_name in ["ar_number", "category", "description"]:
                # Keep as string for these fields
                value = cell_text
            else:
                # Parse as numeric for other fields
                value = parse_numeric_value(cell_text)

            row_data[field_name] = value

        # Add to appropriate collection
        if is_totals_row:
            totals = row_data
        elif row_data.get("ar_number"):  # Only add if has AR number
            recommendations.append(row_data)

    return {
        "headers": headers,
        "standardized_headers": standardized_headers,
        "recommendations": recommendations,
        "totals": totals,
    }


def validate_recommendation_totals(
    recommendation_json: Dict[str, Any], tolerance: float = 0.01
) -> Dict[str, Any]:
    """
    Validate that the sum of individual recommendations matches the totals row.

    Args:
        recommendation_json: Output from get_recommended_summary_table_json function
        tolerance: Acceptable difference for floating point comparisons (default: 0.01)

    Returns:
        Dictionary containing validation results for each numeric column
    """
    if not recommendation_json.get("recommendations") or not recommendation_json.get(
        "totals"
    ):
        return {"error": "Missing recommendations or totals data"}

    recommendations = recommendation_json["recommendations"]
    totals = recommendation_json["totals"]

    # Get all numeric fields (exclude text fields)
    text_fields = {"ar_number", "category", "description"}
    numeric_fields = [
        field
        for field in recommendation_json.get("standardized_headers", [])
        if field not in text_fields
    ]

    # Fields that should NOT be summed (calculated differently)
    non_summable_fields = {
        "payback_period_years"  # Usually weighted average or total_cost/total_savings
    }

    validation_results = []

    for field in numeric_fields:
        # Calculate sum from individual recommendations
        calculated_sum = 0
        valid_values = []

        for rec in recommendations:
            value = rec.get(field, 0)
            if isinstance(value, (int, float)):
                calculated_sum += value
                valid_values.append(value)

        # Get expected total from totals row
        expected_total = totals.get(field, 0)

        # Special handling for non-summable fields
        if field in non_summable_fields:
            validation_results.append(
                {
                    "field_name": field,
                    "calculated_sum": calculated_sum,
                    "expected_total": expected_total,
                    "difference": (
                        abs(calculated_sum - expected_total)
                        if isinstance(expected_total, (int, float))
                        else None
                    ),
                    "is_valid": "not_applicable",
                    "individual_values": valid_values,
                    "note": f"Field '{field}' is not expected to be a simple sum of individual values",
                    "validation_type": "non_summable",
                }
            )
            continue

        # Compare values for summable fields
        if isinstance(expected_total, (int, float)) and isinstance(
            calculated_sum, (int, float)
        ):
            difference = abs(calculated_sum - expected_total)
            is_valid = difference <= tolerance

            validation_results.append(
                {
                    "field_name": field,
                    "calculated_sum": calculated_sum,
                    "expected_total": expected_total,
                    "difference": difference,
                    "is_valid": is_valid,
                    "individual_values": valid_values,
                    "tolerance_used": tolerance,
                    "validation_type": "summable",
                }
            )
        else:
            validation_results.append(
                {
                    "field_name": field,
                    "calculated_sum": calculated_sum,
                    "expected_total": expected_total,
                    "difference": None,
                    "is_valid": False,
                    "individual_values": valid_values,
                    "error": "Non-numeric values found",
                    "validation_type": "error",
                }
            )

    # Overall validation summary (only consider summable fields)
    summable_results = [
        result
        for result in validation_results
        if result.get("validation_type") == "summable"
    ]
    all_valid = all(result.get("is_valid", False) for result in summable_results)
    invalid_fields = [
        result["field_name"]
        for result in summable_results
        if not result.get("is_valid", False)
    ]
    non_summable_count = len(
        [
            result
            for result in validation_results
            if result.get("validation_type") == "non_summable"
        ]
    )

    return {
        "overall_valid": all_valid,
        "invalid_fields": invalid_fields,
        "field_validations": validation_results,
        "summary": {
            "total_fields_checked": len(validation_results),
            "summable_fields_checked": len(summable_results),
            "non_summable_fields": non_summable_count,
            "valid_summable_fields": len(summable_results) - len(invalid_fields),
            "invalid_summable_fields": len(invalid_fields),
        },
    }
