#!/usr/bin/env python3
"""
ITAC Report Validator Web Application

A Flask web application for uploading and comparing DOCX and Excel ITAC reports.
Extracts data using document_extractor.py and excel_keyinfo_extractor.py and 
highlights mismatches between the two sources.
"""

import logging
import os
import tempfile
from pathlib import Path
from typing import Any, Dict, List

import tomli
from flask import Flask, flash, jsonify, redirect, render_template, request, url_for
from icecream import ic
from werkzeug.utils import secure_filename

ic.configureOutput(includeContext=True, prefix='DEBUG: ')



# Import our existing extractors
from doc_extractor_utils import (
    compare_ar_with_summary,
    get_recommended_summary_table_json,
    get_single_ar_summary_table,
    validate_recommendation_totals,
)
from document_extractor import (
    extract_energy_usage,
    extract_general_info_fields,
    extract_itac_report,
)
from excel_keyinfo_extractor import extract_all_structured_info
from link_validator import validate_all_links

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

app = Flask(__name__)
app.secret_key = 'itac-validator-secret-key-2024'  # Change this in production
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

# Configuration
UPLOAD_FOLDER = Path('uploads')
UPLOAD_FOLDER.mkdir(exist_ok=True)
ALLOWED_EXTENSIONS = {'docx', 'xlsx'}

def allowed_file(filename: str) -> bool:
    """Check if file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def compare_values(doc_value: Any, excel_value: Any, tolerance: float = 0.01) -> Dict[str, Any]:
    """
    Compare two values and return comparison result with mismatch detection.
    
    Args:
        doc_value: Value from document extraction
        excel_value: Value from Excel extraction
        tolerance: Relative tolerance for numeric comparisons (1% default)
    
    Returns:
        Dict with comparison results including mismatch status
    """
    result = {
        'doc_value': doc_value,
        'excel_value': excel_value,
        'match': False,
        'mismatch_type': None,
        'difference': None,
        'formatted_doc': str(doc_value) if doc_value is not None else 'N/A',
        'formatted_excel': str(excel_value) if excel_value is not None else 'N/A'
    }
    
    # Handle None values
    if doc_value is None and excel_value is None:
        result['match'] = True
        result['formatted_doc'] = result['formatted_excel'] = 'N/A'
        return result
    
    if doc_value is None or excel_value is None:
        result['mismatch_type'] = 'missing_value'
        return result
    
    # Handle string comparisons
    if isinstance(doc_value, str) or isinstance(excel_value, str):
        doc_str = str(doc_value).strip().lower()
        excel_str = str(excel_value).strip().lower()
        result['match'] = doc_str == excel_str
        if not result['match']:
            result['mismatch_type'] = 'text_mismatch'
        return result
    
    # Handle numeric comparisons
    try:
        doc_num = float(doc_value)
        excel_num = float(excel_value)
        
        # Format numbers nicely
        if doc_num >= 1000000:
            result['formatted_doc'] = f"{doc_num:,.0f}" if doc_num == int(doc_num) else f"{doc_num:,.2f}"
        else:
            result['formatted_doc'] = f"{doc_num:.2f}" if doc_num != int(doc_num) else f"{int(doc_num)}"
            
        if excel_num >= 1000000:
            result['formatted_excel'] = f"{excel_num:,.0f}" if excel_num == int(excel_num) else f"{excel_num:,.2f}"
        else:
            result['formatted_excel'] = f"{excel_num:.2f}" if excel_num != int(excel_num) else f"{int(excel_num)}"
        
        # Calculate difference
        if excel_num != 0:
            relative_diff = abs(doc_num - excel_num) / abs(excel_num)
            result['difference'] = f"{relative_diff:.1%}"
            result['match'] = relative_diff <= tolerance
        else:
            result['match'] = doc_num == excel_num
            
        if not result['match']:
            result['mismatch_type'] = 'numeric_mismatch'
            
    except (ValueError, TypeError):
        # Fallback to string comparison
        result['match'] = str(doc_value) == str(excel_value)
        if not result['match']:
            result['mismatch_type'] = 'type_mismatch'
    
    return result

def compare_general_info(doc_info: Dict[str, Any], excel_info: Dict[str, Any], excel_energy_data: Dict[str, Any] = None) -> Dict[str, Any]:
    """
    Compare general information from document against Excel validation data.
    Only compares fields that are present in the document extraction.
    Handles special calculated fields like value_per_finished_product and total_utility_cost.
    """
    comparison = {
        'fields': {},
        'summary': {
            'total_fields': 0,
            'matched_fields': 0,
            'mismatched_fields': 0,
            'missing_in_excel': 0,
            'validated_fields': 0,
            'skipped_fields': 0
        }
    }
    
    # Fields to skip (not available in Excel for comparison)
    skip_fields = {'total_energy_usage', 'total_utility_cost', 'no_of_assessment_recommendations'}
    
    # Only process fields that exist in the document extraction
    for field, doc_val in doc_info.items():
        # Skip fields that shouldn't be compared
        if field in skip_fields:
            comparison['summary']['skipped_fields'] += 1
            continue
            
        comparison['summary']['total_fields'] += 1
        excel_val = None
        
        # Handle special calculated fields
        if field == 'value_per_finished_product':
            # Calculate: annual_sales / annual_production
            annual_sales = excel_info.get('annual_sales', 0)
            annual_production = excel_info.get('annual_production', 0)
            if annual_sales and annual_production and annual_production != 0:
                excel_val = annual_sales / annual_production
                comparison['fields'][field] = compare_values(doc_val, excel_val)
                comparison['fields'][field]['calculation_note'] = f'Calculated: {annual_sales:,.2f} / {annual_production:,.2f} = {excel_val:,.2f}'
            else:
                comparison['fields'][field] = compare_values(doc_val, None)
                comparison['fields'][field]['validation_status'] = 'cannot_calculate'
                comparison['fields'][field]['calculation_note'] = 'Cannot calculate: missing annual_sales or annual_production in Excel'
                
        else:
            # Regular field comparison
            excel_val = excel_info.get(field)
            comparison['fields'][field] = compare_values(doc_val, excel_val)
        
        # Update summary counts
        if excel_val is not None or comparison['fields'][field].get('calculation_note'):
            comparison['summary']['validated_fields'] += 1
            if comparison['fields'][field]['match']:
                comparison['summary']['matched_fields'] += 1
            else:
                comparison['summary']['mismatched_fields'] += 1
        else:
            comparison['summary']['missing_in_excel'] += 1
            # Mark as validation issue when Excel doesn't have the field
            if 'validation_status' not in comparison['fields'][field]:
                comparison['fields'][field]['validation_status'] = 'not_in_excel'
    
    return comparison

def _has_nonzero_data(item: Dict[str, Any]) -> bool:
    """Check if an energy item has any non-zero cost or usage values."""
    cost = item.get('cost', 0)
    if cost and isinstance(cost, (int, float)) and cost > 0:
        return True
    for v in item.get('usage', {}).values():
        if v and isinstance(v, (int, float)) and v > 0:
            return True
    return False

def _build_type_comparison(doc_item: Dict[str, Any], excel_item: Dict[str, Any], validation_status: str) -> Dict[str, Any]:
    """Build a comparison dict for a single energy type pair."""
    type_comparison = {
        'doc_data': doc_item,
        'excel_data': excel_item,
        'cost_comparison': compare_values(
            doc_item.get('cost'),
            excel_item.get('cost')
        ),
        'usage_comparison': {},
        'validation_status': validation_status
    }

    doc_usage = doc_item.get('usage', {})
    excel_usage = excel_item.get('usage', {})

    if doc_usage and excel_usage:
        excel_usage_lower = {k.lower(): (k, v) for k, v in excel_usage.items()}

        for unit, value in doc_usage.items():
            if unit in excel_usage:
                type_comparison['usage_comparison'][unit] = compare_values(value, excel_usage[unit])
            elif unit.lower() in excel_usage_lower:
                _, excel_val = excel_usage_lower[unit.lower()]
                type_comparison['usage_comparison'][unit] = compare_values(value, excel_val)
            elif 'value' in excel_usage and len(doc_usage) == 1:
                type_comparison['usage_comparison'][unit] = compare_values(value, excel_usage['value'])

        if not type_comparison['usage_comparison'] and len(doc_usage) == 1 and len(excel_usage) == 1:
            doc_unit, doc_val = next(iter(doc_usage.items()))
            excel_val = next(iter(excel_usage.values()))
            type_comparison['usage_comparison'][doc_unit] = compare_values(doc_val, excel_val)

    return type_comparison

def compare_energy_data(doc_energy: Dict[str, Any], excel_energy: Dict[str, Any]) -> Dict[str, Any]:
    """
    Compare energy usage data from document against Excel validation data.

    Uses a two-phase matching strategy:
      Phase 1 - Exact standardized-name match
      Phase 2 - Value-based match (cost within tolerance) for remaining types,
                with a name-mismatch warning when labels differ
    Remaining unmatched types from either side are flagged accordingly.
    """
    comparison: Dict[str, Any] = {
        'energy_types': {},
        'summary': {
            'total_types': 0,
            'matched_types': 0,
            'mismatched_types': 0,
            'missing_in_excel': 0,
            'missing_in_doc': 0,
            'name_mismatch_types': 0,
            'validated_types': 0,
            'total_cost_match': False,
            'doc_total_cost': 0,
            'excel_total_cost': 0
        }
    }

    skip_energy_types = {'total_utility'}

    # Build type dicts, preferring entries with the highest cost when
    # multiple rows share the same standardized type (e.g. several Fuel Oil grades).
    doc_types: Dict[str, Dict[str, Any]] = {}
    for item in doc_energy.get('data', []):
        dtype = item['type']
        if dtype in skip_energy_types:
            continue
        if dtype not in doc_types or (item.get('cost', 0) or 0) > (doc_types[dtype].get('cost', 0) or 0):
            doc_types[dtype] = item

    excel_types: Dict[str, Dict[str, Any]] = {}
    for item in excel_energy.get('data', []):
        etype = item['type']
        if etype in skip_energy_types:
            continue
        if etype not in excel_types or (item.get('cost', 0) or 0) > (excel_types[etype].get('cost', 0) or 0):
            excel_types[etype] = item

    matched_doc: set = set()
    matched_excel: set = set()

    # ── Phase 1: Exact name match ──────────────────────────────────────
    for energy_type, doc_item in doc_types.items():
        if energy_type in excel_types:
            excel_item = excel_types[energy_type]
            tc = _build_type_comparison(doc_item, excel_item, 'validated')
            comparison['energy_types'][energy_type] = tc
            matched_doc.add(energy_type)
            matched_excel.add(energy_type)

            comparison['summary']['total_types'] += 1
            comparison['summary']['validated_types'] += 1
            if tc['cost_comparison']['match']:
                comparison['summary']['matched_types'] += 1
            else:
                comparison['summary']['mismatched_types'] += 1

    # ── Phase 2: Value-based match for remaining types ─────────────────
    unmatched_doc = {k: v for k, v in doc_types.items() if k not in matched_doc}
    unmatched_excel = {k: v for k, v in excel_types.items()
                       if k not in matched_excel and _has_nonzero_data(v)}
    value_matched_excel: set = set()

    for doc_type, doc_item in list(unmatched_doc.items()):
        doc_cost = doc_item.get('cost', 0)
        if not doc_cost:
            continue

        candidates = [
            (etype, eitem) for etype, eitem in unmatched_excel.items()
            if etype not in value_matched_excel
            and compare_values(doc_cost, eitem.get('cost', 0))['match']
        ]

        paired = None
        if len(candidates) == 1:
            paired = candidates[0]
        elif len(candidates) > 1:
            doc_usage_vals = list(doc_item.get('usage', {}).values())
            doc_uv = doc_usage_vals[0] if doc_usage_vals else None
            if doc_uv:
                for etype, eitem in candidates:
                    excel_uv_list = list(eitem.get('usage', {}).values())
                    excel_uv = excel_uv_list[0] if excel_uv_list else None
                    if excel_uv and compare_values(doc_uv, excel_uv)['match']:
                        paired = (etype, eitem)
                        break

        if paired:
            excel_type, excel_item = paired
            doc_label = doc_type.replace('_', ' ').title()
            excel_label = excel_item.get('original_name', excel_type.replace('_', ' ').title())

            tc = _build_type_comparison(doc_item, excel_item, 'name_mismatch')
            tc['name_warning'] = f"Document: '{doc_label}' / Excel: '{excel_label}'"
            tc['doc_type_name'] = doc_type
            tc['excel_type_name'] = excel_type

            comparison['energy_types'][doc_type] = tc
            matched_doc.add(doc_type)
            value_matched_excel.add(excel_type)

            comparison['summary']['total_types'] += 1
            comparison['summary']['validated_types'] += 1
            comparison['summary']['name_mismatch_types'] += 1
            if tc['cost_comparison']['match']:
                comparison['summary']['matched_types'] += 1
            else:
                comparison['summary']['mismatched_types'] += 1

    # ── Remaining unmatched document types ──────────────────────────────
    for doc_type, doc_item in doc_types.items():
        if doc_type in matched_doc:
            continue
        tc = _build_type_comparison(doc_item, {}, 'not_in_excel')
        comparison['energy_types'][doc_type] = tc
        comparison['summary']['total_types'] += 1
        comparison['summary']['missing_in_excel'] += 1

    # ── Remaining unmatched Excel types (non-zero data only) ───────────
    for excel_type, excel_item in excel_types.items():
        if excel_type in matched_excel or excel_type in value_matched_excel:
            continue
        if not _has_nonzero_data(excel_item):
            continue
        tc = {
            'doc_data': {},
            'excel_data': excel_item,
            'cost_comparison': compare_values(None, excel_item.get('cost')),
            'usage_comparison': {},
            'validation_status': 'not_in_document'
        }
        comparison['energy_types'][excel_type] = tc
        comparison['summary']['total_types'] += 1
        comparison['summary']['missing_in_doc'] += 1

    # ── Total cost comparison ──────────────────────────────────────────
    doc_total = sum(
        item.get('cost', 0) for item in doc_energy.get('data', [])
        if item.get('cost') and item.get('type') not in skip_energy_types
    )
    excel_total = sum(
        item.get('cost', 0) for item in excel_energy.get('data', [])
        if item.get('cost') and isinstance(item.get('cost'), (int, float)) and item.get('cost') > 0
        and item.get('type') not in skip_energy_types
    )

    comparison['summary']['doc_total_cost'] = doc_total
    comparison['summary']['excel_total_cost'] = excel_total
    total_comparison = compare_values(doc_total, excel_total)
    comparison['summary']['total_cost_match'] = total_comparison['match']
    comparison['summary']['total_cost_comparison'] = total_comparison
    non_skip_excel_count = len([
        item for item in excel_energy.get('data', [])
        if item.get('cost', 0) > 0 and item.get('type') not in skip_energy_types
    ])
    comparison['summary']['excel_total_calculation_note'] = (
        f'Calculated from {non_skip_excel_count} energy cost entries (excluding total utility)'
    )

    return comparison

def compare_ar_sanity_check(assessment_recommendations: List[str], recommendation_summary_table: str) -> Dict[str, Any]:
    """
    Compare individual AR summary tables against the overall recommendation summary table.
    
    Args:
        assessment_recommendations: List of HTML strings, each containing one AR
        recommendation_summary_table: HTML string containing the summary table
        
    Returns:
        Dictionary containing comparison results with summary stats and detailed comparisons
    """
    result = {
        'summary': {
            'total_ars': 0,
            'matched_ars': 0,
            'mismatched_ars': 0,
            'total_field_matches': 0,
            'total_field_differences': 0,
            'has_data': False
        },
        'ar_comparisons': [],
        'error': None
    }
    
    # Handle edge cases
    if not assessment_recommendations:
        result['error'] = 'No assessment recommendations found in document'
        return result
    
    if not recommendation_summary_table or not recommendation_summary_table.strip():
        result['error'] = (
            'No recommendation summary table found in document. '
            'The table detection looks for captions like "Table 1.3 Summary Table" or "Recommendation Summary Table". '
            'If your document uses different wording, the table may not be detected. '
            'Please verify that your document contains a recommendation summary table in Chapter 1.'
        )
        result['debug_info'] = {
            'total_ars_found': len(assessment_recommendations),
            'table_found': False,
            'suggestion': 'Check if the table caption matches patterns like "Summary Table" or is located in Chapter 1'
        }
        return result
    
    try:
        # Extract the recommendation summary table
        rec_summary = get_recommended_summary_table_json(recommendation_summary_table)
        
        if not rec_summary.get('recommendations'):
            result['error'] = 'Could not extract recommendations from summary table'
            return result
        
        result['summary']['has_data'] = True
        result['summary']['total_ars'] = len(assessment_recommendations)
        
        # Process each AR
        for i, ar_html in enumerate(assessment_recommendations):
            try:
                # Extract AR data
                ar_data = get_single_ar_summary_table(ar_html)
                
                if not ar_data.get('ar_number'):
                    # Skip ARs without numbers
                    continue
                
                # Compare AR with summary
                comparison = compare_ar_with_summary(ar_data, rec_summary['recommendations'])
                
                # Track summary statistics
                if comparison.get('error'):
                    # AR not found in summary or other error
                    result['ar_comparisons'].append({
                        'ar_number': ar_data.get('ar_number'),
                        'status': 'error',
                        'error': comparison['error'],
                        'matches': [],
                        'differences': [],
                        'total_matches': 0,
                        'total_differences': 0
                    })
                else:
                    # Successfully compared
                    result['summary']['total_field_matches'] += comparison['total_matches']
                    result['summary']['total_field_differences'] += comparison['total_differences']
                    
                    if comparison['total_differences'] == 0:
                        result['summary']['matched_ars'] += 1
                        status = 'match'
                    else:
                        result['summary']['mismatched_ars'] += 1
                        status = 'mismatch'
                    
                    result['ar_comparisons'].append({
                        'ar_number': comparison['ar_number'],
                        'status': status,
                        'matches': comparison['matches'],
                        'differences': comparison['differences'],
                        'total_matches': comparison['total_matches'],
                        'total_differences': comparison['total_differences']
                    })
                    
            except Exception as e:
                logging.error(f"Error processing AR {i+1}: {e}")
                continue
        
    except Exception as e:
        logging.error(f"Error in AR sanity check: {e}")
        result['error'] = f'Error performing AR sanity check: {str(e)}'
    
    return result


RESOURCE_STREAM_MAPPING = {
    "electricity usage": {
        "unit_savings_field": "electricity_savings_kwh_per_year",
        "dollar_savings_field": "energy_cost_savings_per_year"
    },
    "electricity demand": {
        "unit_savings_field": "demand_savings_kw_per_year",
        "dollar_savings_field": "demand_cost_savings_per_year"
    },
    "natural gas": {
        "unit_savings_field": "propane_savings_mmbtu_per_year",
        "dollar_savings_field": "propane_cost_savings_per_year"
    },
    "lpg": {
        "unit_savings_field": "propane_savings_mmbtu_per_year",
        "dollar_savings_field": "propane_cost_savings_per_year"
    },
    "propane": {
        "unit_savings_field": "propane_savings_mmbtu_per_year",
        "dollar_savings_field": "propane_cost_savings_per_year"
    },
    "administrative changes": {
        "unit_savings_field": None,
        "dollar_savings_field": "admin_cost_savings_per_year"
    },
    "personnel changes": {
        "unit_savings_field": None,
        "dollar_savings_field": "admin_cost_savings_per_year"
    }
}


def compare_recommendation_info_check(
    excel_recommendation_info: Dict[str, Any],
    recommendation_summary_table: str
) -> Dict[str, Any]:
    """
    Compare Excel Recommendation Info sheet against the DOCX recommendation summary table.

    Args:
        excel_recommendation_info: Output of extract_recommendation_info_dict
        recommendation_summary_table: HTML string containing the summary table

    Returns:
        Dictionary containing comparison results with summary stats and detailed comparisons
    """
    result = {
        'summary': {
            'total_ars': 0,
            'matched_ars': 0,
            'mismatched_ars': 0,
            'total_field_matches': 0,
            'total_field_differences': 0,
            'has_data': False
        },
        'ar_comparisons': [],
        'error': None
    }

    if not excel_recommendation_info or not excel_recommendation_info.get('recommendations'):
        result['error'] = 'No recommendation info data found in Excel file'
        return result

    if not recommendation_summary_table or not recommendation_summary_table.strip():
        result['error'] = (
            'No recommendation summary table found in document. '
            'The table detection looks for captions like "Table 1.3 Summary Table" or "Recommendation Summary Table". '
            'If your document uses different wording, the table may not be detected.'
        )
        return result

    try:
        rec_summary = get_recommended_summary_table_json(recommendation_summary_table)
        if not rec_summary.get('recommendations'):
            result['error'] = 'Could not extract recommendations from summary table'
            return result

        result['summary']['has_data'] = True

        for rec in excel_recommendation_info.get('recommendations', []):
            ar_number = rec.get('ar_number')
            if not ar_number:
                continue

            result['summary']['total_ars'] += 1

            summary_rec = next(
                (item for item in rec_summary['recommendations'] if item.get('ar_number') == ar_number),
                None
            )

            if summary_rec is None:
                result['ar_comparisons'].append({
                    'ar_number': ar_number,
                    'status': 'error',
                    'error': f'No matching AR number {ar_number} found in summary table',
                    'matches': [],
                    'differences': [],
                    'total_matches': 0,
                    'total_differences': 0
                })
                continue

            excel_fields: Dict[str, float] = {}

            for stream in rec.get('resource_streams', []):
                stream_type = stream.get('type', '')
                normalized_type = " ".join(str(stream_type).strip().lower().split())
                mapping = RESOURCE_STREAM_MAPPING.get(normalized_type)
                if not mapping:
                    continue

                unit_field = mapping.get('unit_savings_field')
                dollar_field = mapping.get('dollar_savings_field')

                unit_savings = stream.get('unit_savings')
                dollar_savings = stream.get('dollar_savings')

                if unit_field and isinstance(unit_savings, (int, float)):
                    excel_fields[unit_field] = excel_fields.get(unit_field, 0) + unit_savings

                if dollar_field and isinstance(dollar_savings, (int, float)):
                    excel_fields[dollar_field] = excel_fields.get(dollar_field, 0) + dollar_savings

            if isinstance(rec.get('total_dollar_savings'), (int, float)):
                excel_fields['total_cost_savings_per_year'] = rec['total_dollar_savings']

            if isinstance(rec.get('total_implementation_cost'), (int, float)):
                excel_fields['implementation_cost'] = rec['total_implementation_cost']

            matches = []
            differences = []

            for field, excel_value in excel_fields.items():
                doc_value = summary_rec.get(field)

                if doc_value is None and excel_value is None:
                    continue

                if isinstance(excel_value, (int, float)) and isinstance(doc_value, (int, float)):
                    diff = abs(excel_value - doc_value)
                    if diff < 0.01:
                        matches.append({
                            'field': field,
                            'excel_value': excel_value,
                            'docx_value': doc_value,
                            'match': True
                        })
                    else:
                        differences.append({
                            'field': field,
                            'excel_value': excel_value,
                            'docx_value': doc_value,
                            'difference': diff
                        })
                elif excel_value == doc_value:
                    matches.append({
                        'field': field,
                        'excel_value': excel_value,
                        'docx_value': doc_value,
                        'match': True
                    })
                else:
                    differences.append({
                        'field': field,
                        'excel_value': excel_value,
                        'docx_value': doc_value,
                        'difference': 'type mismatch or different values'
                    })

            if differences:
                result['summary']['mismatched_ars'] += 1
                status = 'mismatch'
            else:
                result['summary']['matched_ars'] += 1
                status = 'match'

            result['summary']['total_field_matches'] += len(matches)
            result['summary']['total_field_differences'] += len(differences)

            result['ar_comparisons'].append({
                'ar_number': ar_number,
                'status': status,
                'matches': matches,
                'differences': differences,
                'total_matches': len(matches),
                'total_differences': len(differences)
            })

    except Exception as e:
        logging.error(f"Error in Recommendation Info check: {e}")
        result['error'] = f'Error performing Recommendation Info check: {str(e)}'

    return result

@app.route('/')
def index():
    """Main upload page."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    """Handle file uploads and process them."""
    try:
        # Check if files were uploaded
        if 'docx_file' not in request.files or 'excel_file' not in request.files:
            flash('Both DOCX and Excel files are required', 'error')
            return redirect(url_for('index'))
        
        docx_file = request.files['docx_file']
        excel_file = request.files['excel_file']
        
        # Check if files are selected
        if docx_file.filename == '' or excel_file.filename == '':
            flash('Please select both files', 'error')
            return redirect(url_for('index'))
        
        # Validate file extensions
        if not (allowed_file(docx_file.filename) and allowed_file(excel_file.filename)):
            flash('Invalid file type. Please upload DOCX and XLSX files only', 'error')
            return redirect(url_for('index'))
        
        # Save uploaded files
        docx_filename = secure_filename(docx_file.filename)
        excel_filename = secure_filename(excel_file.filename)
        
        docx_path = UPLOAD_FOLDER / docx_filename
        excel_path = UPLOAD_FOLDER / excel_filename
        
        docx_file.save(str(docx_path))
        excel_file.save(str(excel_path))
        
        # Extract data from both files
        logging.info(f"Processing DOCX file: {docx_path}")
        doc_data = extract_itac_report(str(docx_path), output="html", save_files=False)
        doc_general_info = extract_general_info_fields(doc_data["general_information"])
        doc_energy_data = extract_energy_usage(doc_data["annual_energy_usages_and_costs"])
        
        
        
        
        logging.info(f"Processing Excel file: {excel_path}")
        excel_data = extract_all_structured_info(str(excel_path))
        excel_general_info = excel_data.get("general_info", {})
        excel_energy_data = excel_data.get("energy_waste_info", {})
        ic("Data comparison:")
        ic(excel_general_info)
        ic(doc_general_info)
        
        ic(excel_energy_data)
        ic(doc_energy_data)
        
        # Perform comparisons
        general_comparison = compare_general_info(doc_general_info, excel_general_info, excel_energy_data)
        energy_comparison = compare_energy_data(doc_energy_data, excel_energy_data)
        
        # Perform AR sanity check
        ar_sanity_check = compare_ar_sanity_check(
            doc_data.get('assessment_recommendations', []),
            doc_data.get('recommendation_summary_table', '')
        )

        # Perform Recommendation Info cross-check (Excel vs DOCX summary table)
        rec_info_check = compare_recommendation_info_check(
            excel_data.get('recommendation_info', {}),
            doc_data.get('recommendation_summary_table', '')
        )
        
        # Perform totals validation
        totals_validation = None
        if doc_data.get('recommendation_summary_table'):
            rec_summary = get_recommended_summary_table_json(doc_data['recommendation_summary_table'])
            if rec_summary.get('recommendations'):
                totals_validation = validate_recommendation_totals(rec_summary)
        
        # Validate web links from ARs
        ar_links = doc_data.get('ar_links', {})
        logging.info(f"AR links extracted: {len(ar_links)} AR(s) with links")
        
        if ar_links:
            logging.info("Validating web links from Assessment Recommendations...")
            try:
                link_validation = validate_all_links(ar_links)
                logging.info(f"Link validation complete: {link_validation['summary']['total_links']} links checked")
            except Exception as e:
                logging.error(f"Error validating links: {e}")
                link_validation = {
                    'results': {},
                    'summary': {
                        'total_links': 0,
                        'unique_urls': 0,
                        'working': 0,
                        'warning': 0,
                        'broken': 0,
                        'has_issues': False,
                        'error': str(e)
                    }
                }
        else:
            logging.info("No links found in Assessment Recommendations")
            link_validation = {
                'results': {},
                'summary': {
                    'total_links': 0,
                    'unique_urls': 0,
                    'working': 0,
                    'warning': 0,
                    'broken': 0,
                    'has_issues': False
                }
            }
        
        # Prepare data for template
        template_data = {
            'docx_filename': docx_filename,
            'excel_filename': excel_filename,
            'doc_data': doc_data,
            'excel_data': excel_data,
            'general_comparison': general_comparison,
            'energy_comparison': energy_comparison,
            'ar_sanity_check': ar_sanity_check,
            'rec_info_check': rec_info_check,
            'totals_validation': totals_validation,
            'link_validation': link_validation,
            'doc_general_info': doc_general_info,
            'excel_general_info': excel_general_info,
            'doc_energy_data': doc_energy_data,
            'excel_energy_data': excel_energy_data
        }
        
        # Clean up uploaded files
        docx_path.unlink()
        excel_path.unlink()
        
        return render_template('comparison.html', **template_data)
        
    except Exception as e:
        logging.error(f"Error processing files: {e}")
        flash(f'Error processing files: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/api/compare', methods=['POST'])
def api_compare():
    """API endpoint for programmatic access."""
    try:
        # Handle file uploads via API
        if 'docx_file' not in request.files or 'excel_file' not in request.files:
            return jsonify({'error': 'Both DOCX and Excel files are required'}), 400
        
        docx_file = request.files['docx_file']
        excel_file = request.files['excel_file']
        
        # Process files similar to upload_files but return JSON
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_docx:
            docx_file.save(temp_docx.name)
            doc_data = extract_itac_report(temp_docx.name, output="json", save_files=False)
            doc_general_info = extract_general_info_fields(doc_data.get("general_information", ""))
            doc_energy_data = extract_energy_usage(doc_data.get("annual_energy_usages_and_costs", ""))
            os.unlink(temp_docx.name)
        
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_excel:
            excel_file.save(temp_excel.name)
            excel_data = extract_all_structured_info(temp_excel.name)
            os.unlink(temp_excel.name)
        
        # Perform comparisons
        general_comparison = compare_general_info(doc_general_info, excel_data.get("general_info", {}), excel_data.get("energy_waste_info", {}))
        energy_comparison = compare_energy_data(doc_energy_data, excel_data.get("energy_waste_info", {}))
        
        # Perform AR sanity check
        ar_sanity_check = compare_ar_sanity_check(
            doc_data.get('assessment_recommendations', []),
            doc_data.get('recommendation_summary_table', '')
        )
        
        # Perform totals validation
        totals_validation = None
        if doc_data.get('recommendation_summary_table'):
            rec_summary = get_recommended_summary_table_json(doc_data['recommendation_summary_table'])
            if rec_summary.get('recommendations'):
                totals_validation = validate_recommendation_totals(rec_summary)
        
        # Validate web links from ARs
        ar_links = doc_data.get('ar_links', {})
        logging.info(f"AR links extracted: {len(ar_links)} AR(s) with links")
        
        if ar_links:
            logging.info("Validating web links from Assessment Recommendations...")
            try:
                link_validation = validate_all_links(ar_links)
                logging.info(f"Link validation complete: {link_validation['summary']['total_links']} links checked")
            except Exception as e:
                logging.error(f"Error validating links: {e}")
                link_validation = {
                    'results': {},
                    'summary': {
                        'total_links': 0,
                        'unique_urls': 0,
                        'working': 0,
                        'warning': 0,
                        'broken': 0,
                        'has_issues': False,
                        'error': str(e)
                    }
                }
        else:
            logging.info("No links found in Assessment Recommendations")
            link_validation = {
                'results': {},
                'summary': {
                    'total_links': 0,
                    'unique_urls': 0,
                    'working': 0,
                    'warning': 0,
                    'broken': 0,
                    'has_issues': False
                }
            }
        
        return jsonify({
            'general_comparison': general_comparison,
            'energy_comparison': energy_comparison,
            'ar_sanity_check': ar_sanity_check,
            'totals_validation': totals_validation,
            'link_validation': link_validation,
            'success': True
        })
        
    except Exception as e:
        logging.error(f"API error: {e}")
        return jsonify({'error': str(e)}), 500

# ============================================================================
# AI Agents API Endpoints
# ============================================================================

def discover_agents() -> List[Dict[str, Any]]:
    """
    Dynamically discover available AI agents in the agents/ folder.
    
    Looks for subdirectories containing both agent.py and config.toml files.
    
    Returns:
        List of agent metadata dictionaries with id, name, description, etc.
    """
    agents_dir = Path(__file__).parent / 'agents'
    discovered_agents = []
    
    if not agents_dir.exists():
        logging.warning(f"Agents directory not found: {agents_dir}")
        return []
    
    for item in agents_dir.iterdir():
        if not item.is_dir():
            continue
        
        # Skip special directories
        if item.name.startswith('_') or item.name.startswith('.'):
            continue
        
        agent_file = item / 'agent.py'
        config_file = item / 'config.toml'
        
        # Check if both required files exist
        if agent_file.exists() and config_file.exists():
            try:
                # Read config to get agent metadata
                with open(config_file, 'rb') as f:
                    config = tomli.load(f)
                
                agent_info = {
                    'id': item.name,
                    'name': config.get('agent', {}).get('name', item.name),
                    'description': config.get('agent', {}).get('description', 'No description available'),
                    'version': config.get('agent', {}).get('version', '1.0.0'),
                    'config_path': str(config_file)
                }
                
                discovered_agents.append(agent_info)
                logging.info(f"Discovered agent: {agent_info['name']} ({agent_info['id']})")
                
            except Exception as e:
                logging.error(f"Error reading agent config from {config_file}: {e}")
                continue
    
    return discovered_agents

@app.route('/api/agents', methods=['GET'])
def list_agents():
    """
    API endpoint to list all available AI agents.
    
    Returns:
        JSON response with list of available agents
    """
    try:
        agents = discover_agents()
        return jsonify({
            'success': True,
            'agents': agents,
            'count': len(agents)
        })
    except Exception as e:
        logging.error(f"Error listing agents: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/agents/<agent_id>/run', methods=['POST'])
def run_agent(agent_id: str):
    """
    API endpoint to run a specific AI agent.
    
    Args:
        agent_id: The ID of the agent to run
        
    Request Body:
        JSON containing document data (doc_data, excel_data, etc.)
        
    Returns:
        JSON response with agent analysis results
    """
    try:
        # Get request data
        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'error': 'No data provided'}), 400
        
        # Discover agents to validate agent_id
        agents = discover_agents()
        agent_info = next((a for a in agents if a['id'] == agent_id), None)
        
        if not agent_info:
            return jsonify({'success': False, 'error': f'Agent not found: {agent_id}'}), 404
        
        logging.info(f"Running agent: {agent_info['name']} ({agent_id})")
        
        # Currently only summary_checker is implemented
        if agent_id == 'summary_checker':
            return run_summary_checker_agent(data)
        else:
            return jsonify({
                'success': False,
                'error': f'Agent {agent_id} execution not yet implemented'
            }), 501
            
    except Exception as e:
        logging.error(f"Error running agent {agent_id}: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

def run_summary_checker_agent(data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Run the Summary Checker agent on the provided document data.
    
    Args:
        data: Dictionary containing doc_data with AR information
        
    Returns:
        JSON response with validation results and analysis
    """
    try:
        from agents.summary_checker import analyze_with_llm, check_all_ar_summaries
        from doc_extractor_utils import (
            get_recommended_summary_table_json,
            get_single_ar_summary_table,
            parse_ar_summaries,
        )
        
        # Extract required data from request
        doc_data = data.get('doc_data', {})
        
        if not doc_data:
            return jsonify({
                'success': False,
                'error': 'Missing doc_data in request'
            }), 400
        
        # Get AR summaries
        ar_summaries_html = doc_data.get('ar_summary', '')
        if not ar_summaries_html:
            return jsonify({
                'success': False,
                'error': 'No AR summaries found in document'
            }), 400
        
        # Parse AR summaries
        ar_summaries = parse_ar_summaries(ar_summaries_html)
        
        if not ar_summaries:
            return jsonify({
                'success': False,
                'error': 'Could not parse AR summaries from document'
            }), 400
        
        # Get recommendation summary table
        rec_summary_html = doc_data.get('recommendation_summary_table', '')
        if not rec_summary_html:
            return jsonify({
                'success': False,
                'error': 'No recommendation summary table found in document'
            }), 400
        
        rec_summary = get_recommended_summary_table_json(rec_summary_html)
        summary_recommendations = rec_summary.get('recommendations', [])
        
        if not summary_recommendations:
            return jsonify({
                'success': False,
                'error': 'Could not extract recommendations from summary table'
            }), 400
        
        # Get individual AR data
        assessment_recommendations = doc_data.get('assessment_recommendations', [])
        ar_data_list = []
        for ar_html in assessment_recommendations:
            ar_data = get_single_ar_summary_table(ar_html)
            if ar_data.get('ar_number'):
                ar_data_list.append(ar_data)
        
        if not ar_data_list:
            return jsonify({
                'success': False,
                'error': 'Could not extract individual AR data from document'
            }), 400
        
        # Run validation
        logging.info(f"Validating {len(ar_summaries)} AR summaries...")
        validation_results = check_all_ar_summaries(
            ar_summaries,
            summary_recommendations,
            ar_data_list
        )
        
        # Run LLM analysis
        logging.info("Running LLM analysis on validation results...")
        analysis_report = analyze_with_llm(
            ar_data_list,
            ar_summaries
        )
        
        return jsonify({
            'success': True,
            'agent_id': 'summary_checker',
            'agent_name': 'AR Summary Checker',
            'validation_results': validation_results,
            'analysis_report': analysis_report,
            'summary': {
                'total_ars': len(ar_summaries),
                'ars_with_issues': sum(1 for r in validation_results 
                                      if r.get('validation', {}).get('has_differences', False)),
                'validation_complete': True
            }
        })
        
    except ImportError as e:
        logging.error(f"Import error in summary checker: {e}")
        return jsonify({
            'success': False,
            'error': f'Failed to import required modules: {str(e)}'
        }), 500
    except Exception as e:
        logging.error(f"Error in summary checker agent: {e}")
        return jsonify({
            'success': False,
            'error': f'Agent execution failed: {str(e)}'
        }), 500

@app.route('/api/agents/run_all', methods=['POST'])
def run_all_agents():
    """
    API endpoint to run all available AI agents.
    
    Request Body:
        JSON containing document data
        
    Returns:
        JSON response with results from all agents
    """
    try:
        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'error': 'No data provided'}), 400
        
        agents = discover_agents()
        results = []
        
        for agent in agents:
            agent_id = agent['id']
            logging.info(f"Running agent {agent_id}...")
            
            try:
                # Run each agent
                if agent_id == 'summary_checker':
                    result = run_summary_checker_agent(data)
                    if isinstance(result, tuple):
                        result_data, status_code = result
                        result_json = result_data.get_json()
                    else:
                        result_json = result.get_json()
                    
                    results.append({
                        'agent_id': agent_id,
                        'agent_name': agent['name'],
                        'result': result_json
                    })
                else:
                    results.append({
                        'agent_id': agent_id,
                        'agent_name': agent['name'],
                        'result': {
                            'success': False,
                            'error': 'Agent execution not yet implemented'
                        }
                    })
            except Exception as e:
                logging.error(f"Error running agent {agent_id}: {e}")
                results.append({
                    'agent_id': agent_id,
                    'agent_name': agent['name'],
                    'result': {
                        'success': False,
                        'error': str(e)
                    }
                })
        
        return jsonify({
            'success': True,
            'results': results,
            'total_agents': len(agents),
            'completed': len(results)
        })
        
    except Exception as e:
        logging.error(f"Error running all agents: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.errorhandler(413)
def too_large(e):
    """Handle file too large error."""
    flash('File is too large. Maximum size is 50MB.', 'error')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8000)
