#!/usr/bin/env python3
"""
ITAC Report Validator Web Application

A Flask web application for uploading and comparing DOCX and Excel ITAC reports.
Extracts data using document_extractor.py and excel_keyinfo_extractor.py and 
highlights mismatches between the two sources.
"""

import os
import json
import tempfile
from pathlib import Path
from typing import Dict, Any, Tuple, List
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from werkzeug.utils import secure_filename
import logging

# Import our existing extractors
from document_extractor import extract_itac_report, extract_general_info_fields, extract_energy_usage
from excel_keyinfo_extractor import extract_all_structured_info

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

app = Flask(__name__)
app.secret_key = 'itac-validator-secret-key-2024'  # Change this in production
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

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

def compare_general_info(doc_info: Dict[str, Any], excel_info: Dict[str, Any]) -> Dict[str, Any]:
    """
    Compare general information from document against Excel validation data.
    Only compares fields that are present in the document extraction.
    """
    comparison = {
        'fields': {},
        'summary': {
            'total_fields': 0,
            'matched_fields': 0,
            'mismatched_fields': 0,
            'missing_in_excel': 0,
            'validated_fields': 0
        }
    }
    
    # Only process fields that exist in the document extraction
    for field, doc_val in doc_info.items():
        excel_val = excel_info.get(field)
        
        comparison['fields'][field] = compare_values(doc_val, excel_val)
        comparison['summary']['total_fields'] += 1
        
        if excel_val is not None:
            comparison['summary']['validated_fields'] += 1
            if comparison['fields'][field]['match']:
                comparison['summary']['matched_fields'] += 1
            else:
                comparison['summary']['mismatched_fields'] += 1
        else:
            comparison['summary']['missing_in_excel'] += 1
            # Mark as validation issue when Excel doesn't have the field
            comparison['fields'][field]['validation_status'] = 'not_in_excel'
    
    return comparison

def compare_energy_data(doc_energy: Dict[str, Any], excel_energy: Dict[str, Any]) -> Dict[str, Any]:
    """
    Compare energy usage data from document against Excel validation data.
    Only compares energy types that are present in the document extraction.
    """
    comparison = {
        'energy_types': {},
        'summary': {
            'total_types': 0,
            'matched_types': 0,
            'mismatched_types': 0,
            'missing_in_excel': 0,
            'validated_types': 0,
            'total_cost_match': False,
            'doc_total_cost': 0,
            'excel_total_cost': 0
        }
    }
    
    # Create mappings by energy type
    doc_types = {item['type']: item for item in doc_energy.get('data', [])}
    excel_types = {item['type']: item for item in excel_energy.get('data', [])}
    
    # Only process energy types that exist in the document extraction
    for energy_type, doc_item in doc_types.items():
        excel_item = excel_types.get(energy_type, {})
        
        type_comparison = {
            'doc_data': doc_item,
            'excel_data': excel_item,
            'cost_comparison': compare_values(
                doc_item.get('cost'), 
                excel_item.get('cost')
            ),
            'usage_comparison': {},
            'validation_status': 'validated' if excel_item else 'not_in_excel'
        }
        
        # Compare usage values if both sources have data
        doc_usage = doc_item.get('usage', {})
        excel_usage = excel_item.get('usage', {})
        
        if doc_usage and excel_usage:
            # Look for matching units or values
            for unit, value in doc_usage.items():
                if unit in excel_usage:
                    type_comparison['usage_comparison'][unit] = compare_values(value, excel_usage[unit])
                elif 'value' in excel_usage and len(doc_usage) == 1:
                    type_comparison['usage_comparison'][unit] = compare_values(value, excel_usage['value'])
        
        comparison['energy_types'][energy_type] = type_comparison
        comparison['summary']['total_types'] += 1
        
        if excel_item:  # Only count as validated if Excel has this energy type
            comparison['summary']['validated_types'] += 1
            if type_comparison['cost_comparison']['match']:
                comparison['summary']['matched_types'] += 1
            else:
                comparison['summary']['mismatched_types'] += 1
        else:
            comparison['summary']['missing_in_excel'] += 1
    
    # Compare total costs (only from document perspective)
    doc_total = sum(item.get('cost', 0) for item in doc_energy.get('data', []) if item.get('cost'))
    excel_total = excel_energy.get('summary', {}).get('total_utility_cost', 0) or \
                 excel_energy.get('summary', {}).get('total_energy_cost', 0)
    
    comparison['summary']['doc_total_cost'] = doc_total
    comparison['summary']['excel_total_cost'] = excel_total
    total_comparison = compare_values(doc_total, excel_total)
    comparison['summary']['total_cost_match'] = total_comparison['match']
    comparison['summary']['total_cost_comparison'] = total_comparison
    
    return comparison

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
        
        # Perform comparisons
        general_comparison = compare_general_info(doc_general_info, excel_general_info)
        energy_comparison = compare_energy_data(doc_energy_data, excel_energy_data)
        
        # Prepare data for template
        template_data = {
            'docx_filename': docx_filename,
            'excel_filename': excel_filename,
            'doc_data': doc_data,
            'excel_data': excel_data,
            'general_comparison': general_comparison,
            'energy_comparison': energy_comparison,
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
        general_comparison = compare_general_info(doc_general_info, excel_data.get("general_info", {}))
        energy_comparison = compare_energy_data(doc_energy_data, excel_data.get("energy_waste_info", {}))
        
        return jsonify({
            'general_comparison': general_comparison,
            'energy_comparison': energy_comparison,
            'success': True
        })
        
    except Exception as e:
        logging.error(f"API error: {e}")
        return jsonify({'error': str(e)}), 500

@app.errorhandler(413)
def too_large(e):
    """Handle file too large error."""
    flash('File is too large. Maximum size is 16MB.', 'error')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8000)
