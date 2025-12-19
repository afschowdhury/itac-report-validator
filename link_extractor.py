#!/usr/bin/env python3
"""
Link Extractor Module

Extracts web links from DOCX Assessment Recommendations (ARs).
Handles both embedded hyperlinks and URL patterns in text.
Includes support for extracting links from footnotes.
"""

import re
from typing import Dict, List, Any, Union, Optional, Tuple
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml import parse_xml
from docx.oxml.ns import qn
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def extract_url_patterns_from_text(text: str) -> List[str]:
    """
    Extract URL patterns from plain text.
    
    Matches:
    - http://example.com
    - https://example.com
    - www.example.com
    - ftp://example.com
    
    Args:
        text: Text content to search
        
    Returns:
        List of unique URLs found
    """
    if not text:
        return []
    
    # Comprehensive URL regex pattern
    # Matches: protocol://domain.tld/path or www.domain.tld/path
    url_pattern = re.compile(
        r'(?:(?:https?|ftp)://)(?:\S+(?::\S*)?@)?(?:(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)(?:\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)*(?:\.(?:[a-z\u00a1-\uffff]{2,}))\.?)(?::\d{2,5})?(?:[/?#]\S*)?|'
        r'www\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+(?:\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)*(?:\.(?:[a-z\u00a1-\uffff]{2,}))(?::\d{2,5})?(?:[/?#]\S*)?',
        re.IGNORECASE
    )
    
    urls = url_pattern.findall(text)
    
    # Clean up URLs (remove trailing punctuation that might be captured)
    cleaned_urls = []
    for url in urls:
        # Remove trailing punctuation
        url = re.sub(r'[.,;:)\]}>]+$', '', url)
        # Add http:// to www. URLs
        if url.startswith('www.'):
            url = 'http://' + url
        cleaned_urls.append(url)
    
    return list(set(cleaned_urls))  # Return unique URLs


def extract_hyperlinks_from_paragraph(paragraph: Paragraph) -> List[Dict[str, str]]:
    """
    Extract embedded hyperlinks from a paragraph.
    
    Args:
        paragraph: DOCX Paragraph object
        
    Returns:
        List of dicts with 'url' and 'text' keys
    """
    hyperlinks = []
    
    try:
        # Get the paragraph's XML element
        for child in paragraph._element.iterchildren():
            # Look for hyperlink elements
            if child.tag == qn('w:hyperlink'):
                # Get the relationship ID
                r_id = child.get(qn('r:id'))
                if r_id:
                    # Get the actual URL from the document relationships
                    try:
                        url = paragraph.part.rels[r_id].target_ref
                        # Get the text content of the hyperlink
                        text_elements = child.findall(f'.//{qn("w:t")}')
                        text = ''.join([elem.text for elem in text_elements if elem.text])
                        
                        hyperlinks.append({
                            'url': url,
                            'text': text or url
                        })
                    except (KeyError, AttributeError) as e:
                        logger.debug(f"Could not resolve hyperlink: {e}")
                        continue
    except Exception as e:
        logger.debug(f"Error extracting hyperlinks from paragraph: {e}")
    
    return hyperlinks


def extract_hyperlinks_from_table(table: Table) -> List[Dict[str, str]]:
    """
    Extract embedded hyperlinks from a table.
    
    Args:
        table: DOCX Table object
        
    Returns:
        List of dicts with 'url' and 'text' keys
    """
    hyperlinks = []
    
    try:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    hyperlinks.extend(extract_hyperlinks_from_paragraph(paragraph))
    except Exception as e:
        logger.debug(f"Error extracting hyperlinks from table: {e}")
    
    return hyperlinks


def get_paragraph_context(paragraph: Paragraph, max_length: int = 100) -> str:
    """
    Get text context around a paragraph for reference.
    
    Args:
        paragraph: DOCX Paragraph object
        max_length: Maximum length of context string
        
    Returns:
        Context string
    """
    text = paragraph.text.strip()
    if len(text) <= max_length:
        return text
    return text[:max_length] + "..."


def extract_links_from_ar_blocks(ar_blocks: List[Union[Paragraph, Table]], ar_number: str) -> List[Dict[str, Any]]:
    """
    Extract all links from a single AR's blocks.
    
    Args:
        ar_blocks: List of Paragraph and Table objects for one AR
        ar_number: AR number (e.g., "01", "02")
        
    Returns:
        List of link dictionaries
    """
    links = []
    
    for block in ar_blocks:
        if isinstance(block, Paragraph):
            # Extract embedded hyperlinks
            hyperlinks = extract_hyperlinks_from_paragraph(block)
            for hl in hyperlinks:
                links.append({
                    'url': hl['url'],
                    'text': hl['text'],
                    'type': 'hyperlink',
                    'location': 'paragraph',
                    'context': get_paragraph_context(block)
                })
            
            # Extract URL patterns from text
            text = block.text
            url_patterns = extract_url_patterns_from_text(text)
            for url in url_patterns:
                # Check if this URL is already captured as a hyperlink
                if not any(link['url'] == url for link in links):
                    links.append({
                        'url': url,
                        'text': url,
                        'type': 'text_url',
                        'location': 'paragraph',
                        'context': get_paragraph_context(block)
                    })
        
        elif isinstance(block, Table):
            # Extract embedded hyperlinks from table
            hyperlinks = extract_hyperlinks_from_table(block)
            for hl in hyperlinks:
                links.append({
                    'url': hl['url'],
                    'text': hl['text'],
                    'type': 'hyperlink',
                    'location': 'table',
                    'context': 'Found in table'
                })
            
            # Extract URL patterns from table text
            for row in block.rows:
                for cell in row.cells:
                    cell_text = cell.text
                    url_patterns = extract_url_patterns_from_text(cell_text)
                    for url in url_patterns:
                        # Check if this URL is already captured
                        if not any(link['url'] == url for link in links):
                            links.append({
                                'url': url,
                                'text': url,
                                'type': 'text_url',
                                'location': 'table',
                                'context': cell_text[:100] + ('...' if len(cell_text) > 100 else '')
                            })
    
    return links


def extract_footnotes_from_document(doc: Document) -> Tuple[Dict[str, Any], Any]:
    """
    Extract all footnotes from a document.
    
    Args:
        doc: Document object
        
    Returns:
        Tuple of (footnotes_dict, footnotes_part) where:
        - footnotes_dict maps footnote IDs to their content, links, and display number
        - footnotes_part is the footnotes part object for resolving hyperlinks
    """
    footnotes_dict = {}
    footnotes_part = None
    
    try:
        # Find the footnotes part via document relationships
        for rel_id, rel in doc.part.rels.items():
            if 'footnote' in rel.reltype.lower() and 'endnote' not in rel.reltype.lower():
                footnotes_part = rel.target_part
                logger.debug(f"Found footnotes part via relationship {rel_id}")
                break
        
        if not footnotes_part:
            logger.debug("No footnotes part found in document")
            return {}, None
        
        # Parse the footnotes XML
        footnotes_xml = footnotes_part.blob
        footnotes_element = parse_xml(footnotes_xml)
        
        # Find all footnote elements
        footnotes = footnotes_element.findall(f'.//{qn("w:footnote")}')
        logger.info(f"Found {len(footnotes)} footnote elements in document")
        
        # Filter out special footnotes (separator, continuationSeparator, etc.)
        # and create a mapping from XML ID to display number
        user_footnote_counter = 0
        
        # Extract content and links from each footnote
        for fn in footnotes:
            fn_id = fn.get(qn('w:id'))
            if not fn_id:
                continue
            
            # Check if this is a special footnote (separator, continuationSeparator, etc.)
            fn_type = fn.get(qn('w:type'))
            if fn_type:
                # Skip special footnotes - they don't have user-visible numbers
                logger.debug(f"Skipping special footnote ID {fn_id} of type {fn_type}")
                continue
            
            # This is a user footnote - assign it a display number
            user_footnote_counter += 1
            display_number = user_footnote_counter
            
            # Get text content
            text_elements = fn.findall(f'.//{qn("w:t")}')
            text = ''.join([t.text for t in text_elements if t.text])
            
            # Extract hyperlinks
            hyperlinks = []
            hyperlink_elements = fn.findall(f'.//{qn("w:hyperlink")}')
            for hl in hyperlink_elements:
                r_id = hl.get(qn('r:id'))
                if r_id and r_id in footnotes_part.rels:
                    try:
                        url = footnotes_part.rels[r_id].target_ref
                        # Get hyperlink text
                        hl_text_elements = hl.findall(f'.//{qn("w:t")}')
                        hl_text = ''.join([t.text for t in hl_text_elements if t.text])
                        hyperlinks.append({'url': url, 'text': hl_text or url})
                    except:
                        pass
            
            # Extract plain text URLs
            plain_urls = extract_url_patterns_from_text(text)
            
            footnotes_dict[fn_id] = {
                'text': text,
                'hyperlinks': hyperlinks,
                'plain_urls': plain_urls,
                'display_number': display_number  # The actual footnote number users see
            }
        
        logger.info(f"Extracted {len(footnotes_dict)} footnotes with content")
        
    except Exception as e:
        logger.error(f"Error extracting footnotes: {e}")
        import traceback
        logger.error(traceback.format_exc())
    
    return footnotes_dict, footnotes_part


def find_footnote_references_in_blocks(ar_blocks: List[Union[Paragraph, Table]]) -> List[str]:
    """
    Find all footnote reference IDs in AR blocks.
    
    Args:
        ar_blocks: List of Paragraph and Table objects
        
    Returns:
        List of footnote IDs referenced in the blocks
    """
    footnote_ids = []
    
    for block in ar_blocks:
        if isinstance(block, Paragraph):
            # Look for footnote references in paragraph
            for child in block._element.iterchildren():
                if child.tag == qn('w:r'):  # Run
                    for subchild in child.iterchildren():
                        if subchild.tag == qn('w:footnoteReference'):
                            fn_id = subchild.get(qn('w:id'))
                            if fn_id:
                                footnote_ids.append(fn_id)
        
        elif isinstance(block, Table):
            # Look for footnote references in table cells
            for row in block.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for child in para._element.iterchildren():
                            if child.tag == qn('w:r'):
                                for subchild in child.iterchildren():
                                    if subchild.tag == qn('w:footnoteReference'):
                                        fn_id = subchild.get(qn('w:id'))
                                        if fn_id:
                                            footnote_ids.append(fn_id)
    
    return footnote_ids


def extract_links_from_all_ars(doc: Document, ar_blocks_list: List[List[Union[Paragraph, Table]]]) -> Dict[str, List[Dict[str, Any]]]:
    """
    Extract links from all ARs in a document, including footnotes.
    
    Args:
        doc: Document object (needed to access footnotes)
        ar_blocks_list: List of AR block lists from document_extractor.extract_ars()
        
    Returns:
        Dictionary mapping AR numbers to lists of links:
        {
            "AR_01": [{url, text, type, location, context}, ...],
            "AR_02": [...],
            ...
        }
    """
    all_links = {}
    
    # Extract all footnotes from the document once
    footnotes_dict, footnotes_part = extract_footnotes_from_document(doc)
    
    for idx, ar_blocks in enumerate(ar_blocks_list):
        ar_number = f"AR_{idx + 1:02d}"  # Format as AR_01, AR_02, etc.
        
        try:
            # Extract links from AR blocks (paragraphs and tables)
            links = extract_links_from_ar_blocks(ar_blocks, ar_number)
            
            # Find footnote references in this AR
            footnote_refs = find_footnote_references_in_blocks(ar_blocks)
            
            if footnote_refs and footnotes_dict:
                logger.debug(f"{ar_number}: Found {len(footnote_refs)} footnote reference(s)")
                
                # Extract links from referenced footnotes
                for fn_id in footnote_refs:
                    if fn_id in footnotes_dict:
                        footnote = footnotes_dict[fn_id]
                        display_num = footnote.get('display_number', fn_id)
                        
                        # Add hyperlinks from footnote
                        for hl in footnote['hyperlinks']:
                            links.append({
                                'url': hl['url'],
                                'text': hl['text'],
                                'type': 'hyperlink',
                                'location': 'footnote',
                                'context': f"Footnote {display_num}: {footnote['text'][:80]}..."
                            })
                        
                        # Add plain text URLs from footnote
                        # Remove duplicates (URLs that are already captured as hyperlinks)
                        existing_urls = {link['url'] for link in links}
                        for url in footnote['plain_urls']:
                            if url not in existing_urls:
                                links.append({
                                    'url': url,
                                    'text': url,
                                    'type': 'text_url',
                                    'location': 'footnote',
                                    'context': f"Footnote {display_num}: {footnote['text'][:80]}..."
                                })
                                existing_urls.add(url)
            
            if links:  # Only add if there are links
                all_links[ar_number] = links
                logger.info(f"Found {len(links)} link(s) in {ar_number} (including footnotes)")
            else:
                logger.debug(f"No links found in {ar_number}")
        
        except Exception as e:
            logger.error(f"Error extracting links from {ar_number}: {e}")
            import traceback
            logger.error(traceback.format_exc())
            continue
    
    return all_links


def get_link_statistics(links_dict: Dict[str, List[Dict[str, Any]]]) -> Dict[str, Any]:
    """
    Calculate statistics about extracted links.
    
    Args:
        links_dict: Dictionary of links by AR
        
    Returns:
        Statistics dictionary
    """
    total_links = sum(len(links) for links in links_dict.values())
    hyperlink_count = sum(
        sum(1 for link in links if link['type'] == 'hyperlink')
        for links in links_dict.values()
    )
    text_url_count = sum(
        sum(1 for link in links if link['type'] == 'text_url')
        for links in links_dict.values()
    )
    
    return {
        'total_links': total_links,
        'total_ars_with_links': len(links_dict),
        'hyperlink_count': hyperlink_count,
        'text_url_count': text_url_count,
        'links_by_ar': {ar: len(links) for ar, links in links_dict.items()}
    }


if __name__ == '__main__':
    # Test the module
    import sys
    from document_extractor import extract_itac_report
    
    if len(sys.argv) > 1:
        docx_path = sys.argv[1]
        print(f"Extracting links from: {docx_path}")
        
        # Extract document data
        doc_data = extract_itac_report(docx_path, output="html", save_files=False)
        
        # This will be updated once integrated with document_extractor
        print("Link extraction standalone test - integration required")
    else:
        print("Usage: python link_extractor.py <path_to_docx>")

