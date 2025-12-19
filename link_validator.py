#!/usr/bin/env python3
"""
Link Validator Module

Validates web links with comprehensive error reporting.
Handles SSL validation, redirects, timeouts, and various error types.
"""

import requests
import time
import logging
from typing import Dict, List, Any, Optional
from urllib.parse import urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed
import ssl
import socket

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configure requests to use a reasonable timeout
DEFAULT_TIMEOUT = 10  # seconds
MAX_REDIRECTS = 10
MAX_CONCURRENT_REQUESTS = 10

# User agent to avoid being blocked by some sites
USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'


def categorize_link_status(validation_result: Dict[str, Any]) -> str:
    """
    Categorize link status based on validation results.
    
    Args:
        validation_result: Validation result dictionary
        
    Returns:
        Status string: "working", "warning", or "broken"
    """
    if validation_result.get('error_type'):
        return 'broken'
    
    http_code = validation_result.get('http_code', 0)
    
    # Working: 2xx status codes
    if 200 <= http_code < 300:
        # Check for warnings (redirects, slow response, SSL issues)
        has_redirects = len(validation_result.get('redirect_chain', [])) > 1
        is_slow = validation_result.get('response_time_ms', 0) > 5000  # > 5 seconds
        ssl_warning = not validation_result.get('ssl_valid', True)
        
        if has_redirects or is_slow or ssl_warning:
            return 'warning'
        return 'working'
    
    # Warning: 3xx redirects without final resolution
    if 300 <= http_code < 400:
        return 'warning'
    
    # Warning: Authentication/permission issues (likely accessible to real users)
    # 401 Unauthorized - requires authentication
    # 403 Forbidden - often bot detection, CAPTCHA, or cookie requirements
    # Other 4xx except 404 - client-side issues but resource may exist
    if 400 <= http_code < 500 and http_code != 404:
        return 'warning'
    
    # Broken: 404 Not Found and 5xx server errors
    if http_code == 404 or http_code >= 500:
        return 'broken'
    
    # Default: treat unknown codes as broken
    return 'broken'


def get_error_suggestion(error_type: str, http_code: Optional[int] = None) -> str:
    """
    Get user-friendly fix suggestion based on error type.
    
    Args:
        error_type: Type of error encountered
        http_code: HTTP status code if applicable
        
    Returns:
        Suggestion string
    """
    suggestions = {
        'timeout': 'The server took too long to respond. Check if the URL is correct or try again later.',
        'dns': 'Domain name could not be resolved. Check if the URL is spelled correctly.',
        'ssl': 'SSL certificate is invalid or expired. The site may be insecure or misconfigured.',
        'connection': 'Could not establish a connection. The server may be down or blocking requests.',
        'too_many_redirects': 'Too many redirects detected. The URL may be misconfigured.',
        'invalid_url': 'The URL format is invalid. Check the URL syntax.',
    }
    
    # HTTP status code specific suggestions
    if http_code:
        if http_code == 404:
            return 'Page not found (404). The URL is broken or the page has been moved/deleted.'
        elif http_code == 403:
            return 'Access forbidden (403). This is often due to bot detection or cookie requirements. The page may be accessible in a regular web browser with popups/CAPTCHA enabled.'
        elif http_code == 401:
            return 'Authentication required (401). This page requires login credentials, but the resource exists.'
        elif http_code == 400:
            return 'Bad request (400). The URL syntax may be malformed or missing required parameters.'
        elif http_code == 405:
            return 'Method not allowed (405). The page exists but doesn\'t accept GET requests.'
        elif http_code == 429:
            return 'Too many requests (429). Rate limiting in effect. Try again later.'
        elif 400 <= http_code < 500:
            return f'Client error ({http_code}). There may be an issue with how the request is formatted, but the resource likely exists.'
        elif 500 <= http_code < 600:
            return f'Server error ({http_code}). The website server is down or experiencing technical issues.'
    
    return suggestions.get(error_type, 'An unknown error occurred. Please verify the URL.')


def validate_url(url: str, timeout: int = DEFAULT_TIMEOUT) -> Dict[str, Any]:
    """
    Comprehensively validate a URL.
    
    Args:
        url: URL to validate
        timeout: Request timeout in seconds
        
    Returns:
        Validation result dictionary with status, errors, redirects, etc.
    """
    result = {
        'url': url,
        'status': 'broken',
        'http_code': None,
        'response_time_ms': None,
        'final_url': url,
        'redirect_chain': [url],
        'ssl_valid': None,
        'error_message': None,
        'error_type': None,
        'suggestion': None
    }
    
    # Basic URL validation
    try:
        parsed = urlparse(url)
        if not parsed.scheme or not parsed.netloc:
            result['error_type'] = 'invalid_url'
            result['error_message'] = 'Invalid URL format'
            result['suggestion'] = get_error_suggestion('invalid_url')
            return result
    except Exception as e:
        result['error_type'] = 'invalid_url'
        result['error_message'] = str(e)
        result['suggestion'] = get_error_suggestion('invalid_url')
        return result
    
    # Perform HTTP request
    start_time = time.time()
    
    try:
        # Create a session to track redirects
        session = requests.Session()
        session.max_redirects = MAX_REDIRECTS
        
        # Configure headers
        headers = {
            'User-Agent': USER_AGENT,
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive',
        }
        
        # Make the request
        response = session.get(
            url,
            headers=headers,
            timeout=timeout,
            allow_redirects=True,
            verify=True  # Verify SSL certificates
        )
        
        # Calculate response time
        response_time_ms = int((time.time() - start_time) * 1000)
        
        # Extract redirect chain
        redirect_chain = [url]
        if response.history:
            for resp in response.history:
                redirect_chain.append(resp.url)
            redirect_chain.append(response.url)
        
        # Update result with success data
        result['http_code'] = response.status_code
        result['response_time_ms'] = response_time_ms
        result['final_url'] = response.url
        result['redirect_chain'] = redirect_chain
        result['ssl_valid'] = parsed.scheme == 'https'
        
        # Categorize status
        result['status'] = categorize_link_status(result)
        
        # Add suggestion if there's a warning or error
        if result['status'] in ['warning', 'broken']:
            result['suggestion'] = get_error_suggestion(None, response.status_code)
        
        logger.debug(f"Validated {url}: {result['status']} ({response.status_code})")
        
    except requests.exceptions.SSLError as e:
        result['error_type'] = 'ssl'
        result['error_message'] = f'SSL certificate error: {str(e)}'
        result['suggestion'] = get_error_suggestion('ssl')
        result['ssl_valid'] = False
        logger.debug(f"SSL error for {url}: {e}")
        
    except requests.exceptions.Timeout:
        result['error_type'] = 'timeout'
        result['error_message'] = f'Request timed out after {timeout} seconds'
        result['suggestion'] = get_error_suggestion('timeout')
        logger.debug(f"Timeout for {url}")
        
    except requests.exceptions.TooManyRedirects:
        result['error_type'] = 'too_many_redirects'
        result['error_message'] = f'Too many redirects (>{MAX_REDIRECTS})'
        result['suggestion'] = get_error_suggestion('too_many_redirects')
        logger.debug(f"Too many redirects for {url}")
        
    except requests.exceptions.ConnectionError as e:
        # Try to determine if it's a DNS error
        error_str = str(e).lower()
        if 'name or service not known' in error_str or 'nodename nor servname provided' in error_str:
            result['error_type'] = 'dns'
            result['error_message'] = 'Domain name could not be resolved'
            result['suggestion'] = get_error_suggestion('dns')
        else:
            result['error_type'] = 'connection'
            result['error_message'] = f'Connection error: {str(e)}'
            result['suggestion'] = get_error_suggestion('connection')
        logger.debug(f"Connection error for {url}: {e}")
        
    except requests.exceptions.RequestException as e:
        result['error_type'] = 'http_error'
        result['error_message'] = f'Request error: {str(e)}'
        result['suggestion'] = get_error_suggestion('http_error')
        logger.debug(f"Request error for {url}: {e}")
        
    except Exception as e:
        result['error_type'] = 'unknown'
        result['error_message'] = f'Unexpected error: {str(e)}'
        result['suggestion'] = 'An unexpected error occurred. Please check the URL and try again.'
        logger.error(f"Unexpected error validating {url}: {e}")
    
    return result


def validate_all_links(links_dict: Dict[str, List[Dict[str, Any]]], max_workers: int = MAX_CONCURRENT_REQUESTS) -> Dict[str, Any]:
    """
    Validate all links from multiple ARs concurrently.
    
    Args:
        links_dict: Dictionary mapping AR numbers to lists of link dictionaries
        max_workers: Maximum number of concurrent requests
        
    Returns:
        Dictionary with validation results and summary statistics
    """
    validation_results = {}
    url_cache = {}  # Cache to avoid validating the same URL multiple times
    
    # Flatten all links with AR context
    links_to_validate = []
    for ar_num, links in links_dict.items():
        for link in links:
            url = link['url']
            links_to_validate.append({
                'ar_num': ar_num,
                'url': url,
                'link_data': link
            })
    
    total_links = len(links_to_validate)
    logger.info(f"Validating {total_links} link(s) from {len(links_dict)} AR(s)...")
    
    # Validate links concurrently
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # Submit validation tasks
        future_to_link = {}
        for link_info in links_to_validate:
            url = link_info['url']
            
            # Check cache first
            if url in url_cache:
                continue
            
            future = executor.submit(validate_url, url)
            future_to_link[future] = link_info
        
        # Collect results
        completed = 0
        for future in as_completed(future_to_link):
            link_info = future_to_link[future]
            ar_num = link_info['ar_num']
            url = link_info['url']
            
            try:
                validation_result = future.result()
                url_cache[url] = validation_result
                completed += 1
                
                if completed % 10 == 0 or completed == len(future_to_link):
                    logger.info(f"Validated {completed}/{len(future_to_link)} unique URLs")
                    
            except Exception as e:
                logger.error(f"Error validating {url}: {e}")
                url_cache[url] = {
                    'url': url,
                    'status': 'broken',
                    'error_type': 'unknown',
                    'error_message': str(e),
                    'suggestion': 'An unexpected error occurred during validation.'
                }
    
    # Build results structure
    for link_info in links_to_validate:
        ar_num = link_info['ar_num']
        url = link_info['url']
        link_data = link_info['link_data']
        
        if ar_num not in validation_results:
            validation_results[ar_num] = []
        
        # Combine link data with validation result
        combined_result = {**link_data, **url_cache.get(url, {})}
        validation_results[ar_num].append(combined_result)
    
    # Calculate summary statistics
    summary = calculate_validation_summary(validation_results)
    
    return {
        'results': validation_results,
        'summary': summary
    }


def calculate_validation_summary(validation_results: Dict[str, List[Dict[str, Any]]]) -> Dict[str, Any]:
    """
    Calculate summary statistics for validation results.
    
    Args:
        validation_results: Dictionary of validation results by AR
        
    Returns:
        Summary statistics dictionary
    """
    total_links = 0
    working_links = 0
    warning_links = 0
    broken_links = 0
    unique_urls = set()
    
    error_types = {}
    ar_stats = {}
    
    for ar_num, links in validation_results.items():
        ar_working = 0
        ar_warning = 0
        ar_broken = 0
        
        for link in links:
            total_links += 1
            unique_urls.add(link.get('url'))
            
            status = link.get('status', 'broken')
            if status == 'working':
                working_links += 1
                ar_working += 1
            elif status == 'warning':
                warning_links += 1
                ar_warning += 1
            else:
                broken_links += 1
                ar_broken += 1
                
                # Track error types
                error_type = link.get('error_type', 'unknown')
                error_types[error_type] = error_types.get(error_type, 0) + 1
        
        ar_stats[ar_num] = {
            'total': len(links),
            'working': ar_working,
            'warning': ar_warning,
            'broken': ar_broken
        }
    
    return {
        'total_links': total_links,
        'unique_urls': len(unique_urls),
        'working': working_links,
        'warning': warning_links,
        'broken': broken_links,
        'error_types': error_types,
        'ar_stats': ar_stats,
        'has_issues': broken_links > 0 or warning_links > 0
    }


if __name__ == '__main__':
    # Test the module
    test_links = {
        'AR_01': [
            {'url': 'https://www.google.com', 'text': 'Google', 'type': 'hyperlink', 'location': 'paragraph', 'context': 'Test link'},
            {'url': 'https://thissitedoesnotexist12345.com', 'text': 'Broken', 'type': 'hyperlink', 'location': 'paragraph', 'context': 'Test broken link'},
        ]
    }
    
    print("Testing link validator...")
    results = validate_all_links(test_links)
    
    print("\nSummary:")
    print(f"  Total links: {results['summary']['total_links']}")
    print(f"  Working: {results['summary']['working']}")
    print(f"  Warnings: {results['summary']['warning']}")
    print(f"  Broken: {results['summary']['broken']}")
    
    print("\nDetailed Results:")
    for ar_num, links in results['results'].items():
        print(f"\n{ar_num}:")
        for link in links:
            print(f"  - {link['url']}")
            print(f"    Status: {link['status']}")
            if link.get('error_message'):
                print(f"    Error: {link['error_message']}")

