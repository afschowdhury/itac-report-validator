/**
 * ITAC Report Validator - Main JavaScript
 */

document.addEventListener('DOMContentLoaded', function() {
    initializeFileUpload();
    initializeTooltips();
    initializeAnimations();
});

/**
 * Initialize file upload functionality
 */
function initializeFileUpload() {
    const dropZones = document.querySelectorAll('.file-drop-zone');
    
    dropZones.forEach(dropZone => {
        const targetInputId = dropZone.getAttribute('data-target');
        const targetInput = document.getElementById(targetInputId);
        const browseButton = dropZone.querySelector('.btn');
        const fileInfo = dropZone.querySelector('.file-info');
        const removeButton = dropZone.querySelector('.remove-file');
        
        if (!targetInput) return;
        
        // Handle browse button click
        browseButton?.addEventListener('click', (e) => {
            e.preventDefault();
            targetInput.click();
        });
        
        // Handle file input change
        targetInput.addEventListener('change', function() {
            handleFileSelect(this.files[0], dropZone, fileInfo);
            checkSubmitButton();
        });
        
        // Handle drag and drop
        dropZone.addEventListener('dragover', handleDragOver);
        dropZone.addEventListener('dragleave', handleDragLeave);
        dropZone.addEventListener('drop', (e) => handleDrop(e, targetInput, dropZone, fileInfo));
        
        // Handle remove file
        removeButton?.addEventListener('click', () => {
            removeFile(targetInput, dropZone, fileInfo);
            checkSubmitButton();
        });
    });
}

/**
 * Handle drag over event
 */
function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    this.classList.add('dragover');
}

/**
 * Handle drag leave event
 */
function handleDragLeave(e) {
    e.preventDefault();
    e.stopPropagation();
    this.classList.remove('dragover');
}

/**
 * Handle drop event
 */
function handleDrop(e, targetInput, dropZone, fileInfo) {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        const file = files[0];
        
        // Validate file type
        const allowedTypes = targetInput.getAttribute('accept').split(',');
        const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
        
        if (allowedTypes.includes(fileExtension)) {
            targetInput.files = files;
            handleFileSelect(file, dropZone, fileInfo);
            checkSubmitButton();
        } else {
            showAlert('Invalid file type. Please upload ' + allowedTypes.join(' or ') + ' files only.', 'error');
        }
    }
}

/**
 * Handle file selection
 */
function handleFileSelect(file, dropZone, fileInfo) {
    if (!file) return;
    
    const fileName = fileInfo.querySelector('.file-name');
    const uploadContent = dropZone.querySelector('.text-center');
    
    if (fileName) {
        fileName.textContent = file.name;
        fileInfo.classList.remove('d-none');
        uploadContent.classList.add('d-none');
        dropZone.classList.add('has-file');
    }
    
    // Validate file size (16MB limit)
    const maxSize = 16 * 1024 * 1024; // 16MB
    if (file.size > maxSize) {
        showAlert('File is too large. Maximum size is 16MB.', 'error');
        removeFile(dropZone.querySelector('input'), dropZone, fileInfo);
        return;
    }
    
    showAlert(`Selected: ${file.name} (${formatFileSize(file.size)})`, 'success');
}

/**
 * Remove selected file
 */
function removeFile(input, dropZone, fileInfo) {
    input.value = '';
    fileInfo.classList.add('d-none');
    dropZone.querySelector('.text-center').classList.remove('d-none');
    dropZone.classList.remove('has-file');
}

/**
 * Check if submit button should be enabled
 */
function checkSubmitButton() {
    const submitBtn = document.getElementById('submitBtn');
    if (!submitBtn) return;
    
    const docxFile = document.getElementById('docx_file');
    const excelFile = document.getElementById('excel_file');
    
    const bothFilesSelected = docxFile?.files.length > 0 && excelFile?.files.length > 0;
    submitBtn.disabled = !bothFilesSelected;
    
    if (bothFilesSelected) {
        submitBtn.innerHTML = '<i class="fas fa-sync-alt me-2"></i>Compare Reports';
    } else {
        submitBtn.innerHTML = '<i class="fas fa-upload me-2"></i>Select Files First';
    }
}

/**
 * Initialize tooltips
 */
function initializeTooltips() {
    const tooltipTriggerList = [].slice.call(document.querySelectorAll('[title]'));
    tooltipTriggerList.forEach(function (tooltipTriggerEl) {
        if (typeof bootstrap !== 'undefined' && bootstrap.Tooltip) {
            new bootstrap.Tooltip(tooltipTriggerEl);
        }
    });
}

/**
 * Initialize animations and interactions
 */
function initializeAnimations() {
    // Smooth scrolling for anchor links
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            e.preventDefault();
            const target = document.querySelector(this.getAttribute('href'));
            if (target) {
                target.scrollIntoView({
                    behavior: 'smooth',
                    block: 'start'
                });
            }
        });
    });
    
    // Card hover effects
    const cards = document.querySelectorAll('.card');
    cards.forEach(card => {
        card.addEventListener('mouseenter', function() {
            this.style.transform = 'translateY(-2px)';
        });
        
        card.addEventListener('mouseleave', function() {
            this.style.transform = 'translateY(0)';
        });
    });
    
    // Section collapse icons
    const sectionToggles = document.querySelectorAll('.section-toggle');
    sectionToggles.forEach(toggle => {
        const target = toggle.getAttribute('data-bs-target');
        const collapseElement = document.querySelector(target);
        
        if (collapseElement) {
            collapseElement.addEventListener('shown.bs.collapse', function() {
                const chevron = toggle.querySelector('.fa-chevron-down, .fa-chevron-up');
                if (chevron) {
                    chevron.classList.remove('fa-chevron-down');
                    chevron.classList.add('fa-chevron-up');
                }
            });
            
            collapseElement.addEventListener('hidden.bs.collapse', function() {
                const chevron = toggle.querySelector('.fa-chevron-down, .fa-chevron-up');
                if (chevron) {
                    chevron.classList.remove('fa-chevron-up');
                    chevron.classList.add('fa-chevron-down');
                }
            });
        }
    });
}

/**
 * Show alert message
 */
function showAlert(message, type = 'info') {
    // Create alert element
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type === 'error' ? 'danger' : type} alert-dismissible fade show`;
    alertDiv.style.position = 'fixed';
    alertDiv.style.top = '20px';
    alertDiv.style.right = '20px';
    alertDiv.style.zIndex = '9999';
    alertDiv.style.maxWidth = '400px';
    
    const icon = type === 'error' ? 'exclamation-triangle' : 'check-circle';
    alertDiv.innerHTML = `
        <i class="fas fa-${icon} me-2"></i>
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    `;
    
    document.body.appendChild(alertDiv);
    
    // Auto remove after 5 seconds
    setTimeout(() => {
        if (alertDiv.parentNode) {
            alertDiv.remove();
        }
    }, 5000);
}

/**
 * Format file size for display
 */
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * Copy text to clipboard
 */
function copyToClipboard(text) {
    if (navigator.clipboard) {
        navigator.clipboard.writeText(text).then(() => {
            showAlert('Copied to clipboard!', 'success');
        }).catch(() => {
            fallbackCopyToClipboard(text);
        });
    } else {
        fallbackCopyToClipboard(text);
    }
}

/**
 * Fallback copy method for older browsers
 */
function fallbackCopyToClipboard(text) {
    const textArea = document.createElement('textarea');
    textArea.value = text;
    textArea.style.position = 'fixed';
    textArea.style.left = '-999999px';
    textArea.style.top = '-999999px';
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    
    try {
        document.execCommand('copy');
        showAlert('Copied to clipboard!', 'success');
    } catch (err) {
        showAlert('Failed to copy to clipboard', 'error');
    }
    
    document.body.removeChild(textArea);
}

/**
 * Export comparison results as JSON
 */
function exportResults() {
    const data = {
        timestamp: new Date().toISOString(),
        general_comparison: window.generalComparison || {},
        energy_comparison: window.energyComparison || {}
    };
    
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `itac-comparison-${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    showAlert('Results exported successfully!', 'success');
}

/**
 * Print comparison results
 */
function printResults() {
    window.print();
}

// Global functions for template access
window.copyToClipboard = copyToClipboard;
window.exportResults = exportResults;
window.printResults = printResults;
