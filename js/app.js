// Global variables
let excelData = null;
let excelHeaders = [];
let wordTemplateContent = null;
let wordPlaceholders = [];
let templateName = ''; // Store template name for filename prefix

// Document ready function
document.addEventListener('DOMContentLoaded', function() {
    // Initialize event listeners
    document.getElementById('excel-file').addEventListener('change', handleExcelUpload);
    document.getElementById('word-template').addEventListener('change', handleWordTemplateUpload);
    document.getElementById('generate-btn').addEventListener('click', generateDocument);
    document.getElementById('debug-btn').addEventListener('click', showDebugInfo);
    
    // Improved file upload/remove logic
    setFileInfo('excel-file', 'excel-file-info', 'remove-excel-file');
    setFileInfo('word-template', 'word-template-info', 'remove-word-file');
});

/**
 * Handle Excel file upload
 */
function handleExcelUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    // Update file info display
    document.getElementById('excel-file-info').textContent = `Đã chọn: ${file.name}`;
    
    // Read the Excel file
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Get first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to JSON
            excelData = XLSX.utils.sheet_to_json(worksheet);
            
            // Extract headers
            if (excelData.length > 0) {
                excelHeaders = Object.keys(excelData[0]);
                
                // Display preview
                displayExcelPreview(excelData);
                
                // Show preview section
                document.getElementById('preview-section').style.display = 'block';
            }
            
            // Check if we can enable mapping
            checkEnableMappingSection();
            
            // Auto-update debug info if it's visible
            updateDebugInfoIfVisible();
        } catch (error) {
            console.error('Error processing Excel file:', error);
            alert('Lỗi khi xử lý tệp Excel. Vui lòng kiểm tra định dạng và thử lại.');
            
            // Auto-update debug info if it's visible
            updateDebugInfoIfVisible();
        }
    };
    
    reader.readAsArrayBuffer(file);
}

/**
 * Handle Word template upload
 */
function handleWordTemplateUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    // Check file extension
    const fileName = file.name.toLowerCase();
    if (!fileName.endsWith('.docx') && !fileName.endsWith('.doc')) {
        alert('Vui lòng tải lên tệp .docx hoặc .doc (định dạng Word). Các định dạng khác không được hỗ trợ.');
        event.target.value = ''; // Clear the file input
        document.getElementById('word-template-info').textContent = 'Chưa chọn tệp tin';
        templateName = ''; // Clear template name
        
        // Auto-update debug info if it's visible
        updateDebugInfoIfVisible();
        return;
    }

    // Update file info display
    document.getElementById('word-template-info').textContent = `Đã chọn: ${file.name}`;
    
    // Store template name for filename prefix (remove extension and clean for filename)
    templateName = file.name.replace(/\.[^/.]+$/, "").replace(/[<>:"/\\|?*]/g, '_').replace(/\s+/g, '_');
    
    // Check if this is a .doc file that needs conversion
    if (fileName.endsWith('.doc')) {
        // Show conversion process for .doc files
        convertDocToDocx(file);
    } else {
        // Process .docx file normally
        processDocxFile(file);
    }
}

/**
 * Convert .doc file to .docx (currently shows user-friendly message)
 */
function convertDocToDocx(file) {
    const conversionProgress = document.getElementById('conversion-progress');
    const progressFill = document.getElementById('conversion-progress-fill');
    const conversionText = document.getElementById('conversion-text');
    
    // Show conversion progress
    conversionProgress.style.display = 'block';
    conversionText.textContent = 'Đang phân tích tệp .doc...';
    progressFill.style.width = '20%';
    
    // Simulate conversion process (this will be replaced with actual conversion later)
    setTimeout(() => {
        conversionText.textContent = 'Đang chuyển đổi định dạng...';
        progressFill.style.width = '60%';
        
        setTimeout(() => {
            conversionText.textContent = 'Hoàn tất chuyển đổi!';
            progressFill.style.width = '100%';
            
            setTimeout(() => {
                // Hide conversion progress
                conversionProgress.style.display = 'none';
                progressFill.style.width = '0%';
                
                // For now, we'll attempt to read the .doc file directly
                // In a future version, this will use actual conversion
                attemptDocProcessing(file);
            }, 1000);
        }, 1500);
    }, 1000);
}

/**
 * Attempt to process .doc file directly (fallback method)
 */
function attemptDocProcessing(file) {
    // Add warning about .doc file limitations
    const existingWarning = document.querySelector('.doc-warning');
    if (!existingWarning) {
        const warningDiv = document.createElement('div');
        warningDiv.className = 'doc-warning';
        warningDiv.style.cssText = `
            margin-top: 10px;
            padding: 10px;
            background-color: #fff3cd;
            border: 1px solid #ffeaa7;
            border-radius: 4px;
            font-size: 0.85rem;
            color: #856404;
        `;
        warningDiv.innerHTML = `
            <strong>Lưu ý:</strong> Tệp .doc có thể có một số hạn chế trong việc xử lý. 
            Để có kết quả tốt nhất, vui lòng chuyển đổi sang .docx trong Microsoft Word.
        `;
        
        const uploadBox = document.querySelector('.upload-box:last-of-type');
        uploadBox.appendChild(warningDiv);
    }
    
    // Try to process the .doc file as if it were a .docx file
    // This may work for some newer .doc files that are actually .docx files with .doc extension
    processDocxFile(file);
}

/**
 * Process .docx file (extracted from original handleWordTemplateUpload)
 */
function processDocxFile(file) {
    // Read the Word template file
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const arrayBuffer = e.target.result;
            wordTemplateContent = arrayBuffer;
            
            // Extract placeholders from template
            extractPlaceholders(arrayBuffer);
            
            // Extract actual document content for preview
            extractActualWordContent(arrayBuffer);
            
            // Make sure the preview section is visible
            document.getElementById('preview-section').style.display = 'block';
            document.getElementById('word-preview-section').style.display = 'block';
            
            // Check if we can enable mapping
            checkEnableMappingSection();
            
            // Auto-update debug info if it's visible
            updateDebugInfoIfVisible();
        } catch (error) {
            console.error('Error processing Word template:', error);
            
            // Provide specific error message for format issues
            if (error.message && error.message.includes('zip')) {
                alert('Tệp tải lên không phải định dạng Word hợp lệ. Vui lòng đảm bảo bạn đang tải lên tệp .docx hoặc .doc hợp lệ. Nếu bạn có tệp .doc cũ, hãy thử mở trong Word và lưu lại.');
            } else {
                alert('Lỗi khi xử lý mẫu Word. Vui lòng kiểm tra định dạng và thử lại.');
            }
            
            // Clear the file input
            event.target.value = '';
            document.getElementById('word-template-info').textContent = 'Chưa chọn tệp tin';
            templateName = ''; // Clear template name
            
            // Remove any existing warnings
            const warningDiv = document.querySelector('.doc-warning');
            if (warningDiv) {
                warningDiv.remove();
            }
            
            // Auto-update debug info if it's visible
            updateDebugInfoIfVisible();
        }
    };
    
    reader.readAsArrayBuffer(file);
}

/**
 * Extract placeholders from Word document
 */
function extractPlaceholders(arrayBuffer) {
    try {
        // Load the document using PizZip and docxtemplater
        const zip = new PizZip(arrayBuffer);
        
        // First try with docxtemplater
        try {
            const doc = new docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
            });
            
            // Get the template variables
            const templateVars = doc.getTemplateVars();
            if (templateVars && templateVars.length > 0) {
                wordPlaceholders = [...new Set(templateVars)]; // Remove duplicates
                console.log('Found placeholders with docxtemplater:', wordPlaceholders);
            }
        } catch (docxError) {
            console.warn('Error using docxtemplater to extract placeholders:', docxError);
        }
        
        // Always try manual search to get more comprehensive results
        try {
            const documentXml = zip.file('word/document.xml').asText();
            console.log('Document XML length:', documentXml.length);
            
            // Multiple regex patterns to catch different placeholder formats
            const placeholderPatterns = [
                /{([^{}]+)}/g,           // Standard {placeholder}
                /\{([^}]+)\}/g,          // Alternative pattern
                /\{\s*([^}]+?)\s*\}/g,   // With potential whitespace
            ];
            
            const foundPlaceholders = new Set();
            
            // Try each pattern
            placeholderPatterns.forEach((pattern, index) => {
                let match;
                while ((match = pattern.exec(documentXml)) !== null) {
                    const placeholder = match[1].trim();
                    if (placeholder && placeholder.length > 0 && !placeholder.includes('<') && !placeholder.includes('>')) {
                        foundPlaceholders.add(placeholder);
                        console.log(`Pattern ${index + 1} found placeholder:`, placeholder);
                    }
                }
                // Reset regex
                pattern.lastIndex = 0;
            });
            
            // Also search for placeholders in plain text (for debugging)
            const textContent = documentXml.replace(/<[^>]*>/g, '');
            const textPattern = /{([^{}]+)}/g;
            let textMatch;
            while ((textMatch = textPattern.exec(textContent)) !== null) {
                const placeholder = textMatch[1].trim();
                if (placeholder && placeholder.length > 0) {
                    foundPlaceholders.add(placeholder);
                    console.log('Text search found placeholder:', placeholder);
                }
            }
            
            if (foundPlaceholders.size > 0) {
                // Merge with existing placeholders from docxtemplater
                const allPlaceholders = new Set([...wordPlaceholders, ...foundPlaceholders]);
                wordPlaceholders = Array.from(allPlaceholders);
                console.log('Manually found placeholders:', Array.from(foundPlaceholders));
                console.log('Combined placeholders:', wordPlaceholders);
            } else {
                console.log('No placeholders found in document XML');
                // Log a sample of the XML for debugging
                console.log('XML sample (first 500 chars):', documentXml.substring(0, 500));
            }
        } catch (xmlError) {
            console.error('Error manually searching for placeholders:', xmlError);
        }
        
        // If still no placeholders found, try searching in all document parts
        if (!wordPlaceholders || wordPlaceholders.length === 0) {
            try {
                console.log('Searching in all document parts...');
                const allFiles = Object.keys(zip.files);
                console.log('Available files in document:', allFiles);
                
                // Search in headers, footers, and other parts
                const searchFiles = allFiles.filter(fileName => 
                    fileName.includes('.xml') && 
                    (fileName.includes('header') || fileName.includes('footer') || fileName.includes('document'))
                );
                
                const allFoundPlaceholders = new Set();
                searchFiles.forEach(fileName => {
                    try {
                        const fileContent = zip.file(fileName).asText();
                        const pattern = /{([^{}]+)}/g;
                        let match;
                        while ((match = pattern.exec(fileContent)) !== null) {
                            const placeholder = match[1].trim();
                            if (placeholder && placeholder.length > 0 && !placeholder.includes('<') && !placeholder.includes('>')) {
                                allFoundPlaceholders.add(placeholder);
                                console.log(`Found in ${fileName}:`, placeholder);
                            }
                        }
                    } catch (e) {
                        // Skip files that can't be read as text
                    }
                });
                
                if (allFoundPlaceholders.size > 0) {
                    wordPlaceholders = Array.from(allFoundPlaceholders);
                    console.log('Found placeholders in document parts:', wordPlaceholders);
                }
            } catch (partError) {
                console.error('Error searching document parts:', partError);
            }
        }
        
        // Extract text content from Word document for preview
        extractActualWordContent(arrayBuffer);
        
    } catch (error) {
        console.error('Error extracting placeholders:', error);
        wordPlaceholders = [];
        
        // Still try to show a preview even on error
        extractWordContent(null);
    }
}

/**
 * Extract actual text content from Word document for preview using mammoth.js
 */
function extractActualWordContent(arrayBuffer) {
    try {
        // Use mammoth.js to extract HTML content from the Word document
        mammoth.convertToHtml({arrayBuffer: arrayBuffer})
            .then(function(result) {
                const htmlContent = result.value; // The generated HTML
                const messages = result.messages; // Any messages, such as warnings during conversion
                
                if (messages.length > 0) {
                    console.log('Mammoth conversion messages:', messages);
                }
                
                // Display the actual content with highlighted placeholders
                displayWordPreview(htmlContent, true); // true indicates this is HTML content
            })
            .catch(function(error) {
                console.error('Error converting Word document with mammoth.js:', error);
                // Fallback to simplified preview
                extractWordContent(null);
            });
    } catch (error) {
        console.error('Error using mammoth.js:', error);
        // Fallback to simplified preview
        extractWordContent(null);
    }
}

/**
 * Extract text content from Word document for preview (fallback method)
 */
function extractWordContent(zip) {
    try {
        // Create a placeholder representation as fallback
        let previewContent = "Document Preview (Simplified View)\n\n";
        previewContent += "This is a simplified preview showing detected placeholders.\n";
        previewContent += "The actual document layout may differ.\n\n";
        previewContent += "------------------------------------------\n\n";
        
        // Add some dummy text with placeholders to visualize
        previewContent += "Dear {name},\n\n";
        previewContent += "Thank you for your interest in our services.\n\n";
        
        // Add all detected placeholders in sample text
        if (wordPlaceholders && wordPlaceholders.length > 0) {
            previewContent += "The following placeholders were detected in your document:\n\n";
            
            wordPlaceholders.forEach(placeholder => {
                previewContent += `This document contains placeholder: {${placeholder}}\n`;
            });
        } else {
            previewContent += "No placeholders were detected in your document.\n";
            previewContent += "Please ensure placeholders are formatted as {placeholder_name}\n";
        }
        
        previewContent += "\n------------------------------------------\n";
        previewContent += "Click on any highlighted placeholder to jump to its mapping field.";
        
        // Display the content with highlighted placeholders
        displayWordPreview(previewContent, false); // false indicates this is plain text
    } catch (error) {
        console.error('Error creating Word preview content:', error);
    }
}

/**
 * Display Word document preview with highlighted placeholders
 */
function displayWordPreview(content, isHtml = false) {
    const wordPreviewSection = document.getElementById('word-preview-section');
    const wordPreviewContent = document.getElementById('word-preview-content');
    
    // Show the preview section
    wordPreviewSection.style.display = 'block';
    
    let highlightedContent = content;
    
    if (isHtml) {
        // Content is already HTML from mammoth.js, just highlight placeholders
        wordPlaceholders.forEach(placeholder => {
            const placeholderPattern = `{${placeholder}}`;
            // Use split and join to avoid issues with special regex characters
            highlightedContent = highlightedContent.split(placeholderPattern).join(
                `<span class="placeholder-highlight" data-placeholder="${placeholder}">{${placeholder}}</span>`
            );
        });
    } else {
        // Escape HTML content for safety (plain text)
        highlightedContent = highlightedContent
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
        
        // Highlight each placeholder
        wordPlaceholders.forEach(placeholder => {
            const placeholderPattern = `{${placeholder}}`;
            // Use split and join to avoid issues with special regex characters
            highlightedContent = highlightedContent.split(placeholderPattern).join(
                `<span class="placeholder-highlight" data-placeholder="${placeholder}">{${placeholder}}</span>`
            );
        });
        
        // Preserve line breaks for proper display
        highlightedContent = highlightedContent.replace(/\n/g, '<br>');
    }
    
    // Set the content
    wordPreviewContent.innerHTML = highlightedContent;
    
    // Add click event to placeholders
    const highlightedPlaceholders = document.querySelectorAll('.placeholder-highlight');
    highlightedPlaceholders.forEach(elem => {
        elem.addEventListener('click', function() {
            const placeholder = this.getAttribute('data-placeholder');
            
            // Find and focus the corresponding mapping field
            const selectField = document.querySelector(`select[data-placeholder="${placeholder}"]`);
            if (selectField) {
                selectField.scrollIntoView({ behavior: 'smooth', block: 'center' });
                selectField.focus();
                selectField.classList.add('highlight-field');
                
                // Remove highlight after a short delay
                setTimeout(() => {
                    selectField.classList.remove('highlight-field');
                }, 2000);
            }
        });
    });
    
    console.log('Word preview displayed with', highlightedPlaceholders.length, 'highlighted placeholders');
}

/**
 * Check if we can enable the mapping section
 */
function checkEnableMappingSection() {
    if (excelHeaders.length > 0 && wordPlaceholders.length > 0) {
        createMappingFields();
        document.getElementById('mapping-section').style.display = 'block';
        document.getElementById('generate-btn').disabled = false;
        
        // Update button text to show how many documents will be generated
        const generateBtn = document.getElementById('generate-btn');
        const rowCount = excelData ? excelData.length : 0;
        generateBtn.textContent = `Tạo ${rowCount} bản tài liệu`;
        
        // Auto-update debug info if it's visible
        updateDebugInfoIfVisible();
    }
}

/**
 * Create mapping fields between Excel columns and Word placeholders
 */
function createMappingFields() {
    const mappingContainer = document.getElementById('mapping-container');
    mappingContainer.innerHTML = '';
    
    wordPlaceholders.forEach(placeholder => {
        const mappingItem = document.createElement('div');
        mappingItem.className = 'mapping-item';
        
        const label = document.createElement('label');
        label.textContent = `Ánh xạ "${placeholder}" với:`;
        label.setAttribute('for', `map-${placeholder}`);
        
        const select = document.createElement('select');
        select.id = `map-${placeholder}`;
        select.name = `map-${placeholder}`;
        select.setAttribute('data-placeholder', placeholder);
        
        // Add event listener to update debug info when mapping changes
        select.addEventListener('change', function() {
            updateDebugInfoIfVisible();
        });
        
        // Add empty option
        const emptyOption = document.createElement('option');
        emptyOption.value = '';
        emptyOption.textContent = '-- Chọn cột Excel --';
        select.appendChild(emptyOption);
        
        // Add options for each Excel header
        excelHeaders.forEach(header => {
            const option = document.createElement('option');
            option.value = header;
            option.textContent = header;
            
            // Auto-select if names match
            if (header.toLowerCase() === placeholder.toLowerCase()) {
                option.selected = true;
            }
            
            select.appendChild(option);
        });
        
        mappingItem.appendChild(label);
        mappingItem.appendChild(select);
        mappingContainer.appendChild(mappingItem);
    });
    
    // Auto-update debug info if it's visible
    updateDebugInfoIfVisible();
}

/**
 * Update debug info if the debug area is currently visible
 */
function updateDebugInfoIfVisible() {
    const debugArea = document.getElementById('debug-area');
    if (debugArea && debugArea.style.display !== 'none') {
        generateDebugInfo();
    }
}

/**
 * Display Excel preview
 */
function displayExcelPreview(data) {
    const previewContainer = document.getElementById('excel-preview');
    previewContainer.innerHTML = '';
    
    if (data.length === 0) {
        previewContainer.innerHTML = '<p>No data found in Excel file</p>';
        return;
    }
    
    // Create controls for showing more/less data
    const controlsDiv = document.createElement('div');
    controlsDiv.className = 'excel-controls';
    
    const showRowsSelect = document.createElement('select');
    showRowsSelect.id = 'show-rows-select';
    [5, 10, 25, 50, 100, 'All'].forEach(value => {
        const option = document.createElement('option');
        option.value = value;
        option.textContent = value === 'All' ? `All (${data.length} rows)` : `${value} rows`;
        if (value === 10) option.selected = true; // Default to 10 rows
        showRowsSelect.appendChild(option);
    });
    
    const label = document.createElement('label');
    label.textContent = 'Show: ';
    label.appendChild(showRowsSelect);
    
    controlsDiv.appendChild(label);
    previewContainer.appendChild(controlsDiv);
    
    // Create table container with scroll
    const tableContainer = document.createElement('div');
    tableContainer.className = 'table-container';
    tableContainer.style.maxHeight = '400px';
    tableContainer.style.overflowY = 'auto';
    
    // Create table
    const table = document.createElement('table');
    
    // Create table header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    
    excelHeaders.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // Function to update table content
    function updateTableContent(rowsToShow) {
        // Remove existing tbody if any
        const existingTbody = table.querySelector('tbody');
        if (existingTbody) {
            existingTbody.remove();
        }
        
        // Create table body
        const tbody = document.createElement('tbody');
        
        // Determine how many rows to show
        const maxRows = rowsToShow === 'All' ? data.length : parseInt(rowsToShow);
        const previewData = data.slice(0, maxRows);
        
        previewData.forEach((row, index) => {
            const tr = document.createElement('tr');
            
            // Add row number column
            const rowNumTd = document.createElement('td');
            rowNumTd.textContent = index + 1;
            rowNumTd.className = 'row-number';
            tr.appendChild(rowNumTd);
            
            excelHeaders.forEach(header => {
                const td = document.createElement('td');
                const cellValue = row[header];
                
                // Handle different data types and formatting
                if (cellValue !== undefined && cellValue !== null) {
                    if (typeof cellValue === 'number') {
                        td.textContent = cellValue.toLocaleString();
                    } else if (cellValue instanceof Date) {
                        td.textContent = cellValue.toLocaleDateString();
                    } else {
                        td.textContent = cellValue.toString();
                    }
                } else {
                    td.textContent = '';
                    td.className = 'empty-cell';
                }
                
                tr.appendChild(td);
            });
            
            tbody.appendChild(tr);
        });
        
        table.appendChild(tbody);
        
        // Update row count info
        const existingRowCount = previewContainer.querySelector('.row-count-info');
        if (existingRowCount) {
            existingRowCount.remove();
        }
        
        const rowCountInfo = document.createElement('div');
        rowCountInfo.className = 'row-count-info';
        rowCountInfo.innerHTML = `
            <p><strong>Showing ${previewData.length} of ${data.length} rows</strong></p>
            <p>Columns: ${excelHeaders.length} | Total cells: ${data.length * excelHeaders.length}</p>
        `;
        previewContainer.appendChild(rowCountInfo);
    }
    
    // Add row number header
    const rowNumTh = document.createElement('th');
    rowNumTh.textContent = '#';
    rowNumTh.className = 'row-number-header';
    headerRow.insertBefore(rowNumTh, headerRow.firstChild);
    
    // Initial table content
    updateTableContent(10);
    
    tableContainer.appendChild(table);
    previewContainer.appendChild(tableContainer);
    
    // Add event listener for changing number of rows
    showRowsSelect.addEventListener('change', function() {
        updateTableContent(this.value);
    });
}

/**
 * Generate document with mapped data
 */
function generateDocument() {
    try {
        // Check if we have required data
        if (!excelData || !wordTemplateContent || excelData.length === 0) {
            alert('Vui lòng tải lên cả dữ liệu Excel và mẫu Word trước khi tạo tài liệu.');
            return;
        }
        
        if (!wordPlaceholders || wordPlaceholders.length === 0) {
            alert('Không tìm thấy placeholder trong mẫu Word. Vui lòng đảm bảo mẫu của bạn chứa placeholder theo định dạng {tên_placeholder}.');
            return;
        }
        
        // Get mapping configuration
        const mappingConfig = {};
        let mappedCount = 0;
        
        wordPlaceholders.forEach(placeholder => {
            const select = document.querySelector(`select[data-placeholder="${placeholder}"]`);
            if (select && select.value) {
                mappingConfig[placeholder] = select.value;
                mappedCount++;
            }
        });
        
        if (mappedCount === 0) {
            alert('Vui lòng ánh xạ ít nhất một cột Excel với một placeholder trong Word trước khi tạo tài liệu.');
            return;
        }
        
        console.log('Mapping configuration:', mappingConfig);
        console.log('Number of mapped placeholders:', mappedCount);
        console.log('Generating documents for', excelData.length, 'rows');
        
        // Show progress
        const downloadArea = document.getElementById('download-area');
        downloadArea.style.display = 'block';
        downloadArea.innerHTML = `
            <h3>Đang tạo ${excelData.length} tài liệu...</h3>
            <div class="progress-bar">
                <div class="progress-fill" id="progress-fill"></div>
            </div>
            <p id="progress-text">Đang bắt đầu...</p>
        `;
        
        // Scroll to download area
        downloadArea.scrollIntoView({ behavior: 'smooth' });
        
        // Generate documents for each row
        generateMultipleDocuments(mappingConfig);
        
        // Auto-update debug info if it's visible
        updateDebugInfoIfVisible();
        
    } catch (error) {
        console.error('Unexpected error generating document:', error);
        
        // Provide more specific error message based on the error type
        let errorMessage = 'Lỗi khi tạo tài liệu. ';
        
        if (error.message) {
            if (error.message.includes('corrupted')) {
                errorMessage += 'Mẫu Word có vẻ bị hỏng. Vui lòng thử tệp khác.';
            } else if (error.message.includes('placeholder') || error.message.includes('template')) {
                errorMessage += 'Có vẻ có vấn đề với placeholder trong mẫu của bạn. Vui lòng đảm bảo chúng được định dạng như {tên_placeholder}.';
            } else if (error.message.includes('zip') || error.message.includes('archive')) {
                errorMessage += 'Không thể xử lý mẫu Word. Vui lòng đảm bảo đó là tệp .docx hợp lệ.';
            } else {
                errorMessage += 'Vui lòng kiểm tra ánh xạ và thử lại. Lỗi: ' + error.message;
            }
        } else {
            errorMessage += 'Vui lòng kiểm tra ánh xạ và thử lại.';
        }
        
        alert(errorMessage);
        
        // Auto-update debug info if it's visible
        updateDebugInfoIfVisible();
    }
}

/**
 * Generate multiple documents - one for each Excel row
 */
async function generateMultipleDocuments(mappingConfig) {
    const progressFill = document.getElementById('progress-fill');
    const progressText = document.getElementById('progress-text');
    const downloadArea = document.getElementById('download-area');
    
    const generatedFiles = [];
    const totalRows = excelData.length;
    
    try {
        // Process each row
        for (let i = 0; i < totalRows; i++) {
            const dataRow = excelData[i];
            
            // Update progress
            const progress = ((i + 1) / totalRows) * 100;
            progressFill.style.width = progress + '%';
            progressText.textContent = `Đang xử lý ${i + 1} / ${totalRows}...`;
            
            // Allow UI to update
            await new Promise(resolve => setTimeout(resolve, 50));
            
            try {
                // Create a completely fresh copy of the template content
                const templateCopy = wordTemplateContent.slice(0);
                
                // Create a fresh PizZip instance from the template copy
                const zip = new PizZip(templateCopy);
                
                // Prepare template data for this row
                const templateData = {};
                
                // Map Excel data to template placeholders
                Object.keys(mappingConfig).forEach(placeholder => {
                    const excelColumn = mappingConfig[placeholder];
                    let value = dataRow[excelColumn];
                    
                    // Handle different data types
                    if (value === null || value === undefined) {
                        value = '';
                    } else if (typeof value === 'number') {
                        value = value.toString();
                    } else if (value instanceof Date) {
                        value = value.toLocaleDateString();
                    } else {
                        value = value.toString();
                    }
                    
                    templateData[placeholder] = value;
                });
                
                console.log(`Row ${i + 1} template data:`, templateData);
                
                // Create new instance of docxtemplater with fresh zip
                const doc = new docxtemplater(zip, {
                    paragraphLoop: true,
                    linebreaks: true,
                    nullGetter: function(part) {
                        // Return empty string for null/undefined values
                        return '';
                    },
                    errorLogging: true
                });
                
                // Set data for template
                doc.setData(templateData);
                
                // Render document
                doc.render();
                
                // Get the generated ZIP
                const generatedZip = doc.getZip();
                
                // Generate blob with proper settings
                const out = generatedZip.generate({
                    type: 'blob',
                    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    compression: 'DEFLATE',
                    compressionOptions: {
                        level: 6
                    }
                });
                
                // Verify the blob is valid
                if (!out || out.size === 0) {
                    throw new Error('Generated document is empty');
                }
                
                // Create filename based on row data or index
                let filename = `document_${i + 1}`;
                
                // Try to use a meaningful name if available
                const nameColumns = ['name', 'Name', 'Họ và tên', 'Họ và tên ', 'STT'];
                for (const col of nameColumns) {
                    if (dataRow[col]) {
                        const nameValue = dataRow[col].toString().trim();
                        if (nameValue) {
                            // Clean filename (remove invalid characters and replace spaces with underscores)
                            filename = nameValue.replace(/[<>:"/\\|?*]/g, '_').replace(/\s+/g, '_');
                            break;
                        }
                    }
                }
                
                // Add template name prefix
                if (templateName) {
                    filename = `${templateName}_${filename}`;
                }
                
                // Format date as YYYY_MM_DD
                const today = new Date();
                const formattedDate = `${today.getFullYear()}_${String(today.getMonth() + 1).padStart(2, '0')}_${String(today.getDate()).padStart(2, '0')}`;
                filename += `_${formattedDate}.docx`;
                
                // Store file info
                generatedFiles.push({
                    blob: out,
                    filename: filename,
                    rowIndex: i + 1,
                    rowData: templateData
                });
                
                console.log(`Successfully generated document ${i + 1}/${totalRows}: ${filename}`);
                
            } catch (rowError) {
                console.error(`Error generating document for row ${i + 1}:`, rowError);
                generatedFiles.push({
                    error: rowError.message,
                    rowIndex: i + 1,
                    filename: `error_row_${i + 1}.txt`
                });
            }
        }
        
        // Update progress to complete
        progressFill.style.width = '100%';
        progressText.textContent = 'Hoàn thành tạo tài liệu!';
        
        // Display download links
        displayDownloadLinks(generatedFiles);
        
        // Auto-update debug info if it's visible
        updateDebugInfoIfVisible();
        
    } catch (error) {
        console.error('Error in batch generation:', error);
        downloadArea.innerHTML = `
            <h3>Lỗi khi tạo tài liệu</h3>
            <p>Đã xảy ra lỗi trong quá trình tạo hàng loạt: ${error.message}</p>
        `;
        
        // Auto-update debug info if it's visible
        updateDebugInfoIfVisible();
    }
}

/**
 * Display download links for all generated documents
 */
function displayDownloadLinks(generatedFiles) {
    const downloadArea = document.getElementById('download-area');
    
    const successfulFiles = generatedFiles.filter(file => !file.error);
    const errorFiles = generatedFiles.filter(file => file.error);
    
    let html = `
        <h3>Đã tạo thành công ${successfulFiles.length} tài liệu!</h3>
    `;
    
    if (successfulFiles.length > 1) {
        html += `
            <div class="download-all-section">
                <button id="download-all-btn" class="download-btn">Tải xuống tất cả dưới dạng ZIP</button>
            </div>
            <hr>
        `;
    }
    
    html += `<div class="individual-downloads">`;
    
    successfulFiles.forEach((file, index) => {
        const url = URL.createObjectURL(file.blob);
        html += `
            <div class="download-item">
                <span class="file-info"> ${file.filename}</span>
                <div class="download-actions">
                    <a href="${url}" download="${file.filename}" class="download-link-small">Tải xuống</a>
                    <button onclick="testDocumentOpen('${url}', '${file.filename}')" class="test-btn" title="Kiểm tra xem tài liệu có mở được không">Kiểm tra</button>
                </div>
            </div>
        `;
    });
    
    html += `</div>`;
    
    if (errorFiles.length > 0) {
        html += `
            <div class="error-section">
                <h4>Lỗi (${errorFiles.length} tệp):</h4>
        `;
        
        errorFiles.forEach(file => {
            html += `
                <div class="error-item">
                    Hàng ${file.rowIndex}: ${file.error}
                </div>
            `;
        });
        
        html += `</div>`;
    }
    
    downloadArea.innerHTML = html;
    
    // Add event listener for download all button
    if (successfulFiles.length > 1) {
        document.getElementById('download-all-btn').addEventListener('click', () => {
            downloadAllAsZip(successfulFiles);
        });
    }
}

/**
 * Download all files as a ZIP archive
 */
async function downloadAllAsZip(files) {
    try {
        // Create a new ZIP file using JSZip
        const zip = new JSZip();
        
        files.forEach(file => {
            zip.file(file.filename, file.blob);
        });
        
        // Generate the ZIP file
        const zipBlob = await zip.generateAsync({
            type: 'blob',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        });
        
        // Create download link
        const url = URL.createObjectURL(zipBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `generated_documents_${new Date().toISOString().slice(0, 10)}.zip`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
    } catch (error) {
        console.error('Error creating ZIP file:', error);
        alert('Lỗi khi tạo tệp ZIP. Vui lòng tải xuống từng tệp riêng lẻ.');
    }
}

/**
 * Show debug information to help troubleshoot issues
 */
function showDebugInfo() {
    const debugArea = document.getElementById('debug-area');
    
    // Toggle visibility and generate new content
    if (debugArea.style.display === 'none' || debugArea.style.display === '') {
        debugArea.style.display = 'block';
        generateDebugInfo();
    } else {
        debugArea.style.display = 'none';
    }
}

/**
 * Generate and display debug information
 */
function generateDebugInfo() {
    const debugContent = document.getElementById('debug-content');
    
    let debugInfo = 'THÔNG TIN DEBUG\n';
    debugInfo += '='.repeat(50) + '\n\n';
    
    // Excel data info
    debugInfo += '📊 DỮ LIỆU EXCEL:\n';
    if (excelData && excelData.length > 0) {
        debugInfo += `- Số hàng: ${excelData.length}\n`;
        debugInfo += `- Số cột: ${excelHeaders.length}\n`;
        debugInfo += `- Tên cột: ${excelHeaders.join(', ')}\n`;
        debugInfo += `- Dữ liệu hàng đầu tiên: ${JSON.stringify(excelData[0], null, 2)}\n`;
    } else {
        debugInfo += '- Chưa tải dữ liệu Excel\n';
    }
    
    debugInfo += '\n📄 MẪU WORD:\n';
    if (wordTemplateContent) {
        debugInfo += `- Đã tải mẫu: Có (${(wordTemplateContent.byteLength / 1024).toFixed(2)} KB)\n`;
        debugInfo += `- Tên mẫu: ${templateName || 'Không xác định'}\n`;
        debugInfo += `- Số placeholder tìm thấy: ${wordPlaceholders.length}\n`;
        if (wordPlaceholders.length > 0) {
            debugInfo += `- Danh sách placeholder: ${wordPlaceholders.join(', ')}\n`;
        } else {
            debugInfo += '- Không phát hiện placeholder!\n';
            debugInfo += '\n🔍 PHÂN TÍCH PHÁT HIỆN PLACEHOLDER:\n';
            
            // Try to analyze the document for common issues
            try {
                const zip = new PizZip(wordTemplateContent);
                const documentXml = zip.file('word/document.xml').asText();
                
                // Check for various bracket formats
                const bracketChecks = [
                    { name: 'Dấu ngoặc nhọn {}', pattern: /{[^}]*}/g },
                    { name: 'Dấu ngoặc vuông []', pattern: /\[[^\]]*\]/g },
                    { name: 'Dấu ngoặc nhọn <>', pattern: /<[^>]*>/g },
                    { name: 'Dấu ngoặc tròn ()', pattern: /\([^)]*\)/g }
                ];
                
                bracketChecks.forEach(check => {
                    const matches = documentXml.match(check.pattern);
                    if (matches && matches.length > 0) {
                        debugInfo += `- Tìm thấy ${matches.length} trường hợp ${check.name}\n`;
                        if (check.name.includes('ngoặc nhọn {}')) {
                            debugInfo += `  Ví dụ: ${matches.slice(0, 5).join(', ')}\n`;
                        }
                    }
                });
                
                // Check document structure
                const hasDocumentBody = documentXml.includes('<w:body>');
                const hasParagraphs = documentXml.includes('<w:p>');
                const hasText = documentXml.includes('<w:t>');
                
                debugInfo += `- Cấu trúc tài liệu: Body=${hasDocumentBody}, Đoạn văn=${hasParagraphs}, Văn bản=${hasText}\n`;
                
                // Extract visible text content
                const textElements = documentXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g);
                if (textElements && textElements.length > 0) {
                    const visibleText = textElements.map(el => el.replace(/<[^>]*>/g, '')).join(' ');
                    debugInfo += `- Mẫu văn bản có thể nhìn thấy: "${visibleText.substring(0, 200)}..."\n`;
                    
                    // Check if the visible text contains potential placeholders
                    const textPlaceholders = visibleText.match(/{[^}]*}/g);
                    if (textPlaceholders) {
                        debugInfo += `- Placeholder tiềm năng trong văn bản: ${textPlaceholders.join(', ')}\n`;
                    }
                } else {
                    debugInfo += '- Không tìm thấy văn bản hiển thị trong tài liệu\n';
                }
                
            } catch (analysisError) {
                debugInfo += `- Lỗi khi phân tích tài liệu: ${analysisError.message}\n`;
                
                // Check for common format issues
                if (analysisError.message.includes('zip') || analysisError.message.includes('central directory')) {
                    debugInfo += '\n❌ PHÁT HIỆN VẤN ĐỀ ĐỊNH DẠNG TẬP TIN:\n';
                    debugInfo += '- Đây có vẻ là tệp .doc (định dạng Word cũ)\n';
                    debugInfo += '- Chỉ hỗ trợ tệp .docx (định dạng Word mới)\n';
                    debugInfo += '- GIẢI PHÁP: Mở tệp .doc trong Microsoft Word\n';
                    debugInfo += '- Vào File > Save As > Chọn định dạng "Word Document (.docx)"\n';
                    debugInfo += '- Sau đó tải lên tệp .docx mới\n';
                }
            }
        }
    } else {
        debugInfo += '- Chưa tải mẫu Word\n';
    }
    
    debugInfo += '\n🔗 CẤU HÌNH ÁNH XẠ:\n';
    const mappingConfig = {};
    let mappedCount = 0;
    
    wordPlaceholders.forEach(placeholder => {
        const select = document.querySelector(`select[data-placeholder="${placeholder}"]`);
        if (select) {
            const value = select.value;
            mappingConfig[placeholder] = value || '(chưa ánh xạ)';
            if (value) mappedCount++;
        }
    });
    
    debugInfo += `- Placeholder đã ánh xạ: ${mappedCount}/${wordPlaceholders.length}\n`;
    if (Object.keys(mappingConfig).length > 0) {
        debugInfo += `- Chi tiết ánh xạ:\n`;
        Object.keys(mappingConfig).forEach(placeholder => {
            debugInfo += `  • ${placeholder} → ${mappingConfig[placeholder]}\n`;
        });
    }
    
    debugInfo += '\n🔧 TRẠNG THÁI THƯ VIỆN:\n';
    debugInfo += `- Thư viện XLSX: ${typeof XLSX !== 'undefined' ? 'Đã tải' : 'Thiếu'}\n`;
    debugInfo += `- docxtemplater: ${typeof docxtemplater !== 'undefined' ? 'Đã tải' : 'Thiếu'}\n`;
    debugInfo += `- PizZip: ${typeof PizZip !== 'undefined' ? 'Đã tải' : 'Thiếu'}\n`;
    debugInfo += `- mammoth: ${typeof mammoth !== 'undefined' ? 'Đã tải' : 'Thiếu'}\n`;
    
    debugInfo += '\n💡 GỢI Ý:\n';
    if (!excelData || excelData.length === 0) {
        debugInfo += '- Tải lên tệp Excel có dữ liệu\n';
    }
    if (!wordTemplateContent) {
        debugInfo += '- Tải lên mẫu Word có placeholder\n';
        debugInfo += '- Đảm bảo sử dụng định dạng .docx (không phải .doc)\n';
    }
    if (wordPlaceholders.length === 0) {
        debugInfo += '- Đảm bảo mẫu Word có placeholder theo định dạng {tên_placeholder}\n';
        debugInfo += '- Kiểm tra rằng placeholder không bị tách ra thành nhiều đoạn văn bản trong Word\n';
        debugInfo += '- Thử tạo placeholder bằng cách gõ trực tiếp (không copy-paste)\n';
        debugInfo += '- Tránh định dạng đặc biệt trong văn bản placeholder\n';
        debugInfo += '- Nếu đang dùng định dạng .doc, hãy chuyển đổi sang .docx trước\n';
    }
    if (mappedCount === 0 && wordPlaceholders.length > 0) {
        debugInfo += '- Ánh xạ ít nhất một cột Excel với một placeholder trong Word\n';
    }
    
    if (excelData && wordTemplateContent && wordPlaceholders.length > 0 && mappedCount > 0) {
        debugInfo += '- Tất cả yêu cầu đã đáp ứng! Thử tạo tài liệu.\n';
    }
    
    debugContent.textContent = debugInfo;
}

/**
 * Test if a document can be opened (basic validation)
 */
function testDocumentOpen(url, filename) {
    // Try to create a temporary link and test
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    
    // Simple test - just try to download and show info
    fetch(url)
        .then(response => response.blob())
        .then(blob => {
            const isValidSize = blob.size > 1000; // At least 1KB
            const hasCorrectType = blob.type.includes('wordprocessingml') || blob.type.includes('document');
            
            let message = `Tệp: ${filename}\n`;
            message += `Kích thước: ${(blob.size / 1024).toFixed(1)} KB\n`;
            message += `Loại: ${blob.type}\n`;
            message += `Kích thước hợp lệ: ${isValidSize ? '✓' : '✗'}\n`;
            message += `Loại đúng: ${hasCorrectType ? '✓' : '✗'}\n\n`;
            
            if (isValidSize && hasCorrectType) {
                message += 'Tài liệu có vẻ hợp lệ. Thử mở nó!';
            } else {
                message += 'Tài liệu có thể bị hỏng. Kiểm tra mẫu và dữ liệu.';
            }
            
            alert(message);
        })
        .catch(error => {
            alert(`Lỗi khi kiểm tra tài liệu: ${error.message}`);
        });
}

// Improved file upload/remove logic
function setFileInfo(inputId, infoId, btnId) {
    const input = document.getElementById(inputId);
    const info = document.getElementById(infoId);

    input.addEventListener('change', function () {
        if (this.files.length) {
            info.innerHTML = `
                ${this.files[0].name}
                <button class="remove-file-btn" id="${btnId}" title="Xóa tệp">×</button>
            `;
            // Re-attach remove event after replacing innerHTML
            document.getElementById(btnId).onclick = function () {
                input.value = '';
                info.innerHTML = `
                    Chưa chọn tệp tin
                    <button class="remove-file-btn" id="${btnId}" style="display:none;" title="Xóa tệp">×</button>
                `;
                if (inputId === 'excel-file') {
                    excelData = null;
                    excelHeaders = [];
                    document.getElementById('excel-preview').innerHTML = '';
                }
                if (inputId === 'word-template') {
                    wordTemplateContent = null;
                    wordPlaceholders = [];
                    templateName = '';
                    const wordPreviewContent = document.getElementById('word-preview-content');
                    if (wordPreviewContent) {
                        wordPreviewContent.innerHTML = '';
                    }
                    document.getElementById('word-preview-section').style.display = 'none';
                    
                    // Hide conversion progress if visible
                    const conversionProgress = document.getElementById('conversion-progress');
                    if (conversionProgress) {
                        conversionProgress.style.display = 'none';
                        document.getElementById('conversion-progress-fill').style.width = '0%';
                    }
                    
                    // Remove any existing warnings
                    const warningDiv = document.querySelector('.doc-warning');
                    if (warningDiv) {
                        warningDiv.remove();
                    }
                }
                updateSectionsVisibility();
                updateDebugInfoIfVisible();
            };

            // Process new file upload after removal
            if (inputId === 'excel-file' && this.files[0]) {
                handleExcelUpload({ target: input });
            }
            
            if (inputId === 'word-template' && this.files[0]) {
                handleWordTemplateUpload({ target: input });
            }
            
            updateSectionsVisibility();
        } else {
            info.innerHTML = `
                Chưa chọn tệp tin
                <button class="remove-file-btn" id="${btnId}" style="display:none;" title="Xóa tệp">×</button>
            `;
            if (inputId === 'excel-file') {
                document.getElementById('excel-preview').innerHTML = '';
            }
            if (inputId === 'word-template') {
                const wordPreviewContent = document.getElementById('word-preview-content');
                if (wordPreviewContent) {
                    wordPreviewContent.innerHTML = '';
                }
                document.getElementById('word-preview-section').style.display = 'none';
            }
            updateSectionsVisibility();
        }
    });

    // Initial state
    info.innerHTML = `
        Chưa chọn tệp tin
        <button class="remove-file-btn" id="${btnId}" style="display:none;" title="Xóa tệp">×</button>
    `;
}

// Add this helper function at the end of your file:
function updateSectionsVisibility() {
    const hasExcel = excelData && excelHeaders.length > 0;
    const hasWord = wordTemplateContent && wordPlaceholders.length > 0;

    // Preview section
    if (hasExcel || hasWord) {
        document.getElementById('preview-section').style.display = 'block';
    } else {
        document.getElementById('preview-section').style.display = 'none';
    }

    // Mapping section and generate button only if both files are present
    if (hasExcel && hasWord) {
        document.getElementById('mapping-section').style.display = 'block';
        document.getElementById('generate-btn').disabled = false;
    } else {
        document.getElementById('mapping-section').style.display = 'none';
        document.getElementById('generate-btn').disabled = true;
    }
}
