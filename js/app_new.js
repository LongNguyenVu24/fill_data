// Global variables
let excelData = null;
let excelHeaders = [];
let wordTemplates = {}; // Store multiple templates: {templateName: {content: arrayBuffer, placeholders: []}}
let selectedTemplateForRows = {}; // Store template selection for each row: {rowIndex: templateName}

// Document ready function
document.addEventListener('DOMContentLoaded', function() {
    // Initialize event listeners
    document.getElementById('excel-file').addEventListener('change', handleExcelUpload);
    document.getElementById('word-templates').addEventListener('change', handleWordTemplatesUpload);
    document.getElementById('generate-btn').addEventListener('click', generateDocument);
    document.getElementById('debug-btn').addEventListener('click', showDebugInfo);
});

/**
 * Handle Excel file upload
 */
function handleExcelUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    // Update file info display
    document.getElementById('excel-file-info').textContent = `Selected: ${file.name}`;
    
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
                
                // Display preview with template selection if templates are available
                updateExcelPreviewWithTemplateSelection();
                
                // Show preview section
                document.getElementById('preview-section').style.display = 'block';
                
                // Check if we can enable mapping
                checkEnableMappingSection();
            }
            
        } catch (error) {
            console.error('Error reading Excel file:', error);
            alert('Error reading Excel file. Please check the format and try again.');
        }
    };
    
    reader.readAsArrayBuffer(file);
}

/**
 * Handle multiple Word templates upload
 */
function handleWordTemplatesUpload(event) {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    // Clear previous templates
    wordTemplates = {};
    selectedTemplateForRows = {};
    
    // Reset UI
    document.getElementById('templates-list').style.display = 'none';
    document.getElementById('templates-list-items').innerHTML = '';
    
    let validFileCount = 0;
    let processedFiles = 0;
    const fileList = [];

    // First, validate all files
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const fileName = file.name.toLowerCase();
        
        if (!fileName.endsWith('.docx')) {
            alert(`File "${file.name}" is not a .docx file. Please upload only .docx files.`);
            continue;
        }
        
        fileList.push(file);
        validFileCount++;
    }

    if (validFileCount === 0) {
        alert('No valid .docx files selected. Please select .docx files only.');
        event.target.value = '';
        document.getElementById('word-templates-info').textContent = 'No files selected';
        return;
    }

    // Update file info display
    document.getElementById('word-templates-info').textContent = `Processing ${validFileCount} template(s)...`;

    // Process each valid file
    fileList.forEach((file, index) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const arrayBuffer = e.target.result;
                const templateName = file.name.replace('.docx', '');
                
                // Store template data
                wordTemplates[templateName] = {
                    content: arrayBuffer,
                    fileName: file.name,
                    placeholders: []
                };
                
                // Extract placeholders from this template
                extractPlaceholdersFromTemplate(arrayBuffer, templateName);
                
                processedFiles++;
                
                // Update UI when all files are processed
                if (processedFiles === validFileCount) {
                    updateTemplatesUI();
                    updateExcelPreviewWithTemplateSelection();
                    checkEnableMappingSection();
                }
                
            } catch (error) {
                console.error(`Error processing template "${file.name}":`, error);
                alert(`Error processing template "${file.name}". Please check the format and try again.`);
                
                processedFiles++;
                if (processedFiles === validFileCount) {
                    updateTemplatesUI();
                }
            }
        };
        
        reader.readAsArrayBuffer(file);
    });
}

/**
 * Extract placeholders from a specific template
 */
function extractPlaceholdersFromTemplate(arrayBuffer, templateName) {
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
                wordTemplates[templateName].placeholders = [...new Set(templateVars)]; // Remove duplicates
                console.log(`Placeholders found in ${templateName}:`, wordTemplates[templateName].placeholders);
            } else {
                console.log(`No placeholders found in ${templateName} using docxtemplater.`);
                wordTemplates[templateName].placeholders = [];
            }
        } catch (docxError) {
            console.log(`Docxtemplater failed for ${templateName}, trying manual extraction:`, docxError);
            // Fallback to manual extraction
            extractPlaceholdersManually(zip, templateName);
        }
    } catch (error) {
        console.error(`Error extracting placeholders from ${templateName}:`, error);
        wordTemplates[templateName].placeholders = [];
    }
}

/**
 * Manual placeholder extraction fallback for a specific template
 */
function extractPlaceholdersManually(zip, templateName) {
    try {
        // Extract text from document.xml
        const documentXml = zip.file("word/document.xml").asText();
        
        // Find all placeholders in format {placeholder_name}
        const placeholderRegex = /\{([^}]+)\}/g;
        const matches = documentXml.match(placeholderRegex);
        
        if (matches) {
            // Extract just the placeholder names (without the curly braces)
            const placeholders = matches.map(match => match.slice(1, -1));
            wordTemplates[templateName].placeholders = [...new Set(placeholders)]; // Remove duplicates
            console.log(`Manual extraction found placeholders in ${templateName}:`, wordTemplates[templateName].placeholders);
        } else {
            console.log(`No placeholders found in ${templateName} with manual extraction.`);
            wordTemplates[templateName].placeholders = [];
        }
    } catch (error) {
        console.error(`Manual placeholder extraction failed for ${templateName}:`, error);
        wordTemplates[templateName].placeholders = [];
    }
}

/**
 * Update the templates list UI
 */
function updateTemplatesUI() {
    const templatesList = document.getElementById('templates-list');
    const templatesListItems = document.getElementById('templates-list-items');
    
    templatesListItems.innerHTML = '';
    
    const templateNames = Object.keys(wordTemplates);
    
    if (templateNames.length > 0) {
        templatesList.style.display = 'block';
        
        templateNames.forEach(templateName => {
            const template = wordTemplates[templateName];
            const li = document.createElement('li');
            li.innerHTML = `
                <span class="template-name">${template.fileName}</span>
                <span class="template-placeholders">(${template.placeholders.length} placeholders)</span>
                <button class="preview-template-btn" onclick="previewTemplate('${templateName}')">Preview</button>
            `;
            templatesListItems.appendChild(li);
        });
        
        document.getElementById('word-templates-info').textContent = 
            `${templateNames.length} template(s) uploaded successfully`;
    } else {
        templatesList.style.display = 'none';
        document.getElementById('word-templates-info').textContent = 'No valid templates uploaded';
    }
}

/**
 * Preview a specific template
 */
function previewTemplate(templateName) {
    if (!wordTemplates[templateName]) {
        alert('Template not found!');
        return;
    }
    
    const template = wordTemplates[templateName];
    
    // Show preview section if not already visible
    document.getElementById('preview-section').style.display = 'block';
    document.getElementById('word-preview-section').style.display = 'block';
    
    // Update preview title to show which template is being previewed
    const previewTitle = document.querySelector('#word-preview-section h3');
    previewTitle.textContent = `📄 Word Template Preview: ${template.fileName}`;
    
    // Extract and display content using mammoth
    mammoth.convertToHtml({arrayBuffer: template.content})
        .then(function(result) {
            const htmlContent = result.value;
            displayWordPreview(htmlContent, true, template.placeholders);
        })
        .catch(function(error) {
            console.error('Error converting Word document:', error);
            displayWordPreview('Error loading template preview', false, template.placeholders);
        });
}

/**
 * Display Word document preview with highlighted placeholders
 */
function displayWordPreview(content, isHtml = false, placeholders = []) {
    const wordPreviewSection = document.getElementById('word-preview-section');
    const wordPreviewContent = document.getElementById('word-preview-content');
    
    // Show the preview section
    wordPreviewSection.style.display = 'block';
    
    let highlightedContent = content;
    
    if (isHtml) {
        // Content is HTML, highlight placeholders
        placeholders.forEach(placeholder => {
            const placeholderPattern = `{${placeholder}}`;
            const regex = new RegExp(placeholderPattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
            highlightedContent = highlightedContent.replace(regex, 
                `<span class="placeholder-highlight" data-placeholder="${placeholder}">{${placeholder}}</span>`
            );
        });
    } else {
        // Escape HTML content for safety (plain text)
        highlightedContent = highlightedContent
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
        
        // Highlight each placeholder
        placeholders.forEach(placeholder => {
            const placeholderPattern = `{${placeholder}}`;
            const regex = new RegExp(placeholderPattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
            highlightedContent = highlightedContent.replace(regex,
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
}

/**
 * Update Excel preview to include template selection
 */
function updateExcelPreviewWithTemplateSelection() {
    if (!excelData || excelData.length === 0) {
        return; // No Excel data to update
    }
    
    // Re-display Excel preview with template selection
    displayExcelPreviewWithTemplateSelection(excelData);
}

/**
 * Check if mapping section should be enabled
 */
function checkEnableMappingSection() {
    const hasExcelData = excelData && excelData.length > 0;
    const hasTemplates = Object.keys(wordTemplates).length > 0;
    
    if (hasExcelData && hasTemplates) {
        // Generate mapping fields for all templates
        generateMappingFields();
        document.getElementById('mapping-section').style.display = 'block';
        document.getElementById('generate-btn').disabled = false;
    } else {
        document.getElementById('mapping-section').style.display = 'none';
        document.getElementById('generate-btn').disabled = true;
    }
}

/**
 * Generate mapping fields for all templates
 */
function generateMappingFields() {
    const mappingContainer = document.getElementById('mapping-container');
    mappingContainer.innerHTML = '';
    
    // Get all unique placeholders from all templates
    const allPlaceholders = new Set();
    Object.values(wordTemplates).forEach(template => {
        template.placeholders.forEach(placeholder => {
            allPlaceholders.add(placeholder);
        });
    });
    
    if (allPlaceholders.size === 0) {
        mappingContainer.innerHTML = '<p>No placeholders found in any template.</p>';
        return;
    }
    
    // Create mapping interface
    const mappingHTML = Array.from(allPlaceholders).map(placeholder => {
        // Check which templates contain this placeholder
        const templatesWithPlaceholder = Object.keys(wordTemplates).filter(templateName => 
            wordTemplates[templateName].placeholders.includes(placeholder)
        );
        
        return `
            <div class="mapping-field">
                <label for="mapping-${placeholder}">
                    <strong>{${placeholder}}</strong>
                    <br><small>Used in: ${templatesWithPlaceholder.join(', ')}</small>
                </label>
                <select id="mapping-${placeholder}" data-placeholder="${placeholder}">
                    <option value="">-- Select Excel Column --</option>
                    ${excelHeaders.map(header => 
                        `<option value="${header}">${header}</option>`
                    ).join('')}
                </select>
            </div>
        `;
    }).join('');
    
    mappingContainer.innerHTML = mappingHTML;
}

/**
 * Display Excel preview with template selection for each row
 */
function displayExcelPreviewWithTemplateSelection(data) {
    const previewContainer = document.getElementById('excel-preview');
    previewContainer.innerHTML = '';
    
    if (data.length === 0) {
        previewContainer.innerHTML = '<p>No data found in Excel file</p>';
        return;
    }
    
    const templateNames = Object.keys(wordTemplates);
    if (templateNames.length === 0) {
        // No templates yet, show regular preview
        displayExcelPreview(data);
        return;
    }
    
    // Create table with template selection
    const table = document.createElement('table');
    table.className = 'excel-table';
    
    // Create header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    
    // Add template selection header
    const templateHeader = document.createElement('th');
    templateHeader.innerHTML = 'Word Template<br><small>Choose template for this row</small>';
    headerRow.appendChild(templateHeader);
    
    // Add data headers
    Object.keys(data[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // Create body
    const tbody = document.createElement('tbody');
    
    data.slice(0, 10).forEach((row, index) => { // Show first 10 rows
        const tr = document.createElement('tr');
        
        // Add template selection cell
        const templateCell = document.createElement('td');
        const templateSelect = document.createElement('select');
        templateSelect.className = 'template-selector';
        templateSelect.setAttribute('data-row-index', index);
        
        // Add default option
        const defaultOption = document.createElement('option');
        defaultOption.value = '';
        defaultOption.textContent = '-- Select Template --';
        templateSelect.appendChild(defaultOption);
        
        // Add template options
        templateNames.forEach(templateName => {
            const option = document.createElement('option');
            option.value = templateName;
            option.textContent = wordTemplates[templateName].fileName;
            templateSelect.appendChild(option);
        });
        
        // Set default selection if exists
        if (selectedTemplateForRows[index]) {
            templateSelect.value = selectedTemplateForRows[index];
        }
        
        // Add event listener for template selection
        templateSelect.addEventListener('change', function() {
            const rowIndex = parseInt(this.getAttribute('data-row-index'));
            const selectedTemplate = this.value;
            
            if (selectedTemplate) {
                selectedTemplateForRows[rowIndex] = selectedTemplate;
            } else {
                delete selectedTemplateForRows[rowIndex];
            }
            
            console.log('Template selection updated:', selectedTemplateForRows);
        });
        
        templateCell.appendChild(templateSelect);
        tr.appendChild(templateCell);
        
        // Add data cells
        Object.values(row).forEach(value => {
            const td = document.createElement('td');
            td.textContent = value;
            tr.appendChild(td);
        });
        
        tbody.appendChild(tr);
    });
    
    table.appendChild(tbody);
    previewContainer.appendChild(table);
    
    // Add info about showing limited rows
    if (data.length > 10) {
        const info = document.createElement('p');
        info.className = 'preview-info';
        info.innerHTML = `<small>Showing first 10 rows of ${data.length} total rows. Template selection will apply to all rows when generating documents.</small>`;
        previewContainer.appendChild(info);
    }
}

/**
 * Display Excel preview (fallback when no templates)
 */
function displayExcelPreview(data) {
    const previewContainer = document.getElementById('excel-preview');
    previewContainer.innerHTML = '';
    
    if (data.length === 0) {
        previewContainer.innerHTML = '<p>No data found in Excel file</p>';
        return;
    }
    
    // Create simple table
    const table = document.createElement('table');
    table.className = 'excel-table';
    
    // Create header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    
    Object.keys(data[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // Create body
    const tbody = document.createElement('tbody');
    
    data.slice(0, 10).forEach(row => { // Show first 10 rows
        const tr = document.createElement('tr');
        
        Object.values(row).forEach(value => {
            const td = document.createElement('td');
            td.textContent = value;
            tr.appendChild(td);
        });
        
        tbody.appendChild(tr);
    });
    
    table.appendChild(tbody);
    previewContainer.appendChild(table);
    
    // Add info about showing limited rows
    if (data.length > 10) {
        const info = document.createElement('p');
        info.className = 'preview-info';
        info.innerHTML = `<small>Showing first 10 rows of ${data.length} total rows.</small>`;
        previewContainer.appendChild(info);
    }
}

/**
 * Generate documents for all rows
 */
function generateDocument() {
    try {
        if (!excelData || excelData.length === 0) {
            alert('Please upload an Excel file first.');
            return;
        }
        
        if (Object.keys(wordTemplates).length === 0) {
            alert('Please upload at least one Word template.');
            return;
        }
        
        // Collect mapping data
        const mappingData = {};
        const mappingFields = document.querySelectorAll('#mapping-container select');
        
        mappingFields.forEach(field => {
            const placeholder = field.getAttribute('data-placeholder');
            const excelColumn = field.value;
            if (placeholder && excelColumn) {
                mappingData[placeholder] = excelColumn;
            }
        });
        
        if (Object.keys(mappingData).length === 0) {
            alert('Please map at least one placeholder to an Excel column.');
            return;
        }
        
        // Generate documents
        const generatedDocs = [];
        
        excelData.forEach((row, index) => {
            // Check if a template is selected for this row
            const selectedTemplate = selectedTemplateForRows[index];
            if (!selectedTemplate) {
                console.log(`No template selected for row ${index + 1}, skipping...`);
                return;
            }
            
            const template = wordTemplates[selectedTemplate];
            if (!template) {
                console.error(`Template ${selectedTemplate} not found for row ${index + 1}`);
                return;
            }
            
            try {
                // Load template
                const zip = new PizZip(template.content);
                const doc = new docxtemplater(zip, {
                    paragraphLoop: true,
                    linebreaks: true,
                });
                
                // Prepare data for this row
                const rowData = {};
                Object.keys(mappingData).forEach(placeholder => {
                    const columnName = mappingData[placeholder];
                    rowData[placeholder] = row[columnName] || '';
                });
                
                // Set the template variables
                doc.setData(rowData);
                
                try {
                    // Render the document
                    doc.render();
                    
                    // Get the generated document
                    const buf = doc.getZip().generate({
                        type: 'blob',
                        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    });
                    
                    generatedDocs.push({
                        blob: buf,
                        filename: `${template.fileName.replace('.docx', '')}_row_${index + 1}.docx`,
                        rowIndex: index + 1
                    });
                    
                } catch (renderError) {
                    console.error(`Error rendering template for row ${index + 1}:`, renderError);
                    alert(`Error rendering template for row ${index + 1}: ${renderError.message}`);
                }
                
            } catch (templateError) {
                console.error(`Error processing template for row ${index + 1}:`, templateError);
                alert(`Error processing template for row ${index + 1}: ${templateError.message}`);
            }
        });
        
        if (generatedDocs.length === 0) {
            alert('No documents were generated. Please check your template selections and mappings.');
            return;
        }
        
        // Download generated documents
        if (generatedDocs.length === 1) {
            // Single document - direct download
            saveAs(generatedDocs[0].blob, generatedDocs[0].filename);
        } else {
            // Multiple documents - create zip
            const zip = new JSZip();
            
            generatedDocs.forEach(doc => {
                zip.file(doc.filename, doc.blob);
            });
            
            zip.generateAsync({type: 'blob'}).then(function(zipBlob) {
                saveAs(zipBlob, 'generated_documents.zip');
            });
        }
        
        // Show success message
        alert(`Successfully generated ${generatedDocs.length} document(s)!`);
        
    } catch (error) {
        console.error('Error generating documents:', error);
        alert('Error generating documents: ' + error.message);
    }
}

/**
 * Show debug information
 */
function showDebugInfo() {
    const debugArea = document.getElementById('debug-area');
    const debugContent = document.getElementById('debug-content');
    
    const debugInfo = {
        excelData: excelData ? `${excelData.length} rows` : 'No data',
        excelHeaders: excelHeaders,
        wordTemplates: Object.keys(wordTemplates).map(name => ({
            name: name,
            fileName: wordTemplates[name].fileName,
            placeholders: wordTemplates[name].placeholders
        })),
        selectedTemplateForRows: selectedTemplateForRows
    };
    
    debugContent.textContent = JSON.stringify(debugInfo, null, 2);
    
    if (debugArea.style.display === 'none') {
        debugArea.style.display = 'block';
    } else {
        debugArea.style.display = 'none';
    }
}
