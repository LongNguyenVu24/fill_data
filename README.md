# Excel to Word Document Filler

A web app to fill Word document templates with data from Excel files, all in your browser.

## Features

- Upload Excel (.xlsx) and Word (.docx, .doc) files
- **NEW**: Automatic conversion of .doc files to .docx format
- Detect and map Excel columns to Word placeholders (`{placeholder_name}`)
- Preview Excel data and Word template with highlighted placeholders
- Auto-match columns and placeholders by name
- Click placeholders in preview to jump to mapping fields
- Generate one Word document per Excel row
- Download individual documents or all as a ZIP file
- Progress bar for batch generation
- Debug panel for troubleshooting

## Quick Start

1. **Prepare Files**
   - Excel: Add your data in columns
   - Word: Use placeholders like `{name}` in your template (supports both .docx and .doc)

2. **Upload Files**
   - Upload Excel and Word files in the app
   - .doc files will be automatically converted to .docx format

3. **Preview & Map**
   - Review Excel data and Word template
   - Map columns to placeholders (auto-matching available)

4. **Generate & Download**
   - Click "Generate Documents for All Rows"
   - Download results individually or as a ZIP

## Template Example

Use curly braces for placeholders in your Word document:

```
Dear {name},

Your account balance is {balance}.

Sincerely,
{company_name}
```

## File Format Support

### Word Documents
- **.docx files**: Full support (recommended)
- **.doc files**: Supported with automatic conversion to .docx format
  - Conversion process is handled automatically in the browser
  - Progress indicator shows conversion status
  - Some legacy .doc files may have limitations

### Excel Files
- **.xlsx files**: Full support
- **.xls files**: Full support

## Troubleshooting

- **Debug Info:** Click "Show Debug Info" for details on file status, placeholders, and mappings.
- **Common Issues:**
  - No placeholders found: Use `{placeholder_name}` format
  - Generation errors: Map at least one column
  - Invalid file: Use supported formats (.xlsx/.xls for Excel, .docx/.doc for Word)
  - .doc conversion issues: Try opening the file in Word and saving as .docx manually
- **File Conversion:** .doc files are automatically converted to .docx format for processing.

## Technology

- HTML5, CSS3, JavaScript
- Libraries:
  - SheetJS (xlsx) for Excel parsing
  - docxtemplater for Word templating
  - PizZip for ZIP/docx handling
  - mammoth.js for Word-to-HTML preview
  - FileSaver.js for downloads
  - Automatic DOC to DOCX conversion capability

## Running

Open `index.html` in your browser. No server needed; all processing is local.

## Notes

- All processing is client-side; files are not uploaded anywhere.
- Large Excel files may take longer to process.
- .doc files undergo automatic conversion which may take a few moments.
- For best results with legacy .doc files, consider manually converting to .docx first.
