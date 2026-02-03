# DOCX to PDF Converter - Claude Skill

A Claude skill for batch converting Word documents (.docx) to PDF format using LibreOffice.

## What is this?

This is a custom skill for Claude that allows you to convert multiple Word documents to PDF in one go, without having to manually open each file and "Save As" → PDF.

## Features

- ✅ **Batch processing** - Convert multiple files at once
- ✅ **Preserves filenames** - Output PDFs keep the same name as input files
- ✅ **Error handling** - Continues processing even if one file fails
- ✅ **Clear feedback** - Shows summary of successful/failed conversions
- ✅ **High quality** - Uses LibreOffice conversion engine (~90-95% identical to Word)

## Installation

1. **Copy the skill folder** to your Claude skills directory:
   ```
   /mnt/skills/user/docx-to-pdf/
   ```

2. **Ensure LibreOffice is installed** on your system:
   ```bash
   # For Ubuntu/Debian
   sudo apt-get install libreoffice
   
   # For macOS
   brew install libreoffice
   ```

3. **File structure** should look like:
   ```
   /mnt/skills/user/docx-to-pdf/
   ├── SKILL.md
   └── scripts/
       └── convert_to_pdf.py
   ```

## Usage

### In Claude

Simply upload one or more Word documents and ask Claude:

- "Convert these to PDF"
- "Export these Word files to PDF"
- "Turn these DOCX files into PDFs"

Claude will automatically:
1. Detect the uploaded DOCX files
2. Run the conversion script
3. Provide you with download links for all the PDFs

### Standalone (Command Line)

You can also run the script directly:

```bash
python3 /mnt/skills/user/docx-to-pdf/scripts/convert_to_pdf.py file1.docx file2.docx

# Or convert all DOCX files in a directory
python3 /mnt/skills/user/docx-to-pdf/scripts/convert_to_pdf.py /path/to/files/*.docx
```

## Example Output

```
Starting conversion of 3 file(s)...

Converting: contract.docx... ✓ Success
Converting: report.docx... ✓ Success
Converting: proposal.docx... ✗ Failed

============================================================
CONVERSION SUMMARY
============================================================
Total files: 3
Successful: 2
Failed: 1

✓ Successfully converted:
  • contract.docx → contract.pdf
  • report.docx → report.pdf

✗ Failed conversions:
  • proposal.docx
    Reason: File corrupted

All PDF files are saved in: /mnt/user-data/outputs/
```

## Technical Details

- **Conversion engine**: LibreOffice (headless mode)
- **Supported input**: `.docx` files only
- **Output location**: `/mnt/user-data/outputs/`
- **Timeout**: 60 seconds per file
- **Quality**: ~90-95% identical to Word's native PDF export

### What converts well:
- Standard text and paragraphs
- Tables and borders
- Images and basic shapes
- Common fonts and styles
- Headers and footers

### Potential limitations:
- Complex embedded objects
- Rare font combinations
- Advanced Word-specific features
- Very large files (>100MB) may timeout

## Contributing

Found a bug or want to improve the skill? Feel free to:
- Open an issue
- Submit a pull request
- Suggest new features

## License

MIT License - feel free to use and modify as needed.

## Author

Eliyahu Biton - Lawyer, AI & Legal Tech Educator

---

**Note**: This skill uses LibreOffice for conversion. While the quality is very high, there may be minor differences compared to Microsoft Word's native PDF export. For most legal documents, contracts, and standard business documents, the output is virtually identical.
