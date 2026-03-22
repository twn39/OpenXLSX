# OpenXLSX Streaming Reader API Implementation Plan

## Architecture: Sliding Window & Micro-DOM
We will bypass the full PugiXML DOM load for the `sheetData` node. Instead, we directly read the raw XML from the ZIP archive in chunks, extract `<row>...</row>` fragments using string searching, and parse only these tiny fragments using PugiXML.

### Phase 1: ZIP Engine Enhancement
- Modify `IZipArchive` and `XLZipArchive` to support `openEntryStream`, `readEntryStream`, and `closeEntryStream`.
- Wraps `zip_fopen`, `zip_fread`, and `zip_fclose` to avoid loading the entire file into memory at once.

### Phase 2: Core StreamReader Class
- Create `XLStreamReader` with a streaming buffer (e.g., 64KB chunks).
- Implement `fetchMoreData()` to pull data from the ZIP stream.
- Implement `extractNextRowXml()` to identify `<row>...</row>` boundaries.
- Implement `nextRow()` which loads the extracted XML into a temporary `pugi::xml_document`, resolves cell references (`r` attribute) to align the `std::vector<XLCellValue>`, and decodes types (e.g., resolving `t="s"` against the `XLSharedStrings` table).

### Phase 3: Worksheet State Bridge
- Add `XLWorksheet::streamReader() const` to instantiate the reader for a specific worksheet.

### Phase 4: Testing & Verification
- Test against large sheets.
- Test edge cases (skipped cells, inline strings, shared strings, booleans, empty rows).
