# OpenXLSX XLStreamWriter Tutorial & Best Practices

The `XLStreamWriter` is a highly optimized, zero-allocation module designed for generating massive Excel worksheets (millions of cells) with virtually zero memory footprint. It bypasses the standard `pugixml` DOM engine and writes XML data directly to a temporary file stream.

## Core Concepts & Limitations

Because `XLStreamWriter` streams data directly to disk:
1. **No In-Memory DOM**: Cells written via `XLStreamWriter` **do not exist** in the `XLDocument`'s memory. You cannot read them back using `wks.cell(...)` during the same session.
2. **Sequential Writing Only**: You must write data row by row, from top to bottom. You cannot go back to modify a previous row.
3. **Exclusive Lock**: While a stream writer is active, you should not modify the worksheet using the standard DOM methods (`wks.cell(...)`, `wks.row(...)`).

## ⚠️ Critical Lifecycle Trap: Destruction Before `save()`

The `XLStreamWriter` keeps the XML stream open and buffers data. It relies on its **destructor** (`~XLStreamWriter()`) to flush the final XML footers (like `</sheetData>`) and register the temporary file with the main `XLDocument` archive.

**If you call `doc.save()` before the `XLStreamWriter` is destroyed, the resulting Excel file will be CORRUPTED.**

### ❌ Incorrect Usage (Will Corrupt File)

```cpp
XLDocument doc;
auto wks = doc.workbook().worksheet("Sheet1");
auto stream = wks.streamWriter();

std::vector<XLCellValue> row = {1, 2, 3};
stream.appendRow(row);

doc.save(); // ERROR: stream is still alive! The XML footers haven't been written.
```

### ✅ Correct Usage (Using Scope Blocks)

Always wrap your streaming logic in a local scope `{ ... }` so the stream writer is destroyed and flushed before you save.

```cpp
XLDocument doc;
auto wks = doc.workbook().worksheet("Sheet1");

{ // Open scope for stream writer
    auto stream = wks.streamWriter();
    
    std::vector<XLCellValue> row = {1, 2, 3};
    stream.appendRow(row);
} // stream is destroyed here. XML footers are written safely to disk.

doc.save(); // SUCCESS: The archive packs the fully completed stream file.
```

## ⚠️ Streaming Data for Pivot Tables

When creating an `XLPivotTable` based on a data source range (`RawData!A1:D1000`), the `addPivotTable()` method needs to read the **Header Row (Row 1)** of the source data to define the internal pivot cache fields.

Since `XLStreamWriter` bypasses the DOM, if you write your header row using the stream writer, `addPivotTable()` **will not find the column names**, resulting in generic names like `"Field1", "Field2"`, which will break your pivot table rendering in Excel.

### ✅ Correct Usage: Hybrid DOM + Stream Approach

Write your Header Row using the standard DOM methods so they are cached in memory. Then, switch to `XLStreamWriter` for the massive data rows.

```cpp
auto wksData = doc.workbook().worksheet("RawData");

// 1. Write headers using DOM (Loaded in memory, accessible by Pivot Cache generator)
wksData.cell("A1").value() = "Region";
wksData.cell("B1").value() = "Sales";

// 2. Write massive data using Stream Writer (Zero memory footprint)
{
    auto stream = wksData.streamWriter();
    for (int i = 0; i < 1000000; ++i) {
        std::vector<XLCellValue> row = {"North", 500.0};
        stream.appendRow(row);
    }
}

// 3. Create Pivot Table safely
auto pivotWks = doc.workbook().worksheet("PivotDashboard");
XLPivotTableOptions options("MyPivot", "RawData!A1:B1000001", "A3");
options.addRowField("Region").addDataField("Sales");

auto pt = pivotWks.addPivotTable(options);
pt.setRefreshOnLoad(true); // Tell Excel to parse the 1,000,000 rows on open

doc.save();
```

## Supported Types & Zero-Allocation Features

Thanks to recent optimizations, `XLStreamWriter` uses `std::to_chars` and stack buffers internally. It performs zero heap allocations during the writing of standard data types.

You can mix and match types (Integers, Floats, Booleans, Strings, and even `XLRichText`) within the same row by using `std::vector<XLCellValue>`.
