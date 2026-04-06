# Quick Start: Pivot Tables & Slicers {#pivottable_tutorial}

[TOC]

## Introduction

OpenXLSX provides a powerful engine for generating dynamic Pivot Tables and interactive Slicers. Unlike static tables, these pivot tables are configured to automatically calculate and refresh their layout when the user opens the Excel file.

This tutorial covers:
- Preparing the source data for a pivot table.
- Defining row, column, and data fields with specific aggregations (Sum, Average, etc.).
- Configuring advanced layout options (Compact vs Tabular, Grand Totals).
- Applying custom Pivot Table styles.
- Adding an interactive Slicer to filter the pivot table visually.

## 1. Preparing the Source Data

A pivot table requires a contiguous range of data with column headers in the first row. We start by creating a worksheet and filling it with our raw dataset.

```cpp
#include <OpenXLSX.hpp>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("PivotDemo.xlsx", XLForceOverwrite);
    
    // Create the raw data sheet
    auto rawWks = doc.workbook().worksheet("Sheet1");
    rawWks.setName("RawData");

    // Add Headers
    rawWks.cell("A1").value() = "Date";
    rawWks.cell("B1").value() = "Region";
    rawWks.cell("C1").value() = "SalesRep";
    rawWks.cell("D1").value() = "Product";
    rawWks.cell("E1").value() = "Revenue";

    // Add some sample data rows...
    rawWks.cell("A2").value() = "2026-Q1";
    rawWks.cell("B2").value() = "North";
    rawWks.cell("C2").value() = "Alice";
    rawWks.cell("D2").value() = "Widget A";
    rawWks.cell("E2").value() = 15000.0;
    
    // ... [Assume more data is added up to row 13] ...
```

## 2. Creating a Standard Pivot Table

To create a pivot table, use the `XLPivotTableOptions` struct. You must provide a name, the source range, and the target cell where the pivot table will be rendered.

```cpp
    doc.workbook().addWorksheet("Summary");
    auto pivotWks = doc.workbook().worksheet("Summary");

    XLPivotTableOptions options;
    options.name = "PivotSummary";
    options.sourceRange = "RawData!A1:E13"; // Include headers and data
    options.targetCell = "B3";            // Top-left cell of the pivot table
    
    // 1. Page/Report Filter
    options.filters.push_back({"Region", XLPivotSubtotal::Sum, ""});

    // 2. Row Fields (Hierarchical grouping)
    options.rows.push_back({"SalesRep", XLPivotSubtotal::Sum, ""});

    // 3. Column Fields
    options.columns.push_back({"Date", XLPivotSubtotal::Sum, ""});

    // 4. Data Fields (Metrics to calculate)
    XLPivotField revField;
    revField.name = "Revenue";
    revField.subtotal = XLPivotSubtotal::Sum;    // Default aggregation
    revField.customName = "Total Revenue";       // Custom label in the table
    revField.numFmtId = 4;                       // Internal Excel ID for '#,##0.00' format
    
    options.data.push_back(revField);

    // Apply the options
    auto pt = pivotWks.addPivotTable(options);
    
    // VERY IMPORTANT: Tell Excel to calculate the pivot table upon opening the file
    pt.setRefreshOnLoad(true); 
```

## 3. Advanced Layouts & Formatting

You can highly customize the visual appearance and layout of your pivot table by toggling flags within `XLPivotTableOptions`.

```cpp
    // Visual Style 
    // Excel provides many built-in styles like PivotStyleMedium14, PivotStyleDark3, etc.
    options.pivotTableStyleName = "PivotStyleMedium14";
    options.showRowStripes = true;  // Enable banded rows
    options.showColStripes = false;

    // Layout configuration
    options.compactData = false;    // Switch from Compact form to Tabular form
    options.rowGrandTotals = false; // Hide grand totals for rows
    options.colGrandTotals = true;  // Keep grand totals for columns
    
    // Multiple Data Fields Formatting
    // If you have multiple data fields, you can choose to stack them in rows instead of columns:
    options.dataOnRows = true; 
```

## 4. Adding Slicers

A Slicer is a visual filtering control that connects to a Pivot Table. Adding a slicer in OpenXLSX is extremely straightforward once the Pivot Table is created.

```cpp
    // Configure the Slicer
    XLSlicerOptions sOpts;
    sOpts.name = "Region";          // The exact name of the source data column
    sOpts.caption = "Select Region"; // The display title of the slicer window

    // Add the slicer to the worksheet, anchoring it at cell "E5"
    pivotWks.addPivotSlicer("E5", pt, "Region", sOpts);
```

## 5. Finalizing the Document

Save and close the document as usual.

```cpp
    doc.save();
    doc.close();
    
    return 0;
}
```

When you open the generated `PivotDemo.xlsx` file in Microsoft Excel, the pivot engine will automatically read the `RawData` sheet, calculate the aggregations, and render the fully functional table alongside the interactive slicer!
