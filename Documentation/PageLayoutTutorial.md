# Quick Start: Page Layout & Printing {#pagelayout_tutorial}

[TOC]

## Introduction

OpenXLSX provides comprehensive support for configuring how a worksheet should be printed. You can define page orientation, paper size, margins, headers and footers, print areas, repeated titles, and manual page breaks.

This tutorial covers:
- Configuring basic Page Setup (Orientation, Paper Size, Scaling).
- Adjusting Page Margins.
- Setting up Print Options (Gridlines, Headings, Centering).
- Creating custom Headers and Footers with dynamic Excel codes.
- Defining Print Areas and repeating title rows/columns.
- Inserting manual Page Breaks.

## 1. Page Setup Configuration

The `XLPageSetup` object handles the core physical properties of the printed page. 

```cpp
#include <OpenXLSX.hpp>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("PageLayoutDemo.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    auto ps = wks.pageSetup();
    
    // Set to Landscape orientation
    ps.setOrientation(XLPageOrientation::Landscape);
    
    // Set paper size to A4 (Internal Excel ID for A4 is 9)
    ps.setPaperSize(9); 
    
    // Force the printed output to fit within 1 page wide, but as many pages tall as needed
    ps.setFitToWidth(1);
    ps.setFitToHeight(0); // 0 indicates automatic scaling for height

    // Define the order in which pages are numbered and printed
    ps.setPageOrder("downThenOver"); 
    
    // Start page numbering from 1
    ps.setFirstPageNumber(1);
    ps.setUseFirstPageNumber(true);
```

## 2. Page Margins

The `XLPageMargins` object allows you to fine-tune the empty space around your content. Values are provided in inches.

```cpp
    auto pm = wks.pageMargins();
    
    pm.setTop(1.5);    // 1.5 inches from the top edge
    pm.setBottom(1.0); // 1.0 inch from the bottom edge
    pm.setLeft(0.75);
    pm.setRight(0.75);
    
    // Space allocated specifically for the header and footer text
    pm.setHeader(0.5);
    pm.setFooter(0.5);
```

## 3. Print Options

The `XLPrintOptions` object controls visual aids that appear on the printed page, such as gridlines and headings.

```cpp
    auto po = wks.printOptions();
    
    po.setGridLines(true);          // Print the cell gridlines
    po.setHeadings(true);           // Print the row/column headings (A, B, C... 1, 2, 3...)
    po.setHorizontalCentered(true); // Center the content horizontally on the paper
    po.setVerticalCentered(false);  // Align content to the top vertically
```

## 4. Headers and Footers

Headers and footers can be heavily customized using the `XLHeaderFooter` object. Excel uses special formatting codes (e.g., `&L`, `&C`, `&R`) to define alignment, and placeholders (e.g., `&P`, `&N`) for dynamic data like page numbers.

```cpp
    auto hf = wks.headerFooter();
    
    // Enable a unique header/footer for the very first printed page
    hf.setDifferentFirst(true);
    
    // First Page Header: Centered (&C), Arial font, Bold, Size 24
    hf.setFirstHeader("&C&\"Arial,Bold\"&24 CONFIDENTIAL REPORT");
    
    // First Page Footer: Centered (&C), Size 12
    hf.setFirstFooter("&C&12 Restricted Distribution");

    // Standard Header (used for page 2 onwards):
    // Left: Tab Name (&A) | Center: Static Text | Right: Date (&D) and Time (&T)
    hf.setOddHeader("&L&A &CHidden Corp &R&D &T"); 
    
    // Standard Footer: Centered Page Number (&P) of Total Pages (&N)
    hf.setOddFooter("&CPage &P of &N");
```

## 5. Print Titles and Print Area

You often want specific rows (like a table header) to repeat at the top of every printed page. You may also want to restrict printing to a specific range of cells.

```cpp
    // Repeat Row 1 at the top of every printed page
    wks.setPrintTitleRows(1, 1);
    
    // Repeat Columns A and B at the left of every printed page
    // wks.setPrintTitleCols(1, 2);
    
    // Restrict the printable area to columns A through F, rows 1 through 50
    wks.setPrintArea("A1:F50");
```

## 6. Manual Page Breaks

If you need absolute control over where a page ends, you can insert manual row and column breaks.

```cpp
    // Force a horizontal page break immediately after row 25
    wks.insertRowBreak(25);
    
    // Force a vertical page break immediately after column 3 (Column C)
    wks.insertColBreak(3);

    doc.save();
    doc.close();
    
    return 0;
}
```

By combining these layout properties, you can ensure that your generated Excel reports look exactly as intended when exported to PDF or sent to a physical printer.
