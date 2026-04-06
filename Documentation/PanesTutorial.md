# Quick Start: Freezing & Splitting Panes {#panes_tutorial}

[TOC]

## Introduction

When working with large datasets, it is often necessary to keep specific rows (like headers) or columns (like identifiers) visible while the user scrolls through the rest of the worksheet. OpenXLSX provides an elegant and straightforward API to handle this layout feature.

This tutorial covers:
- Freezing the top row or the first column.
- Freezing both rows and columns simultaneously.
- Splitting the window into independent scrollable panes.
- Removing all pane configurations from a worksheet.

## 1. Freezing the Top Row

To freeze the top row (Row 1), you specify the cell *immediately below* the row you want to freeze. Thus, you tell OpenXLSX to freeze panes starting at `A2`.

```cpp
#include <OpenXLSX.hpp>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("PanesDemo.xlsx", XLForceOverwrite);
    
    auto wks = doc.workbook().worksheet("Sheet1");
    
    // Add headers in row 1
    wks.cell("A1").value() = "Date";
    wks.cell("B1").value() = "Sales";
    
    // Add data in subsequent rows...
    wks.cell("A2").value() = "2026-01-01";
    wks.cell("B2").value() = 1500;
    
    // Freeze everything above Row 2 (which freezes Row 1)
    wks.freezePanes("A2");
```

## 2. Freezing the First Column

Similarly, to freeze a column, you specify the cell *immediately to the right* of the column you wish to freeze. To freeze Column A, you tell the worksheet to freeze starting at `B1`.

```cpp
    doc.workbook().addWorksheet("FreezeColumn");
    auto wksCol = doc.workbook().worksheet("FreezeColumn");
    
    wksCol.cell("A1").value() = "ID";
    wksCol.cell("B1").value() = "Data 1";
    wksCol.cell("C1").value() = "Data 2";
    
    // Freeze everything to the left of Column B (which freezes Column A)
    wksCol.freezePanes("B1");
```

## 3. Freezing Both Rows and Columns

You can freeze both rows and columns simultaneously by selecting a cell that is not in row 1 and not in column A. 

For example, to freeze the first two rows (Row 1 and 2) AND the first two columns (Column A and B), you choose cell `C3`.

```cpp
    doc.workbook().addWorksheet("FreezeBoth");
    auto wksBoth = doc.workbook().worksheet("FreezeBoth");
    
    // Freeze Row 1, Row 2, Column A, and Column B
    wksBoth.freezePanes("C3");
```

## 4. Splitting Panes

While "Freezing" locks the headers in place, "Splitting" divides the Excel window into up to four independent, scrollable viewports. 

To split panes, you use `splitPanes()`, providing the X and Y distance (in twips, where 1/20th of a point = 1 twip). You can optionally specify the cell that should appear at the top-left of the bottom-right pane, and which pane should be the active (selected) one.

```cpp
    doc.workbook().addWorksheet("SplitScreen");
    auto wksSplit = doc.workbook().worksheet("SplitScreen");
    
    // Split the window horizontally and vertically.
    // Set the active focus to the BottomRight pane.
    wksSplit.splitPanes(3000.0, 3000.0, "E10", XLPane::BottomRight);
```

## 5. Clearing Panes

If you need to programmatically remove any freezing or splitting from a worksheet, you can use the `clearPanes()` method.

```cpp
    doc.workbook().addWorksheet("ClearTest");
    auto wksClear = doc.workbook().worksheet("ClearTest");
    
    // Apply a freeze
    wksClear.freezePanes("D5");
    
    // Remove the freeze completely
    wksClear.clearPanes();
```

## Summary

Finally, save and close your document.

```cpp
    doc.save();
    doc.close();
    
    return 0;
}
```

The resulting `PanesDemo.xlsx` file will correctly lock your headers or split your data windows exactly as defined, making large datasets much easier for your end-users to navigate!
