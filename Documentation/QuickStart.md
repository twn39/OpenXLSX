# Quick Start: Building a Complex Dashboard {#quickstart_tutorial}

[TOC]

## Introduction

This Quick Start guide demonstrates how to use OpenXLSX to generate a complex, real-world Excel dashboard. It covers almost all major features in the library:
- Worksheet management (renaming, hiding)
- Advanced cell styling (fonts, fills, borders, alignments, number formats)
- Cell merging and Rich Text
- Formulas
- Data validation (cross-sheet dropdown lists)
- Conditional formatting (data bars, color scales)
- Charts (Bar and Line)
- Comments
- Pane freezing and Row/Column sizing

By the end of this tutorial, you will have a comprehensive understanding of how to build sophisticated reports.

## 1. Document & Worksheet Setup

First, we create a new workbook, rename the default sheet to "Dashboard", and create a hidden "ReferenceData" sheet to hold our dropdown list categories.

```cpp
#include <OpenXLSX.hpp>
#include <iostream>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("ComplexDemo.xlsx", XLForceOverwrite);
    
    auto wks = doc.workbook().worksheet("Sheet1");
    wks.setName("Dashboard");

    // Add a hidden reference sheet for data validation lists
    doc.workbook().addWorksheet("ReferenceData");
    auto refWks = doc.workbook().worksheet("ReferenceData");
    refWks.cell("A1").value() = "Category A";
    refWks.cell("A2").value() = "Category B";
    refWks.cell("A3").value() = "Category C";
    refWks.setVisibility(XLSheetState::Hidden);
```

## 2. Defining Reusable Styles

OpenXLSX provides a high-level `XLStyle` facade that allows you to easily configure fonts, fills, borders, number formats, and alignments without worrying about low-level style indices.

```cpp
    XLStyle titleStyle;
    titleStyle.font.name = "Arial";
    titleStyle.font.size = 20;
    titleStyle.font.bold = true;
    titleStyle.font.color = XLColor("FFFFFF");
    titleStyle.fill.pattern = XLPatternSolid;
    titleStyle.fill.fgColor = XLColor("4F81BD");
    titleStyle.alignment.horizontal = XLAlignCenter;
    titleStyle.alignment.vertical = XLAlignCenter;
    
    XLStyle headerStyle;
    headerStyle.font.bold = true;
    headerStyle.font.color = XLColor("FFFFFF");
    headerStyle.fill.pattern = XLPatternSolid;
    headerStyle.fill.fgColor = XLColor("95B3D7");
    headerStyle.border.bottom.style = XLLineStyleThick;
    headerStyle.border.top.style = XLLineStyleThin;
    headerStyle.alignment.horizontal = XLAlignCenter;
    headerStyle.alignment.vertical = XLAlignCenter;

    XLStyle currencyStyle;
    currencyStyle.numberFormat = "$#,##0.00";
    currencyStyle.font.color = XLColor("006100");
    currencyStyle.border.bottom.style = XLLineStyleDotted;

    XLStyle dateStyle;
    dateStyle.numberFormat = "yyyy-mm-dd";
    dateStyle.alignment.horizontal = XLAlignCenter;
    dateStyle.border.bottom.style = XLLineStyleDotted;
```

## 3. Merging Cells & Rich Text

Use `mergeCells` to span content across multiple columns or rows. We also use the `XLRichText` API to create a multi-styled subtitle.

```cpp
    // Title
    wks.mergeCells("A1:E2");
    wks.cell("A1").value() = "Executive Sales Dashboard 2026";
    wks.cell("A1").setStyle(titleStyle);
    
    // Rich Text Subtitle
    XLRichText rtSubtitle;
    rtSubtitle.addRun("Confidential: ").setBold().setFontColor(XLColor("FFFF0000"));
    rtSubtitle.addRun("Internal Use Only. Data up to ");
    rtSubtitle.addRun("Q1 2026").setItalic().setUnderlineStyle(XLUnderlineSingle);
    
    wks.mergeCells("A3:E3");
    wks.cell("A3").value() = rtSubtitle;
```

## 4. Populating Data & Applying Styles

Now, let's add headers and some mock data, applying the styles we defined earlier using `.setStyle()`. 

```cpp
    wks.cell("A4").value() = "Date";
    wks.cell("B4").value() = "Product Name";
    wks.cell("C4").value() = "Category";
    wks.cell("D4").value() = "Units Sold";
    wks.cell("E4").value() = "Revenue";
    
    wks.cell("A4").setStyle(headerStyle);
    wks.cell("B4").setStyle(headerStyle);
    wks.cell("C4").setStyle(headerStyle);
    wks.cell("D4").setStyle(headerStyle);
    wks.cell("E4").setStyle(headerStyle);
    
    // Example data row
    wks.cell("A5").value() = 45383.0; // Excel date representation
    wks.cell("A5").setStyle(dateStyle);
    
    wks.cell("B5").value() = "Widget X (Enterprise Edition)";
    wks.cell("B5").setStyle(dateStyle); // Reusing border settings from dateStyle
    
    wks.cell("C5").value() = "Category A";
    
    wks.cell("D5").value() = 1500;
    
    wks.cell("E5").value() = 45000.00;
    wks.cell("E5").setStyle(currencyStyle);
    
    // ... [Add more rows similarly] ...
```

## 5. Formulas

You can assign Excel formulas directly to a cell using the `.formula()` property.

```cpp
    wks.cell("C9").value() = "Total:";
    
    wks.cell("D9").formula() = "SUM(D5:D8)";
    wks.cell("E9").formula() = "SUM(E5:E8)";
```

## 6. Data Validation & Dropdown Lists

To ensure data integrity, we restrict the "Category" column to only allow values from our hidden `ReferenceData` sheet.

```cpp
    wks.dataValidations().add("C5:C8")
        .requireList("ReferenceData!$A$1:$A$3")
        .setErrorAlert("Invalid Category", "Please select a valid category from the dropdown list.", XLDataValidationErrorStyle::Stop)
        .setPromptMessage("Category Selection", "Choose a category from ReferenceData.");
```

## 7. Comments

Cell comments are useful for adding notes to specific data points.

```cpp
    wks.comments().addAuthor("System Generated");
    wks.comments().set("E9", "These totals are automatically calculated using the SUM formula.", 0);
    wks.comments().set("B5", "Top selling product this quarter.", 0);
```

## 8. Conditional Formatting

We can use Data Bars to visualize quantities and Color Scales to represent revenue performance directly in the cells.

```cpp
    // Data Bars for Units Sold
    auto dbRule = XLDataBarRule(XLColor("FF4F81BD"), true);
    wks.addConditionalFormatting("D5:D8", dbRule);
    
    // Color Scale for Revenue (Red to Green)
    auto csRule = XLColorScaleRule(XLColor("FFF8696B"), XLColor("FF63BE7B"));
    wks.addConditionalFormatting("E5:E8", csRule);
```

## 9. Charts

OpenXLSX supports embedding several chart types. We'll add a Bar chart for units and a Line chart for revenue.

```cpp
    // Bar Chart
    auto barChart = wks.addChart(XLChartType::Bar, "Units Sold by Product", 11, 1, 400, 250);
    barChart.addSeries("Dashboard!$D$5:$D$8", "Units", "Dashboard!$B$5:$B$8");

    // Line Chart
    auto lineChart = wks.addChart(XLChartType::Line, "Revenue Trend", 11, 4, 400, 250);
    lineChart.addSeries("Dashboard!$E$5:$E$8", "Revenue", "Dashboard!$A$5:$A$8");
```

## 10. Layout: Freezing Panes & Resizing

Finally, adjust the row heights, column widths, and freeze the top 4 rows so headers remain visible when scrolling.

```cpp
    wks.column(1).setWidth(15);
    wks.column(2).setWidth(35);
    wks.column(5).setWidth(20);
    
    wks.row(1).setHeight(30); 
    wks.row(4).setHeight(25);

    // Freeze panes starting from A5 (keeps rows 1-4 visible)
    wks.freezePanes("A5");

    doc.save();
    doc.close();
    
    return 0;
}
```

Compile and run this code, and you will generate a highly professional Excel Dashboard complete with rich functionality!
