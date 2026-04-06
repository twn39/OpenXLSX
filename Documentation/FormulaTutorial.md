# Quick Start: Formulas & Calculations {#formula_tutorial}

[TOC]

## Introduction

OpenXLSX allows you to easily inject native Microsoft Excel formulas and equations into your spreadsheets. When users open the generated `.xlsx` files, Excel will automatically evaluate these formulas.

Furthermore, OpenXLSX comes with its own **built-in Formula Engine** that is capable of evaluating hundreds of Excel functions *in C++* without ever needing to open the file in Excel.

This tutorial covers:
- Writing mathematical formulas and linking them to cells.
- Using aggregation and range-based functions.
- Performing text manipulation, logical statements, and lookups.
- Referencing data between different cells.

## 1. Writing Formulas to Cells

Writing a formula to a cell is done via the `.formula()` property instead of `.value()`. You do not need to prepend an `=` sign when assigning the string, though OpenXLSX will handle it correctly if you do.

```cpp
#include <OpenXLSX.hpp>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("FormulaDemo.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Basic arithmetic
    wks.cell("A1").formula() = "10 + 20";
    wks.cell("A2").formula() = "(100 - 20) / 4 * 2";
    
    // Mathematical functions
    wks.cell("A3").formula() = "POWER(2, 8)";
    wks.cell("A4").formula() = "SQRT(144)";
    wks.cell("A5").formula() = "MOD(25, 7)";
    wks.cell("A6").formula() = "ABS(-123.45)";
```

## 2. Cell References and Aggregation

You can build formulas that refer to other cells. Excel will update the formula result whenever the referenced cells change.

```cpp
    wks.cell("C1").value() = 100;
    wks.cell("C2").value() = 200;
    wks.cell("C3").value() = 300;
    wks.cell("C4").value() = 400;

    // Simple references
    wks.cell("D1").formula() = "C1 * 2";

    // Range Aggregation
    wks.cell("D2").formula() = "SUM(C1:C4)";
    wks.cell("D3").formula() = "AVERAGE(C1:C4)";
    wks.cell("D4").formula() = "MAX(C1:C4)";
```

## 3. Logical and Text Functions

OpenXLSX supports writing advanced logic and text formatting formulas. 

```cpp
    // Logical
    wks.cell("E1").formula() = "IF(D2 > 500, \"Target Met\", \"Target Missed\")";
    wks.cell("E2").formula() = "SWITCH(2, 1, \"One\", 2, \"Two\", \"Other\")";
    wks.cell("E3").formula() = "IFS(C1=100, \"A\", C2=200, \"B\", C3=300, \"C\")";

    // Text processing
    wks.cell("F1").value() = "OpenXLSX";
    wks.cell("F2").formula() = "LEFT(F1, 4)";               // Returns "Open"
    wks.cell("F3").formula() = "MID(F1, 5, 4)";             // Returns "XLSX"
    wks.cell("F4").formula() = "UPPER(\"hello world\")";
    wks.cell("F5").formula() = "SUBSTITUTE(\"A-B-C\", \"-\", \"_\")";
```

## 4. Date, Financial, and Lookup Functions

You can also leverage Excel's powerful built-in analytical capabilities such as lookups and financial calculators.

```cpp
    // Date & Time
    wks.cell("G1").formula() = "DATE(2026, 4, 1)";
    wks.cell("G2").formula() = "WORKDAY(DATE(2026, 4, 1), 5)";

    // Lookups (Assuming we search for the value 200 in the C1:C4 range)
    wks.cell("H1").formula() = "MATCH(200, C1:C4, 0)";
    wks.cell("H2").formula() = "VLOOKUP(200, C1:D4, 2, FALSE)";
    wks.cell("H3").formula() = "XLOOKUP(200, C1:C4, D1:D4, \"Not Found\")";

    // Financial calculations
    // Calculate monthly mortgage payments: PMT(rate, nper, pv)
    wks.cell("I1").formula() = "PMT(0.05/12, 360, -100000)";
    
    // Calculate future value of an investment
    wks.cell("I2").formula() = "FV(0.05, 10, -100)";
```

## 5. Evaluating Formulas in C++ 

*(Advanced Feature)*
If you need to know the result of a formula *before* the document is opened in Excel, OpenXLSX contains a built-in AST math engine to evaluate strings.

```cpp
    // The built-in engine can calculate expressions
    XLFormulaEngine eng;
    auto result = eng.evaluate("=10 + 20 * 2");
    
    if (result.type() == XLValueType::Float) {
        std::cout << "The calculated answer is: " << result.get<double>() << std::endl; // Outputs 50
    }
    
    // It also supports string and boolean return types
    auto textResult = eng.evaluate("=LEFT(\"Excel\", 3)");
    std::cout << "The text answer is: " << textResult.get<std::string>() << std::endl; // Outputs "Exc"
```

Saving the workbook will commit all `.formula()` properties into the XML so Excel can execute them visually.

```cpp
    doc.save();
    doc.close();
    
    return 0;
}
```
