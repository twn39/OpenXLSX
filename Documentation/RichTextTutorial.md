# Rich Text Quick Start {#richtext_tutorial}

[TOC]

## Introduction

OpenXLSX provides a powerful fluid API for applying rich text formatting to cells. This allows you to combine multiple text styles—such as different fonts, colors, sizes, and effects—within a single cell.

To use rich text, create an `XLRichText` object, append runs of text using `addRun()`, and configure the styling for each run via chainable method calls. Finally, assign the `XLRichText` object directly to a cell's value.

## 1. Basic Formatting

You can easily apply basic text styles like bold, italic, strikethrough, and colors. Colors are defined using `XLColor` with hex codes.

```cpp
#include <OpenXLSX.hpp>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("RichTextDemo.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    XLRichText rt1;
    rt1.addRun("Red & Bold: ").setFontColor(XLColor("FF0000")).setBold();
    rt1.addRun("Blue & Italic: ").setFontColor(XLColor("0000FF")).setItalic();
    rt1.addRun("Strikethrough: ").setStrikethrough(true);
    rt1.addRun("Underline: ").setUnderlineStyle(XLUnderlineSingle);

    wks.cell("A1").value() = rt1;

    doc.save();
    doc.close();
    return 0;
}
```

## 2. Fonts and Sizes

You can change the font family and size on a per-run basis.

```cpp
XLRichText rt2;
rt2.addRun("Arial 10, ").setFontName("Arial").setFontSize(10);
rt2.addRun("Courier New 14, ").setFontName("Courier New").setFontSize(14).setFontColor(XLColor("008000"));
rt2.addRun("Times New Roman 18, ").setFontName("Times New Roman").setFontSize(18).setBold();
rt2.addRun("Comic Sans MS 24").setFontName("Comic Sans MS").setFontSize(24).setItalic();

wks.cell("A2").value() = rt2;
```

## 3. Subscript and Superscript

Subscript and superscript alignments are applied using `setVertAlign()` with the `XLSubscript` or `XLSuperscript` enumerations.

```cpp
XLRichText rt3;
rt3.addRun("Water is H");
rt3.addRun("2").setVertAlign(XLSubscript);
rt3.addRun("O. And E = mc");
rt3.addRun("2").setVertAlign(XLSuperscript);

wks.cell("A3").value() = rt3;
```

## 4. Various Underline Types

OpenXLSX supports several underline styles for rich text, including standard and accounting variants.

```cpp
XLRichText rt4;
rt4.addRun("Single, ").setUnderlineStyle(XLUnderlineSingle);
rt4.addRun("Double, ").setUnderlineStyle(XLUnderlineDouble);
rt4.addRun("Single Accounting, ").setUnderlineStyle(XLUnderlineSingleAccounting);
rt4.addRun("Double Accounting").setUnderlineStyle(XLUnderlineDoubleAccounting);

wks.cell("A4").value() = rt4;
```

## 5. Combining Multiple Styles

Because the API is fluid, you can chain as many styling methods as you need onto a single run.

```cpp
XLRichText rt5;
rt5.addRun("MEGA TEXT")
   .setFontName("Impact")
   .setFontSize(30)
   .setFontColor(XLColor("FF00FF"))
   .setBold()
   .setItalic()
   .setStrikethrough(true)
   .setUnderlineStyle(XLUnderlineDouble);

wks.cell("A5").value() = rt5;
```