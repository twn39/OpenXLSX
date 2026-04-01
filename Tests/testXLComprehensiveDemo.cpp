#include <OpenXLSX.hpp>
#include <iostream>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("OpenXLSX_Comprehensive_Demo.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");
    wks.setName("Demo Features");

    // 1. Basic & Styled Data
    wks.cell("A1").value() = "Feature Description";
    wks.cell("B1").value() = "Live Output Demonstration";
    
    // 2. Rich Text (Multicolored strings in one cell)
    wks.cell("A2").value() = "Rich Text (New Fluent API)";
    XLRichText rt;
    rt.addRun("OpenXLSX is ").setFontColor(XLColor("000000")).setBold(false);
    rt.addRun("Fast").setFontColor(XLColor("00B050")).setBold(true).setItalic(true);
    rt.addRun(", ").setFontColor(XLColor("000000"));
    rt.addRun("Reliable").setFontColor(XLColor("0070C0")).setBold(true).setUnderline(true);
    rt.addRun(", and ").setFontColor(XLColor("000000"));
    rt.addRun("Modern").setFontColor(XLColor("7030A0")).setBold(true).setFontSize(14).setFontName("Arial");
    rt.addRun("!").setFontColor(XLColor("000000"));
    wks.cell("B2").value() = rt;

    // 3. String to Number Coercion (Zip codes)
    wks.cell("A3").value() = "String Preserving Leading Zeros";
    wks.cell("B3").value() = "00123456789"; // Should remain a string, not 123456789

    // 4. Extreme Bounds (Row 1 Million)
    wks.cell("A4").value() = "Extreme Boundary Cell";
    wks.cell("B4").value() = "Scroll to cell XFD1048576 to see the hidden message!";
    wks.cell("XFD1048576").value() = "You found the edge of the Excel Universe! (Row 1048576, Col 16384)";

    // 5. 1900 Leap Year Date Bug 
    wks.cell("A5").value() = "1900 Leap Year (Serial 60.0)";
    XLDateTime date60(60.0); // 1900-02-29
    wks.cell("B5").value() = date60.serial();

    // 6. Formulas
    wks.cell("A6").value() = "Formula Engine Calculation";
    wks.cell("C6").value() = 100;
    wks.cell("D6").value() = 250;
    wks.cell("B6").formula() = "SUM(C6:D6)"; 
    // Let's run the engine so it writes the cached value
    XLFormulaEngine eng;
    auto resolver = XLFormulaEngine::makeResolver(wks);
    wks.cell("B6").value() = eng.evaluate("=SUM(C6:D6)", resolver);
    wks.cell("B6").formula() = "SUM(C6:D6)"; // Restore formula

    // 7. Image Insertion (We'll use the test.png from the Tests folder)
    wks.cell("A8").value() = "Image Insertion (TwoCell Anchor at B8)";
    XLImageOptions opts;
    opts.positioning = XLImagePositioning::TwoCell;
    opts.scaleX = 1.0;
    opts.scaleY = 1.0;
    // We will place this at B8. Since the compiled binary runs in build_asan/, we'll reference the copied png
    wks.insertImage("B8", "Tests/test.png", opts);

    // 8. Image Insertion (Absolute pixels)
    wks.cell("A15").value() = "Image Insertion (Absolute offset)";
    XLImageOptions optsAbs;
    optsAbs.positioning = XLImagePositioning::Absolute;
    optsAbs.offsetX = 400; // X pixels
    optsAbs.offsetY = 300; // Y pixels
    wks.insertImage("A1", "Tests/test.jpg", optsAbs);

    // Save
    doc.save();
    doc.close();

    std::cout << "Successfully generated OpenXLSX_Comprehensive_Demo.xlsx in the current directory!" << std::endl;
    return 0;
}