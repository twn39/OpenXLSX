#include "OpenXLSX.hpp"
#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("Advanced Dynamic Pivot Table Builder", "[XLPivotTable]")
{
    XLDocument doc;
    doc.create("./PivotBuilderTest.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Sample data
    wks.cell("A1").value() = "Category";
    wks.cell("B1").value() = "Product";
    wks.cell("C1").value() = "Sales";
    wks.cell("D1").value() = "Date";

    wks.cell("A2").value() = "Fruit";
    wks.cell("B2").value() = "Apple";
    wks.cell("C2").value() = 120;
    wks.cell("D2").value() = "Q1";

    wks.cell("A3").value() = "Fruit";
    wks.cell("B3").value() = "Banana";
    wks.cell("C3").value() = 250;
    wks.cell("D3").value() = "Q2";

    wks.cell("A4").value() = "Vegetable";
    wks.cell("B4").value() = "Carrot";
    wks.cell("C4").value() = 150;
    wks.cell("D4").value() = "Q1";

    XLPivotTableOptions options;
    options.name        = "AdvancedPivot";
    options.sourceRange = "Sheet1!A1:D4";
    options.targetCell  = "F1";

    // Adding dynamic fields
    options.filters.push_back({"Date", XLPivotSubtotal::Sum, ""}); // Page Field
    options.rows.push_back({"Category", XLPivotSubtotal::Sum, ""}); // Row Field
    options.columns.push_back({"Product", XLPivotSubtotal::Sum, ""}); // Col Field
    options.data.push_back({"Sales", XLPivotSubtotal::Average, "Average Sales"}); // Data Field with Average

    // Configure layout and styles based on our new Builder flags
    options.rowGrandTotals = false;
    options.colGrandTotals = false;
    options.compactData = false;
    options.showRowStripes = true;
    options.pivotTableStyleName = "PivotStyleMedium9";

    REQUIRE_NOTHROW(wks.addPivotTable(options));
    REQUIRE_NOTHROW(doc.save());

    // Verify written values
    XLDocument doc2;
    REQUIRE_NOTHROW(doc2.open("./PivotBuilderTest.xlsx"));

    std::string ptDefXmlStr = doc2.extractXmlFromArchive("xl/pivotTables/pivotTable1.xml");

    // Verify boolean flags translated to XML attributes
    REQUIRE(ptDefXmlStr.find("rowGrandTotals=\"0\"") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("colGrandTotals=\"0\"") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("compactData=\"0\"") != std::string::npos);
    
    // Verify style
    REQUIRE(ptDefXmlStr.find("PivotStyleMedium9") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("showRowStripes=\"1\"") != std::string::npos);

    // Verify Page Field (Filter)
    REQUIRE(ptDefXmlStr.find("<pageFields count=\"1\">") != std::string::npos);
    
    // Check Average subtotal setting
    REQUIRE(ptDefXmlStr.find("subtotal=\"average\"") != std::string::npos);
    REQUIRE(ptDefXmlStr.find("name=\"Average Sales\"") != std::string::npos);
}
