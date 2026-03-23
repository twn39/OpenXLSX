#include <catch2/catch_test_macros.hpp>
#include "OpenXLSX.hpp"
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("Dynamic Pivot Table Generation", "[XLPivotTable]")
{
    // === Functionality Setup ===
    XLDocument doc;
    doc.create("./PivotTest.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Add some sample data
    wks.cell("A1").value() = "Region";
    wks.cell("B1").value() = "Product";
    wks.cell("C1").value() = "Sales";
    
    wks.cell("A2").value() = "North";
    wks.cell("B2").value() = "Apples";
    wks.cell("C2").value() = 100;
    
    wks.cell("A3").value() = "South";
    wks.cell("B3").value() = "Bananas";
    wks.cell("C3").value() = 300;

    wks.cell("A4").value() = "North";
    wks.cell("B4").value() = "Oranges";
    wks.cell("C4").value() = 150;

    XLPivotTableOptions options;
    options.name = "Pivot1";
    options.sourceRange = "Sheet1!A1:C4";
    options.targetCell = "E1";
    
    options.rows.push_back({"Region", XLPivotSubtotal::Sum, ""});
    options.columns.push_back({"Product", XLPivotSubtotal::Sum, ""});
    options.data.push_back({"Sales", XLPivotSubtotal::Sum, "Total Sales"});

    // Execute dynamic construction
    wks.addPivotTable(options);

    doc.save();
    doc.close();

    // === Read-back and OOXML Verification ===
    XLDocument doc2;
    doc2.open("./PivotTest.xlsx");
    
    // Verify XML structure in Pivot Cache Definition
    std::string cacheDefXmlStr = doc2.extractXmlFromArchive("xl/pivotCache/pivotCacheDefinition1.xml");
    
    // Check root node attributes indicating dynamic refresh
    REQUIRE(cacheDefXmlStr.find("saveData=\"0\"") != std::string::npos);
    REQUIRE(cacheDefXmlStr.find("refreshOnLoad=\"1\"") != std::string::npos);
    
    // Check dynamically analyzed fields from worksheet headers
    REQUIRE(cacheDefXmlStr.find("<cacheFields count=\"3\">") != std::string::npos);
    REQUIRE(cacheDefXmlStr.find("<cacheField name=\"Region\" numFmtId=\"0\">") != std::string::npos);
    REQUIRE(cacheDefXmlStr.find("<cacheField name=\"Product\" numFmtId=\"0\">") != std::string::npos);
    REQUIRE(cacheDefXmlStr.find("<cacheField name=\"Sales\" numFmtId=\"0\">") != std::string::npos);
    
    // Check proper missing item placeholders to avoid Excel corruption
    bool hasMissingItems = cacheDefXmlStr.find("<sharedItems containsBlank=\"1\" count=\"0\"><m/></sharedItems>") != std::string::npos ||
                           cacheDefXmlStr.find("<m />") != std::string::npos ||
                           cacheDefXmlStr.find("<m/>") != std::string::npos;
    REQUIRE(hasMissingItems);

    // Verify XML structure in Pivot Table Definition
    std::string ptDefXmlStr = doc2.extractXmlFromArchive("xl/pivotTables/pivotTable1.xml");
    
    // Location and layout
    REQUIRE(ptDefXmlStr.find("<location ref=\"E1\"") != std::string::npos);
    
    // Check axis mapping and specific element omission formatting
    REQUIRE(ptDefXmlStr.find("<pivotField axis=\"axisRow\"") != std::string::npos); // Region
    REQUIRE(ptDefXmlStr.find("<pivotField axis=\"axisCol\"") != std::string::npos); // Product
    bool hasDataField = ptDefXmlStr.find("<pivotField dataField=\"1\" showAll=\"0\" />") != std::string::npos ||
                        ptDefXmlStr.find("<pivotField dataField=\"1\" showAll=\"0\"/>") != std::string::npos;
    REQUIRE(hasDataField); // Sales
    
    // Check proper field mappings
    REQUIRE(ptDefXmlStr.find("<rowFields count=\"1\">") != std::string::npos);
    bool hasRowField = ptDefXmlStr.find("<field x=\"0\" />") != std::string::npos || ptDefXmlStr.find("<field x=\"0\"/>") != std::string::npos;
    REQUIRE(hasRowField);

    REQUIRE(ptDefXmlStr.find("<colFields count=\"1\">") != std::string::npos);
    bool hasColField = ptDefXmlStr.find("<field x=\"1\" />") != std::string::npos || ptDefXmlStr.find("<field x=\"1\"/>") != std::string::npos;
    REQUIRE(hasColField);
    
    // Check correct empty items node handling (fixes Excel warning)
    REQUIRE(ptDefXmlStr.find("<colItems count=\"1\">") != std::string::npos);
    bool hasEmptyItem = ptDefXmlStr.find("<i />") != std::string::npos || ptDefXmlStr.find("<i/>") != std::string::npos;
    REQUIRE(hasEmptyItem);
    
    // Check custom data name overriding
    bool hasTotalSales = ptDefXmlStr.find("<dataField name=\"Total Sales\" fld=\"2\" />") != std::string::npos ||
                         ptDefXmlStr.find("<dataField name=\"Total Sales\" fld=\"2\"/>") != std::string::npos;
    REQUIRE(hasTotalSales);

    doc2.close();
}
