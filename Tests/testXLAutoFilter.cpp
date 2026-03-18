#include <catch2/catch_test_macros.hpp>
#include <OpenXLSX.hpp>
#include <fstream>
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("AutoFilter Basic Tests", "[XLAutoFilter]")
{
    XLDocument doc;
    doc.create("AutoFilterTest.xlsx", XLForceOverwrite);
    auto ws = doc.workbook().worksheet("Sheet1");

    ws.cell("A1").value() = "Header1";
    ws.cell("B1").value() = "Header2";
    ws.cell("C1").value() = "Header3";

    ws.cell("A2").value() = 1;
    ws.cell("B2").value() = 10;
    ws.cell("C2").value() = "Apple";

    ws.cell("A3").value() = 2;
    ws.cell("B3").value() = 20;
    ws.cell("C3").value() = "Banana";

    ws.cell("A4").value() = 3;
    ws.cell("B4").value() = 30;
    ws.cell("C4").value() = "Apple";

    ws.setAutoFilter(ws.range(XLCellReference("A1"), XLCellReference("C4")));

    auto filter = ws.autofilterObject();
    REQUIRE(filter);
    REQUIRE(filter.ref() == "A1:C4");

    // Add value filter on column 2 (C)
    auto col2 = filter.filterColumn(2);
    col2.addFilter("Apple");
    REQUIRE(col2.colId() == 2);

    // Add custom filter on column 1 (B)
    auto col1 = filter.filterColumn(1);
    col1.setCustomFilter("greaterThan", "15");

    // Add top10 on column 0 (A)
    auto col0 = filter.filterColumn(0);
    col0.setTop10(2, false, true); // Top 2

    doc.save();
    doc.close();

    // Reopen and check
    XLDocument doc2;
    doc2.open("AutoFilterTest.xlsx");
    auto ws2 = doc2.workbook().worksheet("Sheet1");
    auto filter2 = ws2.autofilterObject();

    REQUIRE(filter2);
    REQUIRE(filter2.ref() == "A1:C4");

    auto c2 = filter2.filterColumn(2);
    REQUIRE(c2.colId() == 2);
    
    doc2.close();
}

TEST_CASE("Table AutoFilter Tests", "[XLTable][XLAutoFilter]")
{
    XLDocument doc;
    doc.create("TableAutoFilterTest.xlsx", XLForceOverwrite);
    auto ws = doc.workbook().worksheet("Sheet1");

    ws.cell("A1").value() = "Id";
    ws.cell("B1").value() = "Score";

    ws.cell("A2").value() = 1;
    ws.cell("B2").value() = 100;

    ws.cell("A3").value() = 2;
    ws.cell("B3").value() = 80;

    auto& table = ws.tables();
    table.setName("MyTable");
    table.setDisplayName("MyTableDisplay");
    table.setRangeReference("A1:B3");

    auto filter = table.autoFilter();
    REQUIRE(filter);
    REQUIRE(filter.ref() == "A1:B3");

    auto col = filter.filterColumn(1);
    col.setCustomFilter("greaterThan", "90");

    doc.save();
    doc.close();

    XLDocument doc2;
    doc2.open("TableAutoFilterTest.xlsx");
    auto ws2 = doc2.workbook().worksheet("Sheet1");
    auto& table2 = ws2.tables();
    auto filter2 = table2.autoFilter();

    REQUIRE(filter2);
    REQUIRE(filter2.ref() == "A1:B3");

    auto c = filter2.filterColumn(1);
    REQUIRE(c.colId() == 1);

    doc2.close();
}
