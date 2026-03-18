#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <sstream>

using namespace OpenXLSX;

TEST_CASE("XLTables Basic Tests", "[XLTables]")
{
    const std::string filename = "TestTables.xlsx";

    SECTION("Table Creation and Validation")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");

            // Trigger table creation
            REQUIRE_FALSE(wks.hasTables());
            auto& tables = wks.tables();
            REQUIRE(wks.hasTables());
            REQUIRE(tables.valid());

            // Test basic properties
            tables.setName("MyTable");
            tables.setDisplayName("MyTableDisplay");
            tables.setRangeReference("A1:C10");
            tables.setStyleName("TableStyleMedium2");
            tables.setShowRowStripes(true);
            tables.setShowColumnStripes(false);
            tables.setShowFirstColumn(true);
            tables.setShowLastColumn(false);
            tables.setShowHeaderRow(false);
            tables.setShowTotalsRow(true);

            REQUIRE(tables.name() == "MyTable");
            REQUIRE(tables.displayName() == "MyTableDisplay");
            REQUIRE(tables.rangeReference() == "A1:C10");
            REQUIRE(tables.styleName() == "TableStyleMedium2");
            REQUIRE(tables.showRowStripes() == true);
            REQUIRE(tables.showColumnStripes() == false);
            REQUIRE(tables.showFirstColumn() == true);
            REQUIRE(tables.showLastColumn() == false);
            REQUIRE(tables.showHeaderRow() == false);
            REQUIRE(tables.showTotalsRow() == true);

            REQUIRE_THROWS_AS(tables.setName("My Table"), XLInputError);
            
            doc.save();
            doc.close();
        }

        {
            XLDocument doc;
            doc.open(filename);
            auto wks = doc.workbook().worksheet("Sheet1");

            REQUIRE(wks.hasTables());
            auto& tables = wks.tables();
            REQUIRE(tables.valid());

            REQUIRE(tables.name() == "MyTable");
            REQUIRE(tables.displayName() == "MyTableDisplay");
            REQUIRE(tables.rangeReference() == "A1:C10");
            REQUIRE(tables.styleName() == "TableStyleMedium2");
            REQUIRE(tables.showRowStripes() == true);
            REQUIRE(tables.showColumnStripes() == false);
            REQUIRE(tables.showFirstColumn() == true);
            REQUIRE(tables.showLastColumn() == false);
            REQUIRE(tables.showHeaderRow() == false);
            REQUIRE(tables.showTotalsRow() == true);

            doc.close();
        }
    }

    SECTION("Table Columns and Totals")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            auto& tables = wks.tables();
            tables.setRangeReference("A1:C10");
            tables.setShowTotalsRow(true);

            auto col1 = tables.appendColumn("ID");
            auto col2 = tables.appendColumn("Name");
            auto col3 = tables.appendColumn("Score");

            REQUIRE(col1.id() == 1);
            REQUIRE(col1.name() == "ID");
            REQUIRE(col2.id() == 2);
            REQUIRE(col2.name() == "Name");
            REQUIRE(col3.id() == 3);
            REQUIRE(col3.name() == "Score");

            col1.setTotalsRowLabel("Total:");
            col3.setTotalsRowFunction(XLTotalsRowFunction::Sum);

            REQUIRE(col1.totalsRowLabel() == "Total:");
            REQUIRE(col1.totalsRowFunction() == XLTotalsRowFunction::None);
            
            REQUIRE(col3.totalsRowFunction() == XLTotalsRowFunction::Sum);
            REQUIRE(col3.totalsRowLabel() == "");

            auto fetchedCol = tables.column("Score");
            REQUIRE(fetchedCol);
            REQUIRE(fetchedCol.totalsRowFunction() == XLTotalsRowFunction::Sum);

            doc.save();
            doc.close();
        }

        {
            XLDocument doc;
            doc.open(filename);
            auto wks = doc.workbook().worksheet("Sheet1");
            auto& tables = wks.tables();

            REQUIRE(tables.showTotalsRow() == true);
            
            auto col1 = tables.column(1);
            REQUIRE(col1);
            REQUIRE(col1.name() == "ID");
            REQUIRE(col1.totalsRowLabel() == "Total:");
            REQUIRE(col1.totalsRowFunction() == XLTotalsRowFunction::None);

            auto col3 = tables.column("Score");
            REQUIRE(col3);
            REQUIRE(col3.totalsRowFunction() == XLTotalsRowFunction::Sum);

            doc.close();
        }
    }

    SECTION("Table Auto Resize")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            
            // Initial tiny table setup
            auto& tables = wks.tables();
            tables.setRangeReference("B2:C3"); // Table starts at B2
            
            wks.cell("B2").value() = "Header1";
            wks.cell("C2").value() = "Header2";
            wks.cell("B3").value() = "Val1";
            wks.cell("C3").value() = "Val2";

            REQUIRE(tables.rangeReference() == "B2:C3");

            // User expands the data programmatically
            wks.cell("B100").value() = "Val99";
            wks.cell("E100").value() = "Val99_colE";

            // Automatically resize the table to fit the new data
            tables.resizeToFitData(wks);

            // B2 was top left. lastCell is E100.
            REQUIRE(tables.rangeReference() == "B2:E100");
            
            // Ensure internal nodes sync up
            REQUIRE(tables.autoFilter().ref() == "B2:E100");

            doc.save();
            doc.close();
        }

        {
            XLDocument doc;
            doc.open(filename);
            auto wks = doc.workbook().worksheet("Sheet1");
            auto& tables = wks.tables();

            REQUIRE(tables.rangeReference() == "B2:E100");
            REQUIRE(tables.autoFilter().ref() == "B2:E100");

            doc.close();
        }
    }

    SECTION("Comprehensive Demo Generation")
    {
        XLDocument doc;
        doc.create("TableFeatureDemo_Final.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        // 1. Write headers and data
        wks.cell("B2").value() = "Employee ID";
        wks.cell("C2").value() = "Name";
        wks.cell("D2").value() = "Sales";
        wks.cell("E2").value() = "Bonus";

        wks.cell("B3").value() = 1001;
        wks.cell("C3").value() = "Alice";
        wks.cell("D3").value() = 5000;
        wks.cell("E3").value() = 500;

        wks.cell("B4").value() = 1002;
        wks.cell("C4").value() = "Bob";
        wks.cell("D4").value() = 4200;
        wks.cell("E4").value() = 420;

        wks.cell("B5").value() = 1003;
        wks.cell("C5").value() = "Charlie";
        wks.cell("D5").value() = 7500;
        wks.cell("E5").value() = 750;

        // 2. Create and configure table
        auto& table = wks.tables();
        table.setName("SalesTable");
        table.setDisplayName("SalesTable");
        
        // Use resizeToFitData to find the current data range (B2:E5)
        table.resizeToFitData(wks); 

        // 3. Enable Totals Row
        // NOTE: Since the data is in B2:E5, enabling the totals row will 
        // expand the table range to B2:E6. The resizeToFitData call should be done
        // OR we manually set the range. 
        // Actually, resizeToFitData only looks at cells with values. 
        // Let's manually ensure the range is correct for a totals row.
        table.setRangeReference("B2:E6"); 
        table.setShowTotalsRow(true);

        // 4. Style and formatting
        table.setStyleName("TableStyleMedium9");
        table.setShowRowStripes(true);
        table.setShowFirstColumn(true);

        // 5. Column specifics
        auto col1 = table.appendColumn("Employee ID");
        col1.setTotalsRowLabel("Grand Total:");
        wks.cell("B6").value() = "Grand Total:"; // Manually write to the cell

        auto col2 = table.appendColumn("Name");
        col2.setTotalsRowFunction(XLTotalsRowFunction::Count);
        // SUBTOTAL(103, ...) is COUNTA. For demo, we can just write the static result or formula.
        // Excel table totals usually use SUBTOTAL to respect filters.
        wks.cell("C6").formula() = "SUBTOTAL(103,SalesTable[Name])"; 

        auto col3 = table.appendColumn("Sales");
        col3.setTotalsRowFunction(XLTotalsRowFunction::Sum);
        wks.cell("D6").formula() = "SUBTOTAL(109,SalesTable[Sales])"; // 109 is SUM

        auto col4 = table.appendColumn("Bonus");
        col4.setTotalsRowFunction(XLTotalsRowFunction::Average);
        wks.cell("E6").formula() = "SUBTOTAL(101,SalesTable[Bonus])"; // 101 is AVERAGE

        // 6. AutoFilter with Custom rule
        auto filter = table.autoFilter();
        // Column Index 2 is "Sales". Filter for > 4500
        filter.filterColumn(2).setCustomFilter("greaterThan", "4500");

        doc.save();
        doc.close();

        REQUIRE(true);
    }
}
