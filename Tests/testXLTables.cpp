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

            REQUIRE(tables.name() == "MyTable");
            REQUIRE(tables.displayName() == "MyTableDisplay");
            REQUIRE(tables.rangeReference() == "A1:C10");
            REQUIRE(tables.styleName() == "TableStyleMedium2");
            REQUIRE(tables.showRowStripes() == true);
            REQUIRE(tables.showColumnStripes() == false);
            REQUIRE(tables.showFirstColumn() == true);
            REQUIRE(tables.showLastColumn() == false);

            // Test invalid name (spaces)
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

            doc.close();
        }
    }
}
