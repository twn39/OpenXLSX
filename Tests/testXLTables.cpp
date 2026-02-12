#include <catch.hpp>
#include <OpenXLSX.hpp>
#include <sstream>

using namespace OpenXLSX;

TEST_CASE("XLTables Basic Tests", "[XLTables]") {
    const std::string filename = "TestTables.xlsx";

    SECTION("Table Creation and Validation") {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            
            // Trigger table creation
            REQUIRE_FALSE(wks.hasTables());
            auto& tables = wks.tables();
            REQUIRE(wks.hasTables());
            REQUIRE(tables.valid());

            // Even if print is empty for now (no columns/data), we've covered the creation path.
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
            doc.close();
        }
    }
}
