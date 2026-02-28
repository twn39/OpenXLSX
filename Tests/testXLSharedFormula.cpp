#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>

using namespace OpenXLSX;

TEST_CASE("Shared Formulas Reading and Shifting", "[XLFormula]")
{
    // The shared_formula_test.xlsx file is created manually and placed in the Tests directory.
    // Structure:
    // A1 (Master, si=0): B1+C1+$D$1+E$1
    // A2 (Slave, si=0): (Empty in XML) -> Should be B2+C2+$D$1+E$1
    
    SECTION("Read from pre-existing shared formula file")
    {
        XLDocument doc;
        doc.open("./Tests/shared_formula_test.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        // 1. Verify Master Cell (A1)
        auto cellA1 = wks.cell("A1");
        REQUIRE(cellA1.hasFormula());
        REQUIRE(cellA1.formula().get() == "B1+C1+$D$1+E$1");

        // 2. Verify Slave Cell (A2)
        auto cellA2 = wks.cell("A2");
        REQUIRE(cellA2.hasFormula());
        // B1 -> B2, C1 -> C2, $D$1 -> $D$1 (abs), E$1 -> E$1 (row abs)
        REQUIRE(cellA2.formula().get() == "B2+C2+$D$1+E$1");

        doc.close();
    }
}
