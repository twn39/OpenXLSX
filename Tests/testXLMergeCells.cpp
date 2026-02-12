#include <catch.hpp>
#include <OpenXLSX.hpp>

using namespace OpenXLSX;

TEST_CASE("XLMergeCells Tests", "[XLMergeCells]") {
    SECTION("Basic Merge and Unmerge") {
        XLDocument doc;
        doc.create("./testXLMergeCells.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.mergeCells("A1:B2");
        REQUIRE(wks.merges().count() == 1);
        REQUIRE(std::string(wks.merges().merge(0)) == "A1:B2");
        REQUIRE(wks.merges().mergeExists("A1:B2") == true);

        wks.unmergeCells("A1:B2");
        REQUIRE(wks.merges().count() == 0);
        REQUIRE(wks.merges().mergeExists("A1:B2") == false);

        doc.close();
    }

    SECTION("Multiple Merges") {
        XLDocument doc;
        doc.create("./testXLMultipleMerges.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.mergeCells("A1:A2");
        wks.mergeCells("C1:C2");
        REQUIRE(wks.merges().count() == 2);

        REQUIRE(wks.merges().findMerge("A1:A2") == 0);
        REQUIRE(wks.merges().findMerge("C1:C2") == 1);

        doc.close();
    }

    SECTION("Find Merge by Cell") {
        XLDocument doc;
        doc.create("./testXLMergeByCell.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.mergeCells("B2:C3");
        
        REQUIRE(wks.merges().findMergeByCell("B2") == 0);
        REQUIRE(wks.merges().findMergeByCell("C3") == 0);
        REQUIRE(wks.merges().findMergeByCell("B3") == 0);
        REQUIRE(wks.merges().findMergeByCell("C2") == 0);
        REQUIRE(wks.merges().findMergeByCell("A1") == XLMergeNotFound);

        doc.close();
    }

    SECTION("Overlap Detection") {
        XLDocument doc;
        doc.create("./testXLMergeOverlap.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.mergeCells("B2:D4");
        
        // Overlapping cases
        REQUIRE_THROWS_AS(wks.mergeCells("C3:E5"), XLInputError); // Partial overlap
        REQUIRE_THROWS_AS(wks.mergeCells("B2:D4"), XLInputError); // Exact same
        REQUIRE_THROWS_AS(wks.mergeCells("C3:C3"), XLInputError); // Inside (though 1x1 is also invalid)
        
        // Non-overlapping
        REQUIRE_NOTHROW(wks.mergeCells("E5:F6"));

        doc.close();
    }

    SECTION("Empty Hidden Cells") {
        XLDocument doc;
        doc.create("./testXLEmptyHiddenCells.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("A1").value() = 1;
        wks.cell("A2").value() = 2;
        wks.cell("B1").value() = 3;
        wks.cell("B2").value() = 4;

        wks.mergeCells("A1:B2", XLEmptyHiddenCells);

        REQUIRE(wks.cell("A1").value().get<int>() == 1);
        REQUIRE(wks.cell("A2").value().type() == XLValueType::Empty);
        REQUIRE(wks.cell("B1").value().type() == XLValueType::Empty);
        REQUIRE(wks.cell("B2").value().type() == XLValueType::Empty);

        doc.close();
    }
}
