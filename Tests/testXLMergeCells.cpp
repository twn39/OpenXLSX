#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>

using namespace OpenXLSX;

TEST_CASE("XLMergeCells Tests", "[XLMergeCells]")
{
    SECTION("Basic Merge and Unmerge")
    {
        XLDocument doc;
        doc.create("./testXLMergeCells.xlsx", XLForceOverwrite);
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

    SECTION("Multiple Merges")
    {
        XLDocument doc;
        doc.create("./testXLMultipleMerges.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.mergeCells("A1:A2");
        wks.mergeCells("C1:C2");
        REQUIRE(wks.merges().count() == 2);

        REQUIRE(wks.merges().findMerge("A1:A2") == 0);
        REQUIRE(wks.merges().findMerge("C1:C2") == 1);

        doc.close();
    }

    SECTION("Find Merge by Cell")
    {
        XLDocument doc;
        doc.create("./testXLMergeByCell.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.mergeCells("B2:C3");

        // Test with strings
        REQUIRE(wks.merges().findMergeByCell("B2") == 0);
        REQUIRE(wks.merges().findMergeByCell("C3") == 0);
        REQUIRE(wks.merges().findMergeByCell("B3") == 0);
        REQUIRE(wks.merges().findMergeByCell("C2") == 0);
        REQUIRE(wks.merges().findMergeByCell("A1") == XLMergeNotFound);

        // Test with XLCellReference (Numerical path)
        REQUIRE(wks.merges().findMergeByCell(XLCellReference(2, 2)) == 0);
        REQUIRE(wks.merges().findMergeByCell(XLCellReference(3, 3)) == 0);
        REQUIRE(wks.merges().findMergeByCell(XLCellReference(1, 1)) == XLMergeNotFound);

        doc.close();
    }

    SECTION("Numerical Cache Consistency after Deletion")
    {
        XLDocument doc;
        doc.create("./testXLMergeCacheConsistency.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.mergeCells("A1:B2");
        wks.mergeCells("D4:E5");
        wks.mergeCells("G7:H8");

        REQUIRE(wks.merges().count() == 3);
        REQUIRE(wks.merges().findMergeByCell("E5") == 1);

        // Delete middle one
        wks.unmergeCells("D4:E5");
        REQUIRE(wks.merges().count() == 2);

        // Check that subsequent index shifted correctly in cache
        REQUIRE(wks.merges().findMerge("G7:H8") == 1);
        REQUIRE(wks.merges().findMergeByCell("H8") == 1);
        REQUIRE(wks.merges().findMergeByCell("E5") == XLMergeNotFound);

        doc.close();
    }

    SECTION("Overlap Detection (Optimized Path)")
    {
        XLDocument doc;
        doc.create("./testXLMergeOverlap.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.mergeCells("B2:D4");

        // Overlapping cases (triggers XLRect::overlaps)
        REQUIRE_THROWS_AS(wks.mergeCells("C3:E5"), XLInputError);    // Partial overlap
        REQUIRE_THROWS_AS(wks.mergeCells("B2:D4"), XLInputError);    // Exact same
        REQUIRE_THROWS_AS(wks.mergeCells("A1:E5"), XLInputError);    // Encompassing

        // Touch but not overlap
        REQUIRE_NOTHROW(wks.mergeCells("E2:F4"));    // Right touch
        REQUIRE_NOTHROW(wks.mergeCells("B5:D6"));    // Bottom touch

        doc.close();
    }

    SECTION("Empty Hidden Cells")
    {
        XLDocument doc;
        doc.create("./testXLEmptyHiddenCells.xlsx", XLForceOverwrite);
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
