#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <filesystem>

using namespace OpenXLSX;

TEST_CASE("Rich Text Fluid API Tests", "[XLRichText]")
{
    SECTION("Create and Serialize Rich Text")
    {
        XLDocument doc;
        doc.create("richtext_test.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        // Use the fluid API to construct a rich text object
        XLRichText rt;
        rt.addRun("Red Bold ").setFontColor(XLColor("FF0000")).setBold();
        rt.addRun("Blue Italic ").setFontColor(XLColor("0000FF")).setItalic();
        rt.addRun("Normal text");

        wks.cell("A1").value() = rt;
        doc.save();
        doc.close();

        // Re-open and verify
        XLDocument doc2;
        doc2.open("richtext_test.xlsx");
        auto wks2 = doc2.workbook().worksheet("Sheet1");

        auto cellType = wks2.cell("A1").value().type();
        REQUIRE(cellType == XLValueType::RichText);

        auto rtRead = wks2.cell("A1").value().get<XLRichText>();
        auto runs   = rtRead.runs();

        REQUIRE(runs.size() == 3);

        REQUIRE(runs[0].text() == "Red Bold ");
        REQUIRE(runs[0].bold() == true);
        REQUIRE(runs[0].fontColor()->hex() == "FFFF0000");    // OpenXLSX often prepends FF for alpha

        REQUIRE(runs[1].text() == "Blue Italic ");
        REQUIRE(runs[1].italic() == true);
        REQUIRE(runs[1].bold().has_value() == false);

        REQUIRE(runs[2].text() == "Normal text");
        REQUIRE(runs[2].bold().has_value() == false);
        REQUIRE(runs[2].italic().has_value() == false);
        REQUIRE(runs[2].fontColor().has_value() == false);

        doc2.close();
        std::filesystem::remove("richtext_test.xlsx");
    }
}