#include <catch.hpp>
#include <OpenXLSX.hpp>

using namespace OpenXLSX;

TEST_CASE("XLStyles Integration Tests", "[XLStyles]") {
    SECTION("Font Style Round-trip") {
        const std::string filename = "StyleIntegrationTest.xlsx";
        
        // 1. Create and apply style
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            auto styles = doc.styles();
            
            // Create a new font
            size_t fontIdx = styles.fonts().create();
            auto font = styles.fonts()[fontIdx];
            font.setFontName("Courier New");
            font.setFontSize(16);
            font.setBold(true);
            
            // Create a cell format (xf) that uses this font
            size_t xfIdx = styles.cellFormats().create();
            styles.cellFormats()[xfIdx].setFontIndex(fontIdx);
            styles.cellFormats()[xfIdx].setApplyFont(true);
            
            // Apply format to A1
            wks.cell("A1").value() = "Styled Text";
            wks.cell("A1").setCellFormat(xfIdx);
            
            doc.save();
            doc.close();
        }
        
        // 2. Re-open and verify
        {
            XLDocument doc;
            doc.open(filename);
            auto wks = doc.workbook().worksheet("Sheet1");
            auto styles = doc.styles();
            
            auto cell = wks.cell("A1");
            size_t xfIdx = cell.cellFormat();
            REQUIRE(xfIdx > 0); // Should not be default 0
            
            auto xf = styles.cellFormats()[xfIdx];
            size_t fontIdx = xf.fontIndex();
            
            auto font = styles.fonts()[fontIdx];
            REQUIRE(font.fontName() == "Courier New");
            REQUIRE(font.fontSize() == 16);
            REQUIRE(font.bold() == true);
            
            doc.close();
        }
    }

    SECTION("Fill Style Round-trip") {
        const std::string filename = "FillIntegrationTest.xlsx";
        
        // 1. Create and apply fill
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            auto styles = doc.styles();
            
            size_t fillIdx = styles.fills().create();
            auto fill = styles.fills()[fillIdx];
            fill.setPatternType(XLPatternSolid);
            fill.setColor(XLColor(255, 255, 0)); // Yellow
            
            size_t xfIdx = styles.cellFormats().create();
            styles.cellFormats()[xfIdx].setFillIndex(fillIdx);
            styles.cellFormats()[xfIdx].setApplyFill(true);
            
            wks.cell("B2").setCellFormat(xfIdx);
            
            doc.save();
            doc.close();
        }
        
        // 2. Re-open and verify
        {
            XLDocument doc;
            doc.open(filename);
            auto wks = doc.workbook().worksheet("Sheet1");
            auto styles = doc.styles();
            
            auto cell = wks.cell("B2");
            size_t xfIdx = cell.cellFormat();
            auto fillIdx = styles.cellFormats()[xfIdx].fillIndex();
            
            auto fill = styles.fills()[fillIdx];
            REQUIRE(fill.patternType() == XLPatternSolid);
            // REQUIRE(fill.color().hex() == "ffffff00"); 
            
            doc.close();
        }
    }
}
