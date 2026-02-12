#include <catch.hpp>
#include <OpenXLSX.hpp>

using namespace OpenXLSX;

TEST_CASE("XLStyles Tests", "[XLStyles]") {
    SECTION("Number Formats") {
        XLDocument doc;
        doc.create("./testXLStylesNumFmt.xlsx");
        auto styles = doc.styles();

        auto numFmts = styles.numberFormats();
        size_t initialCount = numFmts.count();
        
        XLNumberFormat nf;
        size_t idx = numFmts.create(nf);
        numFmts[idx].setNumberFormatId(100);
        numFmts[idx].setFormatCode("0.00%");
        
        REQUIRE(numFmts.count() == initialCount + 1);
        REQUIRE(numFmts[idx].numberFormatId() == 100);
        REQUIRE(numFmts[idx].formatCode() == "0.00%");
        
        doc.close();
    }

    SECTION("Fonts") {
        XLDocument doc;
        doc.create("./testXLStylesFonts.xlsx");
        auto styles = doc.styles();
        auto fonts = styles.fonts();
        
        size_t idx = fonts.create();
        auto font = fonts[idx];
        
        font.setFontName("Courier New");
        font.setFontSize(14);
        font.setBold(true);
        font.setItalic(true);
        font.setFontColor(XLColor(255, 0, 0));
        font.setUnderline(XLUnderlineDouble);
        
        REQUIRE(font.fontName() == "Courier New");
        REQUIRE(font.fontSize() == 14);
        REQUIRE(font.bold() == true);
        REQUIRE(font.italic() == true);
        REQUIRE(font.underline() == XLUnderlineDouble);
        
        doc.close();
    }

    SECTION("Fills") {
        XLDocument doc;
        doc.create("./testXLStylesFills.xlsx");
        auto styles = doc.styles();
        auto fills = styles.fills();
        
        size_t idx = fills.create();
        auto fill = fills[idx];
        
        fill.setPatternType(XLPatternSolid);
        fill.setColor(XLColor(0, 255, 0));
        fill.setBackgroundColor(XLColor(0, 0, 255));
        
        REQUIRE(fill.patternType() == XLPatternSolid);
        
        doc.close();
    }

    SECTION("Borders") {
        XLDocument doc;
        doc.create("./testXLStylesBorders.xlsx");
        auto styles = doc.styles();
        auto borders = styles.borders();
        
        size_t idx = borders.create();
        auto border = borders[idx];
        
        border.setLeft(XLLineStyleThick, XLColor(0, 0, 0));
        border.setRight(XLLineStyleDashed, XLColor(255, 0, 0));
        border.setTop(XLLineStyleDouble, XLColor(0, 255, 0));
        border.setBottom(XLLineStyleThin, XLColor(0, 0, 255));
        
        REQUIRE(border.left().style() == XLLineStyleThick);
        REQUIRE(border.right().style() == XLLineStyleDashed);
        
        doc.close();
    }

    SECTION("Cell Formats (Xf)") {
        XLDocument doc;
        doc.create("./testXLStylesXf.xlsx");
        auto styles = doc.styles();
        
        auto cellFormats = styles.cellFormats();
        size_t idx = cellFormats.create();
        auto xf = cellFormats[idx];
        
        xf.setFontIndex(1);
        xf.setFillIndex(1);
        xf.setBorderIndex(1);
        xf.setApplyFont(true);
        xf.setApplyFill(true);
        xf.setApplyBorder(true);
        
        REQUIRE(xf.fontIndex() == 1);
        REQUIRE(xf.fillIndex() == 1);
        REQUIRE(xf.borderIndex() == 1);
        REQUIRE(xf.applyFont() == true);
        
        doc.close();
    }
}
