#include <catch2/catch_test_macros.hpp>
#include <OpenXLSX.hpp>
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("Conditional Formatting OOXML Structure Validation", "[ConditionalFormatting][OOXML]") {
    XLDocument doc;
    doc.create("CF_OOXML_Validation.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // 1. Color Scale Rule
    auto csRule = XLColorScaleRule(XLColor("FFFF0000"), XLColor("FF00FF00"));
    wks.addConditionalFormatting("A1:A10", csRule);

    // 2. Data Bar Rule (Should not output empty val="")
    auto dbRule = XLDataBarRule(XLColor("FF0000FF"), true);
    wks.addConditionalFormatting("B1:B10", dbRule);

    // 3. Cell Is Rule (Checks uppercase colors and fgColor/bgColor ordering)
    XLDxf dxfC;
    dxfC.font().setFontColor(XLColor("FFFF0000"));
    dxfC.fill().setFillType(XLPatternFill);
    dxfC.fill().setPatternType(XLPatternSolid);
    dxfC.fill().setColor(XLColor("FFFFFF00"));
    dxfC.fill().setBackgroundColor(XLColor("FFFFFF00"));
    auto cellIsRule = XLCellIsRule(">=", "5");
    wks.addConditionalFormatting("C1:C10", cellIsRule, dxfC);

    // 4. Formula Rule (Checks boolean val="1")
    XLDxf dxfD;
    dxfD.font().setBold(true);
    dxfD.font().setUnderline(XLUnderlineSingle);
    auto formulaRule = XLFormulaRule("MOD(ROW(),2)=0");
    wks.addConditionalFormatting(wks.range("D1:D10"), formulaRule, dxfD);

    doc.save();
    
    // Now validate the underlying XML structures directly
    auto cfList = wks.conditionalFormats();
    REQUIRE(cfList.count() == 4);

    // Validate Priority assignments (should be global 1, 2, 3, 4)
    REQUIRE(cfList[0].cfRules()[0].priority() == 1);
    REQUIRE(cfList[1].cfRules()[0].priority() == 2);
    REQUIRE(cfList[2].cfRules()[0].priority() == 3);
    REQUIRE(cfList[3].cfRules()[0].priority() == 4);

    // Validate Data Bar empty value absence
    auto dataBarNode = cfList[1].cfRules()[0].node().child("dataBar");
    auto cfvoMinNode = dataBarNode.child("cfvo");
    REQUIRE(cfvoMinNode.attribute("val").empty() == true); // Should be completely omitted, not val=""

    // Validate DXF configurations in Styles
    auto styles = doc.styles(); // styles is on doc, not workbook
    REQUIRE(styles.dxfs().count() == 2);
    
    // First DXF (Color and Fill)
    auto dxf1 = styles.dxfs()[0].node();
    REQUIRE(std::string(dxf1.child("font").child("color").attribute("rgb").value()) == "FFFF0000"); // Must be uppercase
    auto fillNode = dxf1.child("fill").child("patternFill");
    REQUIRE(std::string(fillNode.first_child().name()) == "fgColor"); // fgColor MUST precede bgColor in sequence
    REQUIRE(std::string(fillNode.first_child().next_sibling().name()) == "bgColor");

    // Second DXF (Font booleans)
    auto dxf2 = styles.dxfs()[1].node();
    REQUIRE(std::string(dxf2.child("font").child("b").attribute("val").value()) == "1"); // Boolean must be "1" or "0" in this context

    doc.close();
}
