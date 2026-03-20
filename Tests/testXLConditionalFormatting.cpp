#include <catch2/catch_test_macros.hpp>
#include <OpenXLSX.hpp>
#include <iostream>

using namespace OpenXLSX;

TEST_CASE("XLConditionalFormatting Tests", "[ConditionalFormatting]") {
    SECTION("XLCfvo Property Tests") {
        XLCfvo cfvo;
        cfvo.setType(XLCfvoType::Percent);
        cfvo.setValue("50");
        cfvo.setGte(false);

        REQUIRE(cfvo.type() == XLCfvoType::Percent);
        REQUIRE(cfvo.value() == "50");
        REQUIRE(cfvo.gte() == false);
    }

    SECTION("XLCfDataBar Property Tests") {
        XLCfDataBar dataBar;
        dataBar.setMin(XLCfvoType::Min, "0");
        dataBar.setMax(XLCfvoType::Max, "0");
        dataBar.setColor(XLColor(255, 0, 0));
        dataBar.setShowValue(false);

        REQUIRE(dataBar.min().type() == XLCfvoType::Min);
        REQUIRE(dataBar.max().type() == XLCfvoType::Max);
        REQUIRE(dataBar.color() == XLColor(255, 0, 0));
        REQUIRE(dataBar.showValue() == false);
    }

    SECTION("XLCfColorScale Property Tests") {
        XLCfColorScale colorScale;
        colorScale.addValue(XLCfvoType::Min, "0", XLColor(255, 0, 0));
        colorScale.addValue(XLCfvoType::Max, "0", XLColor(0, 255, 0));

        auto cfvos = colorScale.cfvos();
        auto colors = colorScale.colors();

        REQUIRE(cfvos.size() == 2);
        REQUIRE(colors.size() == 2);
        REQUIRE(cfvos[0].type() == XLCfvoType::Min);
        REQUIRE(colors[0] == XLColor(255, 0, 0));
        REQUIRE(colors[1] == XLColor(0, 255, 0));
    }

    SECTION("XLCfIconSet Property Tests") {
        XLCfIconSet iconSet;
        iconSet.setIconSet("3Arrows");
        iconSet.addValue(XLCfvoType::Percent, "0");
        iconSet.addValue(XLCfvoType::Percent, "33");
        iconSet.setShowValue(false);
        iconSet.setReverse(true);

        REQUIRE(iconSet.iconSet() == "3Arrows");
        REQUIRE(iconSet.cfvos().size() == 2);
        REQUIRE(iconSet.showValue() == false);
        REQUIRE(iconSet.reverse() == true);
    }

    SECTION("Worksheet Integration and OOXML Validation") {
        XLDocument doc;
        doc.create("CFTest.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        auto cfIdx = wks.conditionalFormats().create();
        auto cf = wks.conditionalFormats().conditionalFormatByIndex(cfIdx);
        cf.setSqref("A1:A10");

        auto ruleIdx = cf.cfRules().create();
        auto rule = cf.cfRules().cfRuleByIndex(ruleIdx);
        rule.setType(XLCfType::DataBar);

        XLCfDataBar db;
        db.setColor(XLColor(0, 0, 255));
        db.setMin(XLCfvoType::Number, "10");
        db.setMax(XLCfvoType::Number, "90");
        rule.setDataBar(db);

        // Check if XML is correctly generated before saving
        XMLNode dbNode = rule.dataBar().node();
        REQUIRE_FALSE(dbNode.empty());
        REQUIRE(std::string(dbNode.name()) == "dataBar");
        
        auto cfvoNodes = dbNode.children("cfvo");
        int count = 0;
        for (auto& n : cfvoNodes) {
            if (count == 0) {
                REQUIRE(std::string(n.attribute("type").value()) == "num");
                REQUIRE(std::string(n.attribute("val").value()) == "10");
            }
            count++;
        }
        REQUIRE(count == 2);
        REQUIRE(std::string(dbNode.child("color").attribute("rgb").value()) == "ff0000ff");

        doc.save();
        doc.close();
    }
}
