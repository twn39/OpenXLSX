#include <OpenXLSX.hpp>
#include <catch2/catch_test_macros.hpp>
#include <iostream>

using namespace OpenXLSX;

namespace
{
    // Hack to access protected member without inheritance (since XLDocument is final)
    template<typename Tag, typename Tag::type M>
    struct Rob
    {
        friend typename Tag::type get_impl(Tag) { return M; }
    };

    struct XLDocument_extractXmlFromArchive
    {
        typedef std::string (XLDocument::*type)(std::string_view);
    };

    template struct Rob<XLDocument_extractXmlFromArchive, &XLDocument::extractXmlFromArchive>;

    // Prototype declaration for the friend function
    std::string (XLDocument::* get_impl(XLDocument_extractXmlFromArchive))(std::string_view);

    // Function to call the protected member
    std::string getRawXml(XLDocument& doc, const std::string& path)
    {
        static auto fn = get_impl(XLDocument_extractXmlFromArchive());
        return (doc.*fn)(path);
    }
}    // namespace

TEST_CASE("XLRichText Tests", "[XLRichText]")
{
    SECTION("Create and Modify XLRichText")
    {
        XLRichText    rt;
        XLRichTextRun run1("Hello ");
        run1.setBold(true);
        run1.setFontColor(XLColor(255, 0, 0));

        XLRichTextRun run2("World!");
        run2.setItalic(true);
        run2.setFontSize(16);

        rt.addRun(run1);
        rt.addRun(run2);

        REQUIRE(rt.runs().size() == 2);
        REQUIRE(rt.plainText() == "Hello World!");
        REQUIRE(rt.runs()[0].bold() == true);
        REQUIRE(rt.runs()[1].italic() == true);
    }

    SECTION("Write and Read RichText in Cell")
    {
        XLDocument doc;
        doc.create("./RichTextTest.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        XLRichText    rt;
        XLRichTextRun run1("Red Bold ");
        run1.setBold(true);
        run1.setFontColor(XLColor(255, 0, 0));

        XLRichTextRun run2("Blue Italic");
        run2.setItalic(true);
        run2.setFontColor(XLColor(0, 0, 255));

        rt.addRun(run1);
        rt.addRun(run2);

        wks.cell("A1").value() = rt;
        doc.save();
        doc.close();

        XLDocument doc2;
        doc2.open("./RichTextTest.xlsx");
        auto wks2 = doc2.workbook().worksheet("Sheet1");

        XLCellValue val = wks2.cell("A1").value();
        REQUIRE(val.type() == XLValueType::RichText);

        OpenXLSX::XLRichText rt2 = val.get<OpenXLSX::XLRichText>();
        REQUIRE(rt2.runs().size() == 2);
        REQUIRE(rt2.plainText() == "Red Bold Blue Italic");

        REQUIRE(rt2.runs()[0].text() == "Red Bold ");
        REQUIRE(rt2.runs()[0].bold() == true);

        REQUIRE(rt2.runs()[1].text() == "Blue Italic");
        REQUIRE(rt2.runs()[1].italic() == true);

        doc2.close();
    }

    SECTION("OOXML Structure Verification")
    {
        XLDocument doc;
        doc.create("./structure_test.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        // Rich Text
        XLRichText rt;
        rt.addRun(XLRichTextRun("Bold"));
        rt.runs().back().setBold(true);
        wks.cell("A1").value() = rt;

        // AutoFilter
        wks.setAutoFilter(wks.range("A1", "B2"));

        // MergeCells (to check ordering)
        wks.merges().appendMerge("C1:D1");

        doc.save();

        // Use internal API to peek into the XML
        // We expect autoFilter to appear BEFORE mergeCells in the XML
        std::string xmlContent = getRawXml(doc, "xl/worksheets/sheet1.xml");

        doc.close();

        // 1. Check Rich Text structure: <is><r><rPr><b></b></rPr><t>Bold</t></r>
        REQUIRE(xmlContent.find("<is>") != std::string::npos);
        REQUIRE(xmlContent.find("<r>") != std::string::npos);
        REQUIRE(xmlContent.find("<rPr>") != std::string::npos);
        REQUIRE(xmlContent.find("<b />") != std::string::npos);    // pugi might use <b /> for empty element
        REQUIRE(xmlContent.find("<t>Bold</t>") != std::string::npos);

        // 2. Check AutoFilter structure
        REQUIRE(xmlContent.find("<autoFilter ref=\"A1:B2\"") != std::string::npos);

        // 3. Check node ordering: autoFilter should come BEFORE mergeCells
        size_t autoFilterPos = xmlContent.find("<autoFilter");
        size_t mergeCellsPos = xmlContent.find("<mergeCells");
        REQUIRE(autoFilterPos != std::string::npos);
        REQUIRE(mergeCellsPos != std::string::npos);
        REQUIRE(autoFilterPos < mergeCellsPos);
    }
}
