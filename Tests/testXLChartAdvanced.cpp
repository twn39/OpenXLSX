#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <filesystem>
#include <fstream>

using namespace OpenXLSX;

namespace
{
    // Hack to access protected member without inheritance
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

    std::string (XLDocument::* get_impl(XLDocument_extractXmlFromArchive))(std::string_view);

    std::string getRawXml(XLDocument& doc, const std::string& path)
    {
        static auto fn = get_impl(XLDocument_extractXmlFromArchive());
        return (doc.*fn)(path);
    }
}

class XLChartAdvTestDoc
{
public:
    XLDocument  doc;
    void        open(const std::string& filename) { doc.open(filename); }
    void        close() { doc.close(); }
    std::string getRawXml(const std::string& path) { return ::getRawXml(doc, path); }
};

TEST_CASE("Advanced Chart Visual Elements", "[XLChart][OOXML]")
{
    std::string filename = "test_chart_advanced.xlsx";

    SECTION("Chart Title and Rich Text Format")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            auto chart = wks.addChart(XLChartType::Bar, "Bar", 1, 1, 400, 300);
            
            chart.setTitle("Strategic Overview");
            
            doc.save();
            doc.close();
        }

        {
            XLChartAdvTestDoc testDoc;
            testDoc.open(filename);
            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

            // Verify title node architecture
            REQUIRE(chartXml.find("<c:title>") != std::string::npos);
            REQUIRE(chartXml.find("<c:tx>") != std::string::npos);
            REQUIRE(chartXml.find("<c:rich>") != std::string::npos);
            REQUIRE(chartXml.find("<a:t>Strategic Overview</a:t>") != std::string::npos);
            REQUIRE(chartXml.find("<c:overlay val=\"0\"") != std::string::npos);
            
            testDoc.close();
        }
        std::filesystem::remove(filename);
    }
    
    SECTION("Legend Positions")
    {
        std::map<XLLegendPosition, std::string> posMap = {
            {XLLegendPosition::Bottom, "b"},
            {XLLegendPosition::Left, "l"},
            {XLLegendPosition::Right, "r"},
            {XLLegendPosition::Top, "t"},
            {XLLegendPosition::TopRight, "tr"}
        };

        for (auto const& [enumPos, xmlVal] : posMap) {
            {
                XLDocument doc;
                doc.create(filename, XLForceOverwrite);
                auto wks = doc.workbook().worksheet("Sheet1");
                auto chart = wks.addChart(XLChartType::Line, "Line", 1, 1, 400, 300);
                
                chart.setLegendPosition(enumPos);
                
                doc.save();
                doc.close();
            }

            {
                XLChartAdvTestDoc testDoc;
                testDoc.open(filename);
                std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

                std::string expectedStr = "<c:legendPos val=\"" + xmlVal + "\"";
                REQUIRE(chartXml.find(expectedStr) != std::string::npos);
                
                testDoc.close();
            }
            std::filesystem::remove(filename);
        }
    }
    
    SECTION("Hide Legend Functionality")
    {
        {
            XLDocument doc;
            doc.create(filename, XLForceOverwrite);
            auto wks = doc.workbook().worksheet("Sheet1");
            auto chart = wks.addChart(XLChartType::Pie, "Pie", 1, 1, 400, 300);
            
            chart.setLegendPosition(XLLegendPosition::Hidden);
            
            doc.save();
            doc.close();
        }

        {
            XLChartAdvTestDoc testDoc;
            testDoc.open(filename);
            std::string chartXml = testDoc.getRawXml("xl/charts/chart1.xml");

            // Ensure the legend node does not exist at all
            REQUIRE(chartXml.find("<c:legend>") == std::string::npos);
            REQUIRE(chartXml.find("<c:legendPos") == std::string::npos);
            
            testDoc.close();
        }
        std::filesystem::remove(filename);
    }
}
