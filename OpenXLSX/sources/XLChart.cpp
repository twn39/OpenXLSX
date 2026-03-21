// ===== External Includes ===== //
#include <pugixml.hpp>
#include <string>
#include <string_view>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLChart.hpp"

using namespace OpenXLSX;

namespace OpenXLSX
{

    XLChart::XLChart(XLXmlData* xmlData) : XLXmlFile(xmlData)
    {
        if (xmlData && xmlData->getXmlType() != XLContentType::Chart) {
            throw XLInternalError("XLChart constructor: Invalid XML data.");
        }
    }

    void XLChart::initXml(XLChartType type)
    {
        if (getXmlPath().empty()) return;
        XMLDocument& doc = xmlDocument();
        if (doc.document_element().empty()) {
            constexpr std::string_view baseTemplate = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:plotArea></c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:layout/>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>)";

            doc.load_string(baseTemplate.data(), pugi_parse_settings);
            
            XMLNode plotArea = doc.document_element().child("c:chart").child("c:plotArea");
            bool hasAxes = true;
            XMLNode chartNode;
            
            switch (type) {
                case XLChartType::Line:
                    chartNode = plotArea.append_child("c:lineChart");
                    chartNode.append_child("c:grouping").append_attribute("val").set_value("standard");
                    break;
                case XLChartType::Pie:
                    chartNode = plotArea.append_child("c:pieChart");
                    chartNode.append_child("c:varyColors").append_attribute("val").set_value("1");
                    hasAxes = false;
                    break;
                case XLChartType::Scatter:
                    chartNode = plotArea.append_child("c:scatterChart");
                    chartNode.append_child("c:scatterStyle").append_attribute("val").set_value("lineMarker");
                    chartNode.append_child("c:varyColors").append_attribute("val").set_value("0");
                    break;
                case XLChartType::Area:
                    chartNode = plotArea.append_child("c:areaChart");
                    chartNode.append_child("c:grouping").append_attribute("val").set_value("standard");
                    break;
                case XLChartType::Doughnut:
                    chartNode = plotArea.append_child("c:doughnutChart");
                    chartNode.append_child("c:varyColors").append_attribute("val").set_value("1");
                    chartNode.append_child("c:holeSize").append_attribute("val").set_value("75");
                    hasAxes = false;
                    break;
                case XLChartType::Bar:
                default:
                    chartNode = plotArea.append_child("c:barChart");
                    chartNode.append_child("c:barDir").append_attribute("val").set_value("col");
                    chartNode.append_child("c:grouping").append_attribute("val").set_value("clustered");
                    break;
            }

            if (hasAxes) {
                chartNode.append_child("c:axId").append_attribute("val").set_value("100000000");
                chartNode.append_child("c:axId").append_attribute("val").set_value("100000001");
                
                std::string axesTemplate;
                if (type == XLChartType::Scatter) {
                    axesTemplate = R"(<dummy>
      <c:valAx>
        <c:axId val="100000000"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="General" sourceLinked="0"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="100000001"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:valAx>
      <c:valAx>
        <c:axId val="100000001"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:numFmt formatCode="General" sourceLinked="0"/>
        <c:majorGridlines/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="100000000"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
</dummy>)";
                } else {
                    axesTemplate = R"(<dummy>
      <c:catAx>
        <c:axId val="100000000"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="General" sourceLinked="0"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="100000001"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="100000001"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:numFmt formatCode="General" sourceLinked="0"/>
        <c:majorGridlines/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="100000000"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
</dummy>)";
                }
                XMLDocument axesDoc;
                axesDoc.load_string(axesTemplate.data(), pugi_parse_settings);
                for (auto child : axesDoc.child("dummy").children()) {
                    plotArea.append_copy(child);
                }
            }
        }
    }

    XMLNode getChartNode(const XMLDocument& doc) {
        XMLNode plotArea = doc.document_element().child("c:chart").child("c:plotArea");
        for (auto child : plotArea.children()) {
            std::string_view name = child.name();
            if (name.length() > 5 && name.substr(name.length() - 5) == "Chart") {
                return child;
            }
        }
        return XMLNode();
    }

    uint32_t XLChart::seriesCount() const
    {
        uint32_t count = 0;
        XMLNode chartNode = getChartNode(xmlDocument());
        for ([[maybe_unused]] auto child : chartNode.children("c:ser")) {
            count++;
        }
        return count;
    }

    void XLChart::addSeries(std::string_view valuesRef, std::string_view title, std::string_view categoriesRef)
    {
        XMLNode chartNode = getChartNode(xmlDocument());
        if (chartNode.empty()) return;

        const uint32_t idx = seriesCount();

        XMLNode serNode = chartNode.append_child("c:ser");
        serNode.append_child("c:idx").append_attribute("val").set_value(idx);
        serNode.append_child("c:order").append_attribute("val").set_value(idx);

        // 1. Series Title (c:tx)
        if (!title.empty()) {
            XMLNode txNode = serNode.append_child("c:tx");
            if (title.find('!') != std::string_view::npos) {
                txNode.append_child("c:strRef").append_child("c:f").text().set(std::string(title).c_str());
            } else {
                txNode.append_child("c:v").text().set(std::string(title).c_str());
            }
        }

        std::string_view chartType = chartNode.name();

        // 2. Categories (X-Axis) and Values (Y-Axis)
        if (chartType == "c:scatterChart") {
            if (!categoriesRef.empty()) {
                XMLNode xValNode = serNode.append_child("c:xVal");
                XMLNode numRefNode = xValNode.append_child("c:numRef");
                numRefNode.append_child("c:f").text().set(std::string(categoriesRef).c_str());
            }
            XMLNode yValNode = serNode.append_child("c:yVal");
            XMLNode numRefNode = yValNode.append_child("c:numRef");
            numRefNode.append_child("c:f").text().set(std::string(valuesRef).c_str());
        } else {
            if (!categoriesRef.empty()) {
                XMLNode catNode = serNode.append_child("c:cat");
                XMLNode strRefNode = catNode.append_child("c:strRef");
                strRefNode.append_child("c:f").text().set(std::string(categoriesRef).c_str());
            }
            XMLNode valNode = serNode.append_child("c:val");
            XMLNode numRefNode = valNode.append_child("c:numRef");
            numRefNode.append_child("c:f").text().set(std::string(valuesRef).c_str());
        }
    }

    void XLChart::setTitle(std::string_view title)
    {
        if (title.empty()) return;
        XMLNode chartNode = xmlDocument().document_element().child("c:chart");
        
        // Remove existing title if any
        chartNode.remove_child("c:title");

        // Insert title at the top of the chart node
        XMLNode titleNode = chartNode.insert_child_before("c:title", chartNode.first_child());
        XMLNode txNode = titleNode.append_child("c:tx");
        XMLNode richNode = txNode.append_child("c:rich");
        
        richNode.append_child("a:bodyPr");
        richNode.append_child("a:lstStyle");
        
        XMLNode pNode = richNode.append_child("a:p");
        XMLNode rNode = pNode.append_child("a:r");
        rNode.append_child("a:t").text().set(std::string(title).c_str());
        
        titleNode.append_child("c:overlay").append_attribute("val").set_value("0");
    }

    void XLChart::setLegendPosition(XLLegendPosition position)
    {
        XMLNode chartNode = xmlDocument().document_element().child("c:chart");
        XMLNode legendNode = chartNode.child("c:legend");
        
        if (position == XLLegendPosition::Hidden) {
            chartNode.remove_child(legendNode);
            return;
        }

        if (legendNode.empty()) {
            legendNode = chartNode.insert_child_before("c:legend", chartNode.child("c:plotVisOnly"));
        }

        XMLNode posNode = legendNode.child("c:legendPos");
        if (posNode.empty()) {
            posNode = legendNode.insert_child_before("c:legendPos", legendNode.first_child());
        }

        switch (position) {
            case XLLegendPosition::Bottom:   posNode.attribute("val").set_value("b"); break;
            case XLLegendPosition::Left:     posNode.attribute("val").set_value("l"); break;
            case XLLegendPosition::Right:    posNode.attribute("val").set_value("r"); break;
            case XLLegendPosition::Top:      posNode.attribute("val").set_value("t"); break;
            case XLLegendPosition::TopRight: posNode.attribute("val").set_value("tr"); break;
            default: break;
        }
        
        if (legendNode.child("c:overlay").empty()) {
            legendNode.append_child("c:overlay").append_attribute("val").set_value("0");
        }
    }

} // namespace OpenXLSX

    void XLChart::setShowDataLabels(bool showValue, bool showCategory, bool showPercent)
    {
        XMLNode chartNode = getChartNode(xmlDocument());
        if (chartNode.empty()) return;

        XMLNode dLbls = chartNode.child("c:dLbls");
        if (dLbls.empty()) {
            XMLNode insertBeforeNode;
            for (auto child : chartNode.children()) {
                std::string_view name = child.name();
                if (name != "c:barDir" && name != "c:grouping" && name != "c:scatterStyle" && name != "c:varyColors" && name != "c:ser") {
                    insertBeforeNode = child;
                    break;
                }
            }

            if (!insertBeforeNode.empty()) {
                dLbls = chartNode.insert_child_before("c:dLbls", insertBeforeNode);
            } else {
                dLbls = chartNode.append_child("c:dLbls");
            }
        } else {
            // Clear existing data labels to maintain exact OOXML sequence
            dLbls.remove_children();
        }

        // Must follow strict OOXML sequence: showLegendKey, showVal, showCatName, showSerName, showPercent, showBubbleSize, showLeaderLines
        dLbls.append_child("c:showLegendKey").append_attribute("val").set_value("0");
        dLbls.append_child("c:showVal").append_attribute("val").set_value(showValue ? "1" : "0");
        dLbls.append_child("c:showCatName").append_attribute("val").set_value(showCategory ? "1" : "0");
        dLbls.append_child("c:showSerName").append_attribute("val").set_value("0");
        dLbls.append_child("c:showPercent").append_attribute("val").set_value(showPercent ? "1" : "0");
        dLbls.append_child("c:showBubbleSize").append_attribute("val").set_value("0");
        dLbls.append_child("c:showLeaderLines").append_attribute("val").set_value("1");
    }
