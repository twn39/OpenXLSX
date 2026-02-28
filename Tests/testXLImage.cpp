#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>
#include <pugixml.hpp>
#include <vector>

using namespace OpenXLSX;

TEST_CASE("XLDocument Image Structure and Insertion", "[XLDocument]")
{
    SECTION("Comprehensive structural check")
    {
        XLDocument doc;
        doc.create("./testImageStruct.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        // Read image from file Tests/test.jpg
        std::ifstream file("Tests/test.jpg", std::ios::binary | std::ios::ate);
        REQUIRE(file.is_open());
        std::streamsize size = file.tellg();
        file.seekg(0, std::ios::beg);
        std::string imageData(size, '\0');
        file.read(&imageData[0], size);

        // 1. Add image
        std::string imagePath = doc.addImage("test.jpg", imageData);

        // 2. Get drawing
        auto&       drawing         = wks.drawing();
        std::string drawingPath     = "xl/drawings/drawing1.xml";
        std::string drawingRelsPath = "xl/drawings/_rels/drawing1.xml.rels";

        // 3. Link image to drawing
        std::string relPath = "../media/test.jpg";
        auto        rel     = drawing.relationships().addRelationship(XLRelationshipType::Image, relPath);
        drawing.addImage(rel.id(), "test.jpg", "Test Image", 1, 1, 400, 300);

        doc.save();
        doc.close();

        // 4. Re-open and perform structural verification
        XLDocument doc2;
        doc2.open("./testImageStruct.xlsx");

        // A. Check image file existence and integrity
        std::string readData = doc2.getImage("xl/media/test.jpg");
        REQUIRE(readData.size() == imageData.size());
        REQUIRE(readData == imageData);

        // B. Check Worksheet XML structure
        auto wks2 = doc2.workbook().worksheet("Sheet1");
        REQUIRE(wks2.hasDrawing());

        pugi::xml_document sheetDoc;
        sheetDoc.load_string(wks2.xmlData().c_str());
        auto worksheet = sheetDoc.child("worksheet");
        REQUIRE_FALSE(worksheet.empty());
        REQUIRE(worksheet.attribute("xmlns:xdr").empty());    // Should not have redundant namespace

        auto drawingNode = worksheet.child("drawing");
        REQUIRE_FALSE(drawingNode.empty());
        REQUIRE_FALSE(drawingNode.attribute("r:id").empty());

        // C. Check Drawing XML structure
        auto& drawing2 = wks2.drawing();
        REQUIRE(drawing2.imageCount() == 1);

        pugi::xml_document drawingDoc;
        drawingDoc.load_string(drawing2.xmlData().c_str());
        auto wsDr = drawingDoc.child("xdr:wsDr");
        REQUIRE_FALSE(wsDr.empty());
        REQUIRE_FALSE(wsDr.attribute("xmlns:r").empty());

        auto anchor = wsDr.child("xdr:oneCellAnchor");
        REQUIRE_FALSE(anchor.empty());
        REQUIRE_FALSE(anchor.child("xdr:pic").empty());

        // D. Check Content Types XML structure
        auto& ct = doc2.contentTypes();
        REQUIRE(ct.hasDefault("jpg"));

        pugi::xml_document ctDoc;
        ctDoc.load_string(ct.xmlData().c_str());
        auto types = ctDoc.child("Types");
        REQUIRE_FALSE(types.empty());

        // Ensure jpg default is before overrides
        bool foundJpg      = false;
        bool foundOverride = false;
        for (auto node : types.children()) {
            if (std::string(node.name()) == "Default") {
                if (std::string(node.attribute("Extension").value()) == "jpg") {
                    foundJpg = true;
                    REQUIRE_FALSE(foundOverride);    // Default should be before Override
                }
            }
            else if (std::string(node.name()) == "Override") {
                foundOverride = true;
            }
        }
        REQUIRE(foundJpg);

        // E. Check for standalone="yes" in declarations
        REQUIRE(wks2.xmlData().find("standalone=\"yes\"") != std::string::npos);
        REQUIRE(drawing2.xmlData().find("standalone=\"yes\"") != std::string::npos);
        REQUIRE(ct.xmlData().find("standalone=\"yes\"") != std::string::npos);

        doc2.close();
    }

    SECTION("Image dimension detection and scaling")
    {
        // 1. Test PNG dimension detection (test.png is 1440x446)
        std::ifstream pngFile("Tests/test.png", std::ios::binary | std::ios::ate);
        REQUIRE(pngFile.is_open());
        std::streamsize pngSize = pngFile.tellg();
        pngFile.seekg(0, std::ios::beg);
        std::string pngData(pngSize, '\0');
        pngFile.read(&pngData[0], pngSize);

        auto [pngW, pngH] = getImageDimensions(pngData);
        REQUIRE(pngW == 1440);
        REQUIRE(pngH == 446);

        // 2. Test JPEG dimension detection
        std::ifstream jpgFile("Tests/test.jpg", std::ios::binary | std::ios::ate);
        REQUIRE(jpgFile.is_open());
        std::streamsize jpgSize = jpgFile.tellg();
        jpgFile.seekg(0, std::ios::beg);
        std::string jpgData(jpgSize, '\0');
        jpgFile.read(&jpgData[0], jpgSize);

        auto [jpgW, jpgH] = getImageDimensions(jpgData);
        REQUIRE(jpgW > 0);
        REQUIRE(jpgH > 0);

        // 3. Test addScaledImage
        XLDocument doc;
        doc.create("./testImageScaling.xlsx", XLForceOverwrite);
        auto  wks     = doc.workbook().worksheet("Sheet1");
        auto& drawing = wks.drawing();

        std::string imagePath = doc.addImage("scaled_test.png", pngData);
        std::string relPath   = "../media/scaled_test.png";
        auto        rel       = drawing.relationships().addRelationship(XLRelationshipType::Image, relPath);

        // Scale to 25% (1440 * 0.25 = 360, 446 * 0.25 = 111.5 -> 111)
        drawing.addScaledImage(rel.id(), "scaled_test.png", "Scaled PNG", pngData, 1, 1, 0.25);

        REQUIRE(drawing.imageCount() == 1);
        auto img = drawing.image(0);
        REQUIRE(img.width() == 360);
        REQUIRE(img.height() == 111);

        doc.save();
        doc.close();
    }
}
