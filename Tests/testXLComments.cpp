#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>

using namespace OpenXLSX;

TEST_CASE("XLComments Tests", "[XLComments]")
{
    SECTION("Basic Comment Operations")
    {
        XLDocument doc;
        doc.create("./testXLComments.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        REQUIRE(wks.hasComments() == false);

        // Add a comment
        wks.comments().addAuthor("Author1");
        wks.comments().set("A1", "This is a comment", 0);

        REQUIRE(wks.hasComments() == true);
        REQUIRE(wks.comments().count() == 1);
        REQUIRE(wks.comments().get("A1") == "This is a comment");
        REQUIRE(wks.comments().authorCount() == 1);
        REQUIRE(wks.comments().author(0) == "Author1");

        // Delete comment
        wks.comments().deleteComment("A1");
        REQUIRE(wks.comments().count() == 0);
        REQUIRE(wks.comments().get("A1") == "");

        doc.close();
    }

    SECTION("Multiple Authors and Comments")
    {
        XLDocument doc;
        doc.create("./testXLCommentsMultiple.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        auto&    comments = wks.comments();
        uint16_t auth1    = comments.addAuthor("Author One");
        uint16_t auth2    = comments.addAuthor("Author Two");

        comments.set("B2", "Comment by Auth 1", auth1);
        comments.set("C3", "Comment by Auth 2", auth2);

        REQUIRE(comments.count() == 2);
        REQUIRE(comments.authorId("B2") == auth1);
        REQUIRE(comments.authorId("C3") == auth2);

        REQUIRE(comments.get(0).ref() == "B2");
        REQUIRE(comments.get(1).ref() == "C3");

        doc.close();
    }

    SECTION("Comment Shape Properties")
    {
        XLDocument doc;
        doc.create("./testXLCommentsShapes.xlsx", XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.comments().set("D4", "Shape test");
        auto shape = wks.comments().shape("D4");

        // Test shape visibility
        shape.style().show();
        REQUIRE(shape.style().visible() == true);
        shape.style().hide();
        REQUIRE(shape.style().visible() == false);

        doc.close();
    }
}
