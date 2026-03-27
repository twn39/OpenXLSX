#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <cmath>

using namespace OpenXLSX;

// Helper: create a no-cell resolver (for pure arithmetic tests)
static const XLCellResolver noResolver{};

// Helper: return a resolver that maps refs to fixed values
static XLCellResolver makeMapResolver(std::initializer_list<std::pair<std::string, XLCellValue>> init)
{
    auto m = std::make_shared<std::unordered_map<std::string, XLCellValue>>();
    for (auto& p : init) m->insert_or_assign(p.first, p.second);
    return [m](std::string_view ref) -> XLCellValue {
        auto it = m->find(std::string(ref));
        return it != m->end() ? it->second : XLCellValue{};
    };
}

TEST_CASE("XLFormulaEngine – Lexer", "[XLFormulaEngine]")
{
    SECTION("Numbers") {
        auto toks = XLFormulaLexer::tokenize("=3.14");
        REQUIRE(toks[0].kind == XLTokenKind::Number);
        REQUIRE(toks[0].number == Catch::Approx(3.14));
    }

    SECTION("String literal") {
        auto toks = XLFormulaLexer::tokenize("=\"hello\"");
        REQUIRE(toks[0].kind == XLTokenKind::String);
        REQUIRE(toks[0].text == "hello");
    }

    SECTION("Bool") {
        auto toks = XLFormulaLexer::tokenize("=TRUE");
        REQUIRE(toks[0].kind == XLTokenKind::Bool);
        REQUIRE(toks[0].boolean == true);
    }

    SECTION("Cell ref") {
        auto toks = XLFormulaLexer::tokenize("=A1");
        REQUIRE(toks[0].kind == XLTokenKind::CellRef);
        REQUIRE(toks[0].text == "A1");
    }

    SECTION("Range ref") {
        auto toks = XLFormulaLexer::tokenize("=A1:C3");
        REQUIRE(toks[0].kind == XLTokenKind::CellRef);
        REQUIRE(toks[0].text == "A1:C3");
    }

    SECTION("Operators") {
        auto toks = XLFormulaLexer::tokenize("=<>=");
        REQUIRE(toks[0].kind == XLTokenKind::NEq);
        REQUIRE(toks[1].kind == XLTokenKind::Eq);
    }
}

TEST_CASE("XLFormulaEngine – Arithmetic", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    SECTION("Addition")         { REQUIRE(eng.evaluate("=1+2").get<double>() == Catch::Approx(3.0)); }
    SECTION("Multiplication")   { REQUIRE(eng.evaluate("=2*3").get<double>() == Catch::Approx(6.0)); }
    SECTION("Precedence")       { REQUIRE(eng.evaluate("=1+2*3").get<double>() == Catch::Approx(7.0)); }
    SECTION("Parentheses")      { REQUIRE(eng.evaluate("=(1+2)*3").get<double>() == Catch::Approx(9.0)); }
    SECTION("Subtraction")      { REQUIRE(eng.evaluate("=10-4").get<double>() == Catch::Approx(6.0)); }
    SECTION("Division")         { REQUIRE(eng.evaluate("=10/4").get<double>() == Catch::Approx(2.5)); }
    SECTION("Power")            { REQUIRE(eng.evaluate("=2^10").get<double>() == Catch::Approx(1024.0)); }
    SECTION("Unary minus")      { REQUIRE(eng.evaluate("=-5").get<double>() == Catch::Approx(-5.0)); }
    SECTION("Percent")          { REQUIRE(eng.evaluate("=50%").get<double>() == Catch::Approx(0.5)); }
    SECTION("Div by zero")      { REQUIRE(eng.evaluate("=1/0").type() == XLValueType::Error); }
    SECTION("No leading =")     { REQUIRE(eng.evaluate("1+1").get<double>() == Catch::Approx(2.0)); }
}

TEST_CASE("XLFormulaEngine – Comparison", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    SECTION("Equal numeric")       { REQUIRE(eng.evaluate("=1=1").get<bool>() == true); }
    SECTION("Not equal")           { REQUIRE(eng.evaluate("=1<>2").get<bool>() == true); }
    SECTION("Less than")           { REQUIRE(eng.evaluate("=1<2").get<bool>() == true); }
    SECTION("Greater equal")       { REQUIRE(eng.evaluate("=2>=2").get<bool>() == true); }
    SECTION("String equal (CI)")   { REQUIRE(eng.evaluate("=\"A\"=\"a\"").get<bool>() == true); }
    SECTION("String less")         { REQUIRE(eng.evaluate("=\"abc\"<\"abd\"").get<bool>() == true); }
}

TEST_CASE("XLFormulaEngine – String concat", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;
    SECTION("Ampersand") {
        auto r = eng.evaluate("=\"Hello\"&\" \"&\"World\"");
        REQUIRE(r.get<std::string>() == "Hello World");
    }
}

TEST_CASE("XLFormulaEngine – Cell refs", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;
    auto resolver = makeMapResolver({ {"A1", XLCellValue(10.0)}, {"B1", XLCellValue(5.0)} });

    SECTION("Single ref")   { REQUIRE(eng.evaluate("=A1", resolver).get<double>() == Catch::Approx(10.0)); }
    SECTION("Two refs")     { REQUIRE(eng.evaluate("=A1+B1", resolver).get<double>() == Catch::Approx(15.0)); }
    SECTION("Ref * literal"){ REQUIRE(eng.evaluate("=A1*2", resolver).get<double>() == Catch::Approx(20.0)); }
}

TEST_CASE("XLFormulaEngine – Range functions", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;
    // A1=1, B1=2, C1=3
    auto resolver = makeMapResolver({
        {"A1", XLCellValue(1.0)}, {"B1", XLCellValue(2.0)}, {"C1", XLCellValue(3.0)},
        {"A2", XLCellValue(4.0)}, {"B2", XLCellValue(5.0)}, {"C2", XLCellValue(6.0)},
    });

    SECTION("SUM")     { REQUIRE(eng.evaluate("=SUM(A1:C1)", resolver).get<double>() == Catch::Approx(6.0)); }
    SECTION("AVERAGE") { REQUIRE(eng.evaluate("=AVERAGE(A1:C1)", resolver).get<double>() == Catch::Approx(2.0)); }
    SECTION("MIN")     { REQUIRE(eng.evaluate("=MIN(A1:C1)", resolver).get<double>() == Catch::Approx(1.0)); }
    SECTION("MAX")     { REQUIRE(eng.evaluate("=MAX(A1:C1)", resolver).get<double>() == Catch::Approx(3.0)); }
    SECTION("COUNT")   { REQUIRE(eng.evaluate("=COUNT(A1:C1)", resolver).get<int64_t>() == 3); }
    SECTION("SUM 2D")  { REQUIRE(eng.evaluate("=SUM(A1:C2)", resolver).get<double>() == Catch::Approx(21.0)); }

    SECTION("SUM multiple args") {
        REQUIRE(eng.evaluate("=SUM(A1,B1,C1)", resolver).get<double>() == Catch::Approx(6.0));
    }
}

TEST_CASE("XLFormulaEngine – Math functions", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    SECTION("ABS negative") { REQUIRE(eng.evaluate("=ABS(-5)").get<double>() == Catch::Approx(5.0)); }
    SECTION("ABS positive") { REQUIRE(eng.evaluate("=ABS(5)").get<double>() == Catch::Approx(5.0)); }
    SECTION("SQRT")         { REQUIRE(eng.evaluate("=SQRT(9)").get<double>() == Catch::Approx(3.0)); }
    SECTION("SQRT neg")     { REQUIRE(eng.evaluate("=SQRT(-1)").type() == XLValueType::Error); }
    SECTION("INT")          { REQUIRE(eng.evaluate("=INT(3.9)").get<int64_t>() == 3); }
    SECTION("MOD")          { REQUIRE(eng.evaluate("=MOD(10,3)").get<double>() == Catch::Approx(1.0)); }
    SECTION("POWER")        { REQUIRE(eng.evaluate("=POWER(2,8)").get<double>() == Catch::Approx(256.0)); }
    SECTION("ROUND")        { REQUIRE(eng.evaluate("=ROUND(3.567,2)").get<double>() == Catch::Approx(3.57)); }
    SECTION("ROUNDUP")      { REQUIRE(eng.evaluate("=ROUNDUP(3.111,2)").get<double>() == Catch::Approx(3.12)); }
    SECTION("ROUNDDOWN")    { REQUIRE(eng.evaluate("=ROUNDDOWN(3.999,2)").get<double>() == Catch::Approx(3.99)); }
}

TEST_CASE("XLFormulaEngine – Logical functions", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    SECTION("IF true")   { REQUIRE(eng.evaluate("=IF(1>0,\"yes\",\"no\")").get<std::string>() == "yes"); }
    SECTION("IF false")  { REQUIRE(eng.evaluate("=IF(0,\"yes\",\"no\")").get<std::string>() == "no"); }
    SECTION("AND true")  { REQUIRE(eng.evaluate("=AND(1,1,1)").get<bool>() == true); }
    SECTION("AND false") { REQUIRE(eng.evaluate("=AND(1,0,1)").get<bool>() == false); }
    SECTION("OR true")   { REQUIRE(eng.evaluate("=OR(0,0,1)").get<bool>() == true); }
    SECTION("OR false")  { REQUIRE(eng.evaluate("=OR(0,0,0)").get<bool>() == false); }
    SECTION("NOT")       { REQUIRE(eng.evaluate("=NOT(FALSE)").get<bool>() == true); }
    SECTION("IFERROR ok")  { REQUIRE(eng.evaluate("=IFERROR(1+1,0)").get<double>() == Catch::Approx(2.0)); }
    SECTION("IFERROR err") { REQUIRE(eng.evaluate("=IFERROR(1/0,99)").get<double>() == Catch::Approx(99.0)); }
}

TEST_CASE("XLFormulaEngine – Text functions", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;

    SECTION("LEN")         { REQUIRE(eng.evaluate("=LEN(\"hello\")").get<int64_t>() == 5); }
    SECTION("LEFT")        { REQUIRE(eng.evaluate("=LEFT(\"Excel\",3)").get<std::string>() == "Exc"); }
    SECTION("RIGHT")       { REQUIRE(eng.evaluate("=RIGHT(\"Excel\",2)").get<std::string>() == "el"); }
    SECTION("MID")         { REQUIRE(eng.evaluate("=MID(\"OpenXLSX\",5,4)").get<std::string>() == "XLSX"); }
    SECTION("UPPER")       { REQUIRE(eng.evaluate("=UPPER(\"hello\")").get<std::string>() == "HELLO"); }
    SECTION("LOWER")       { REQUIRE(eng.evaluate("=LOWER(\"HELLO\")").get<std::string>() == "hello"); }
    SECTION("TRIM")        { REQUIRE(eng.evaluate("=TRIM(\"  hi  \")").get<std::string>() == "hi"); }
    SECTION("CONCATENATE") { REQUIRE(eng.evaluate("=CONCATENATE(\"A\",\"B\",\"C\")").get<std::string>() == "ABC"); }
}

TEST_CASE("XLFormulaEngine – Info functions", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;
    auto resolver = makeMapResolver({ {"A1", XLCellValue(42.0)}, {"B1", XLCellValue("hello")}, {"C1", XLCellValue{}} });

    SECTION("ISNUMBER true")  { REQUIRE(eng.evaluate("=ISNUMBER(A1)", resolver).get<bool>() == true); }
    SECTION("ISNUMBER false") { REQUIRE(eng.evaluate("=ISNUMBER(B1)", resolver).get<bool>() == false); }
    SECTION("ISBLANK true")   { REQUIRE(eng.evaluate("=ISBLANK(C1)", resolver).get<bool>() == true); }
    SECTION("ISBLANK false")  { REQUIRE(eng.evaluate("=ISBLANK(A1)", resolver).get<bool>() == false); }
    SECTION("ISERROR true")   { REQUIRE(eng.evaluate("=ISERROR(1/0)").get<bool>() == true); }
    SECTION("ISERROR false")  { REQUIRE(eng.evaluate("=ISERROR(1+1)").get<bool>() == false); }
    SECTION("ISTEXT true")    { REQUIRE(eng.evaluate("=ISTEXT(B1)", resolver).get<bool>() == true); }
}

TEST_CASE("XLFormulaEngine – MATCH", "[XLFormulaEngine]")
{
    XLFormulaEngine eng;
    // Array: A1=10, B1=20, C1=30
    auto resolver = makeMapResolver({
        {"A1", XLCellValue(10.0)}, {"B1", XLCellValue(20.0)}, {"C1", XLCellValue(30.0)}
    });
    SECTION("MATCH exact") {
        REQUIRE(eng.evaluate("=MATCH(20,A1:C1,0)", resolver).get<int64_t>() == 2);
    }
    SECTION("MATCH not found") {
        REQUIRE(eng.evaluate("=MATCH(99,A1:C1,0)", resolver).type() == XLValueType::Error);
    }
}

TEST_CASE("XLFormulaEngine – VLOOKUP", "[XLFormulaEngine]")
{
    //  Table (A1:B3):  1,"apple"  / 2,"banana"  / 3,"cherry"
    //  colIdx=2, exact match
    XLFormulaEngine eng;
    auto resolver = makeMapResolver({
        {"A1", XLCellValue(1.0)}, {"B1", XLCellValue("apple")},
        {"A2", XLCellValue(2.0)}, {"B2", XLCellValue("banana")},
        {"A3", XLCellValue(3.0)}, {"B3", XLCellValue("cherry")},
    });
    SECTION("VLOOKUP found") {
        REQUIRE(eng.evaluate("=VLOOKUP(2,A1:B3,2,0)", resolver).get<std::string>() == "banana");
    }
    SECTION("VLOOKUP not found") {
        REQUIRE(eng.evaluate("=VLOOKUP(99,A1:B3,2,0)", resolver).type() == XLValueType::Error);
    }
}

TEST_CASE("XLFormulaEngine – Integration with XLDocument", "[XLFormulaEngine]")
{
    XLDocument doc;
    doc.create("./testXLFormulaEngine_integration.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = 10.0;
    wks.cell("B1").value() = 20.0;
    wks.cell("C1").value() = 30.0;
    wks.cell("A2").value() = std::string("hello");

    XLFormulaEngine eng;
    auto resolver = XLFormulaEngine::makeResolver(wks);

    SECTION("SUM via worksheet") {
        REQUIRE(eng.evaluate("=SUM(A1:C1)", resolver).get<double>() == Catch::Approx(60.0));
    }
    SECTION("AVERAGE via worksheet") {
        REQUIRE(eng.evaluate("=AVERAGE(A1:C1)", resolver).get<double>() == Catch::Approx(20.0));
    }
    SECTION("Cell arithmetic") {
        REQUIRE(eng.evaluate("=A1*B1", resolver).get<double>() == Catch::Approx(200.0));
    }
    SECTION("ISTEXT on string cell") {
        REQUIRE(eng.evaluate("=ISTEXT(A2)", resolver).get<bool>() == true);
    }

    doc.close();
}
