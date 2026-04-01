#include <OpenXLSX.hpp>
#include <string>
#include <vector>

extern "C" int LLVMFuzzerTestOneInput(const uint8_t* data, size_t size)
{
    if (size == 0) return 0;

    // Create a string from the fuzzer input
    std::string input(reinterpret_cast<const char*>(data), size);

    // Ensure it starts with '=' so the engine actually tries to parse it as a formula
    if (input.front() != '=') { input = "=" + input; }

    try {
        OpenXLSX::XLFormulaEngine engine;
        auto                      result = engine.evaluate(input);
        (void)result;    // Ignore result to prevent nodiscard warnings
    }
    catch (const OpenXLSX::XLException&) {
        // Expected behavior for garbage input: throw an XLException (e.g. XLFormulaError, XLInputError)
        // We catch it because we only care if it causes a SEGFAULT, ASSERT FAILURE, or HANG.
    }
    catch (...) {
        // Any other C++ exception is technically a bug in error handling, but we ignore it for pure memory fuzzing
    }

    return 0;
}