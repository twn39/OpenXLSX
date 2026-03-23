#ifndef OPENXLSX_XLSTRINGARENA_HPP
#define OPENXLSX_XLSTRINGARENA_HPP

#include <vector>
#include <memory>
#include <string_view>
#include <algorithm>
#include <cstdint>

namespace OpenXLSX {

/**
 * @brief String memory pool (Arena Allocator)
 * @details Adheres to C++ Core Guidelines:
 *          1. RAII resource management (no naked new/delete)
 *          2. No copy construction/assignment to prevent dangling pointers (Rule of 5)
 *          3. Provides zero-copy access via std::string_view
 */
class XLStringArena {
public:
    // Default block size set to 4MB to accommodate many short strings
    explicit XLStringArena(size_t blockSize = 4 * 1024 * 1024) 
        : m_blockSize(blockSize), m_currentOffset(blockSize) {}

    // Disable copy operations to prevent dangling string_view
    XLStringArena(const XLStringArena&) = delete;
    XLStringArena& operator=(const XLStringArena&) = delete;
    
    // Allow move operations
    XLStringArena(XLStringArena&&) noexcept = default;
    XLStringArena& operator=(XLStringArena&&) noexcept = default;

    /**
     * @brief Store string in the arena, returning a non-owning view
     * @param str The string view to store
     * @return std::string_view pointing to the stored characters in the arena
     */
    [[nodiscard]] std::string_view store(std::string_view str) {
        if (str.empty()) {
            static constexpr char emptyStr[] = "";
            return {emptyStr, 0};
        }

        // Allocate a new block if the current one has insufficient capacity (+1 for null terminator)
        if (m_currentOffset + str.size() + 1 > m_blockSize) {
            // Handle edge case for extremely large strings
            const size_t allocSize = std::max(m_blockSize, str.size() + 1);
            m_blocks.push_back(std::make_unique<char[]>(allocSize));
            m_currentOffset = 0;
        }

        // Get the destination pointer and copy data
        char* destPtr = m_blocks.back().get() + m_currentOffset;
        std::copy(str.begin(), str.end(), destPtr);
        destPtr[str.size()] = '\0'; // ensure null-terminated
        
        std::string_view storedView(destPtr, str.size());
        m_currentOffset += str.size() + 1; // advance offset including null terminator

        return storedView;
    }

    /**
     * @brief Clear the arena, releasing all resources
     */
    void clear() noexcept {
        m_blocks.clear();
        m_currentOffset = m_blockSize; // Force new allocation on next store
    }

private:
    std::vector<std::unique_ptr<char[]>> m_blocks; // RAII memory blocks
    size_t m_blockSize;                            // Default size per block
    size_t m_currentOffset;                        // Write offset in current block
};

} // namespace OpenXLSX

#endif // OPENXLSX_XLSTRINGARENA_HPP