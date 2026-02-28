#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.15)
project(OpenXLSX)

#=======================================================================================================================
# OPTIONS
#=======================================================================================================================
option(OPENXLSX_BUILD_SAMPLES "Build the OpenXLSX sample programs" ON)
option(OPENXLSX_BUILD_TESTS "Build the OpenXLSX test programs" OFF)
option(OPENXLSX_BUILD_BENCHMARKS "Build the OpenXLSX benchmark programs" OFF)
option(OPENXLSX_BUILD_DOCS "Build the OpenXLSX documentation" OFF)
option(OPENXLSX_ENABLE_NOWIDE "Enable the Nowide library for UTF-8 support on Windows" ON)
option(OPENXLSX_ENABLE_LTO "Enable Link-Time Optimization (LTO)" OFF)

set(OPENXLSX_LIBRARY_TYPE "STATIC" CACHE STRING "The type of library to build (STATIC or SHARED)")
set_property(CACHE OPENXLSX_LIBRARY_TYPE PROPERTY STRINGS STATIC SHARED)

#=======================================================================================================================
# SOURCES
#=======================================================================================================================
file(GLOB OPENXLSX_SOURCES sources/*.cpp)

#=======================================================================================================================
# STATIC LIBRARY
#   Define the static library
#=======================================================================================================================
if ("${OPENXLSX_LIBRARY_TYPE}" STREQUAL "STATIC")
    add_library(OpenXLSX STATIC "")
    add_library(OpenXLSX::OpenXLSX ALIAS OpenXLSX)
    target_sources(OpenXLSX PRIVATE ${OPENXLSX_SOURCES})
    target_include_directories(OpenXLSX
            PUBLIC
            $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}>
            $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/headers>
            $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>     # For export header
            SYSTEM PUBLIC
            $<BUILD_INTERFACE:${fast_float_SOURCE_DIR}/include>
            $<BUILD_INTERFACE:${fmt_SOURCE_DIR}/include>
            $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/../third_party/pugixml/src>
            $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/../third_party/zippy>)
    target_link_libraries(OpenXLSX
            PUBLIC
            $<BUILD_INTERFACE:pugixml>
            PRIVATE
            $<BUILD_INTERFACE:Zippy>)

    if (OPENXLSX_ENABLE_NOWIDE)
        target_link_libraries(OpenXLSX
                PRIVATE
                $<BUILD_INTERFACE:NoWide>)
    endif ()

    target_compile_definitions(OpenXLSX PUBLIC OPENXLSX_STATIC_DEFINE FMT_HEADER_ONLY)

endif ()

#=======================================================================================================================
# SHARED LIBRARY
#   Define the shared library
#=======================================================================================================================
if ("${OPENXLSX_LIBRARY_TYPE}" STREQUAL "SHARED")
    add_library(OpenXLSX SHARED "")
    add_library(OpenXLSX::OpenXLSX ALIAS OpenXLSX)
    target_sources(OpenXLSX PRIVATE ${OPENXLSX_SOURCES})
    target_include_directories(OpenXLSX
            PUBLIC
            $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}>
            $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/headers>
            $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>     # For export header
            SYSTEM PUBLIC
            $<BUILD_INTERFACE:${fast_float_SOURCE_DIR}/include>
            $<BUILD_INTERFACE:${fmt_SOURCE_DIR}/include>
            $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/../third_party/pugixml/src>
            $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/../third_party/zippy>)
    target_link_libraries(OpenXLSX
            PUBLIC
            $<BUILD_INTERFACE:pugixml>
            PRIVATE
            $<BUILD_INTERFACE:Zippy>)

    if (OPENXLSX_ENABLE_NOWIDE)
        target_link_libraries(OpenXLSX
                PRIVATE
                $<BUILD_INTERFACE:NoWide>)
    endif ()

    # Enable Link-Time Optimization (LTO)
    if(${OPENXLSX_ENABLE_LTO})
        include(CheckIPOSupported)
        check_ipo_supported(RESULT result OUTPUT output)
        if (result)
            set_property(TARGET OpenXLSX PROPERTY INTERPROCEDURAL_OPTIMIZATION TRUE)
        endif ()
    endif()

    target_compile_definitions(OpenXLSX PUBLIC FMT_HEADER_ONLY)

endif ()


if (OPENXLSX_ENABLE_NOWIDE)
    target_compile_definitions(OpenXLSX PRIVATE ENABLE_NOWIDE)
endif ()


# Generate export header
include(GenerateExportHeader)
generate_export_header(OpenXLSX
        BASE_NAME openxlsx
        EXPORT_FILE_NAME OpenXLSX-Exports.hpp
        EXPORT_MACRO_NAME OPENXLSX_EXPORT
        NO_EXPORT_MACRO_NAME OPENXLSX_HIDDEN)

#=======================================================================================================================
# COMPILER FLAGS
#   Set compiler debug flags for GCC, Clang and MSVC.
#=======================================================================================================================
list(APPEND OPENXLSX_DEBUG_FLAGS_GNU
        "-Wmisleading-indentation"
        "-Wduplicated-cond"
        "-Wduplicated-branches"
        "-Wlogical-op"
        "-Wnull-dereference")
list(APPEND OPENXLSX_DEBUG_FLAGS_GNUCLANG
        "-Wall"
        "-Wextra"
        "-Wshadow"
        "-Wnon-virtual-dtor"
        "-Wold-style-cast"
        "-Wcast-align"
        "-Wunused"
        "-Woverloaded-virtual"
        "-Wpedantic"
        "-Wconversion"
        "-Wsign-conversion"
        "-Wmisleading-indentation"
        "-Wnull-dereference"
        "-Wdouble-promotion"
        "-Wformat=2")
list(APPEND OPENXLSX_DEBUG_FLAGS_MSVC
        "/W4"
        "/permissive-")

if (CMAKE_CXX_COMPILER_ID STREQUAL "GNU")
    target_compile_options(OpenXLSX PRIVATE ${OPENXLSX_DEBUG_FLAGS_GNU})
elseif (CMAKE_CXX_COMPILER_ID STREQUAL "Clang" OR CMAKE_CXX_COMPILER_ID STREQUAL "AppleClang")
    target_compile_options(OpenXLSX PRIVATE ${OPENXLSX_DEBUG_FLAGS_GNUCLANG})
elseif (CMAKE_CXX_COMPILER_ID STREQUAL "MSVC")
    target_compile_options(OpenXLSX PRIVATE ${OPENXLSX_DEBUG_FLAGS_MSVC})
endif ()

#=======================================================================================================================
# INSTALLATION
#   Set up installation of the library
#=======================================================================================================================
set(ConfigPackageLocation lib/cmake/OpenXLSX)
install(
        TARGETS OpenXLSX
        EXPORT OpenXLSXTargets
        LIBRARY DESTINATION lib
        ARCHIVE DESTINATION lib
        RUNTIME DESTINATION bin
        INCLUDES DESTINATION include
)
install(
        DIRECTORY headers/
        DESTINATION include/OpenXLSX/headers
        FILES_MATCHING PATTERN "*.hpp"
)
install(
        FILES OpenXLSX.hpp
        DESTINATION include/OpenXLSX
)
install(
        FILES "${CMAKE_CURRENT_BINARY_DIR}/OpenXLSX-Exports.hpp"
        DESTINATION include/OpenXLSX
)
install(
        EXPORT OpenXLSXTargets
        FILE OpenXLSXTargets.cmake
        NAMESPACE OpenXLSX::
        DESTINATION ${ConfigPackageLocation}
)
export(
        EXPORT OpenXLSXTargets
        FILE "${CMAKE_CURRENT_BINARY_DIR}/OpenXLSX/OpenXLSXTargets.cmake"
        NAMESPACE OpenXLSX::
)
install(
        EXPORT OpenXLSXTargets
        FILE OpenXLSXTargets.cmake
        NAMESPACE OpenXLSX::
        DESTINATION ${ConfigPackageLocation}
)
