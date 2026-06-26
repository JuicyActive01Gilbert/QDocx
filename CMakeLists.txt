cmake_minimum_required(VERSION 3.21)

project(QDocx
    VERSION 1.0.0
    DESCRIPTION "Qt wrapper around Microsoft Word COM automation"
    LANGUAGES CXX
)

option(QDOCX_BUILD_EXAMPLES "Build QDocx examples" ON)
option(QDOCX_BUILD_TESTS "Build QDocx local test projects" ON)
option(QDOCX_BUILD_APPS "Build QDocx demonstration applications" ON)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

find_package(Qt6 REQUIRED COMPONENTS Core Gui Widgets AxContainer)

add_library(QDocx SHARED
    include/QDocx/QDocx
    include/QDocx/qdocxdocument.h
    include/QDocx/qdocxglobal.h
    include/QDocx/qdocxresult.h
    include/QDocx/qdocxstyle.h
    include/QDocx/qdocxtable.h
    src/qdocxdocument.cpp
    src/qdocxresult.cpp
    src/qdocxtable.cpp
    src/word/qdocxwordbackend.cpp
    src/word/qdocxwordbackend.h
)

add_library(QDocx::QDocx ALIAS QDocx)

target_compile_definitions(QDocx
    PRIVATE QDOCX_LIBRARY
)

target_include_directories(QDocx
    PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
        $<INSTALL_INTERFACE:include>
    PRIVATE
        ${CMAKE_CURRENT_SOURCE_DIR}/src
)

target_link_libraries(QDocx
    PUBLIC
        Qt6::Core
        Qt6::Gui
        Qt6::AxContainer
)

set_target_properties(QDocx PROPERTIES
    OUTPUT_NAME QDocx
    WINDOWS_EXPORT_ALL_SYMBOLS OFF
    DEBUG_POSTFIX d
)

if(MSVC)
    target_compile_options(QDocx PRIVATE /W4 /permissive-)
endif()

install(TARGETS QDocx
    EXPORT QDocxTargets
    RUNTIME DESTINATION bin
    LIBRARY DESTINATION lib
    ARCHIVE DESTINATION lib
)

install(DIRECTORY include/ DESTINATION include)

install(EXPORT QDocxTargets
    NAMESPACE QDocx::
    DESTINATION lib/cmake/QDocx
)

include(CMakePackageConfigHelpers)

configure_package_config_file(
    cmake/QDocxConfig.cmake.in
    ${CMAKE_CURRENT_BINARY_DIR}/QDocxConfig.cmake
    INSTALL_DESTINATION lib/cmake/QDocx
)

write_basic_package_version_file(
    ${CMAKE_CURRENT_BINARY_DIR}/QDocxConfigVersion.cmake
    VERSION ${PROJECT_VERSION}
    COMPATIBILITY SameMajorVersion
)

install(FILES
    ${CMAKE_CURRENT_BINARY_DIR}/QDocxConfig.cmake
    ${CMAKE_CURRENT_BINARY_DIR}/QDocxConfigVersion.cmake
    DESTINATION lib/cmake/QDocx
)

if(QDOCX_BUILD_EXAMPLES)
    add_subdirectory(examples/basic)
endif()

if(QDOCX_BUILD_TESTS)
    enable_testing()
    add_subdirectory(tests/manual-report)
endif()

if(QDOCX_BUILD_APPS)
    add_subdirectory(apps/report-studio)
endif()
