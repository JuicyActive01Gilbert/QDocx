add_executable(qdocx_report_studio
    main.cpp
)

target_link_libraries(qdocx_report_studio
    PRIVATE
        QDocx::QDocx
        Qt6::Core
        Qt6::Gui
        Qt6::Widgets
)

if(MSVC)
    target_compile_options(qdocx_report_studio PRIVATE /utf-8)
endif()
