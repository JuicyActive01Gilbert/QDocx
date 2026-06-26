add_executable(qdocx_manual_report_test
    main.cpp
)

target_link_libraries(qdocx_manual_report_test
    PRIVATE
        QDocx::QDocx
        Qt6::Core
)

add_test(NAME qdocx_manual_report_api
    COMMAND qdocx_manual_report_test
)
