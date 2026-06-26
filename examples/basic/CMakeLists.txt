add_executable(qdocx_basic
    main.cpp
)

target_link_libraries(qdocx_basic
    PRIVATE
        QDocx::QDocx
        Qt6::Core
)
