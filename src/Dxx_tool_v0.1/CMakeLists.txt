add_executable(Dxx_tool_v0.1 Entrypoint.cpp Core.cpp)

target_link_libraries(Dxx_tool_v0.1 OpenXLSX::OpenXLSX)

target_include_directories(Dxx_tool_v0.1 PUBLIC ${PROJECT_SOURCE_DIR}/src/Dxx_tool_v0.1)
