add_executable(TestOpenXLSX test.cpp)

add_executable(Version1 version1.cpp)

target_link_libraries(TestOpenXLSX OpenXLSX::OpenXLSX)

target_link_libraries(Version1 OpenXLSX::OpenXLSX)

# message(${OpenXLSX_INCLUDE_DIR})
