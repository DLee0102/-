# Hazel_engine's CMakeLists.txt 作者(Author)：丁利(github: DLee0102)
# 四川大学软件学院(Software Engineering College of Sichuan University)

cmake_minimum_required(VERSION 3.10)

set(CMAKE_CXX_STANDARD 17)

# Set the build output location to a common directory
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)
# set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin)

# Linux和Windows平台参数定义不同, 要注意区分; Windows下既是可执行文件，也是动态库文件存放处
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin)

project(青年大学习统计工具 CXX)

add_subdirectory(vendors/OpenXLSX)

add_subdirectory(src/Test)

add_subdirectory(src/Dxx_tool_v0.1)

add_subdirectory(src/Dxx_tool_v0.2)