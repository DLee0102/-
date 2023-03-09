
# set(CMAKE_PREFIX_PATH ${PROJECT_SOURCE_DIR}/vendors)        # For QT
set(QT_QMAKE_EXECUTABBLE ${PROJECT_SOURCE_DIR}/vendors/QT/bin/qmake.exe)       # For QT

set(CMAKE_INCLUDE_CURRENT_DIR ON)      # For QT
set(CMAKE_AUTOUIC ON)      # For QT
set(CMAKE_AUTOMOC ON)      # For QT
set(CMAKE_AUTORCC ON)      # For QT
set(CMAKE_CXX_STANDARD 17)     # For QT
set(CMAKE_CXX_STANDARD_REQUIRED ON)         # For QT

# add_subdirectory(external/nowide) # 用于libzip
# message("${CMAKE_PREFIX_PATH}")

# 该库（libzip）必须先安装才能使用，安装后安装程序会自动将LIBZIP_DIR设置为安装路径下的config.cmake文件路径
# set(LIBZIP_DIR ${PROJECT_SOURCE_DIR}/vendors/libzip-main/build)
find_package(Qt5 COMPONENTS Widgets LinguistTools REQUIRED)         # For QT
# find_package(LIBZIP REQUIRED)

set(TS_FILES Dxx_tool_v0.2_zh_CN.ts)        # For QT

add_executable(Dxx_tool_v0.2 Core.cpp Filemanagement.cpp main.cpp mainwindow.cpp mainwindow.h mainwindow.ui app.rc ${TS_FILES})

target_link_libraries(Dxx_tool_v0.2 OpenXLSX::OpenXLSX Qt5::Widgets)
# lib::zip

qt5_create_translation(QM_FILES ${CMAKE_SOURCE_DIR}/src/Dxx_tool_v0.2 ${TS_FILES})         # For QT

target_include_directories(Dxx_tool_v0.2 PUBLIC ${PROJECT_SOURCE_DIR}/src/Dxx_tool_v0.2
                                                )