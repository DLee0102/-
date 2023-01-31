
set(CMAKE_PREFIX_PATH ${PROJECT_SOURCE_DIR}/vendors)        # For QT
set(QT_QMAKE_EXECUTABBLE ${PROJECT_SOURCE_DIR}/vendors/QT/bin/qmake.exe)       # For QT

set(CMAKE_INCLUDE_CURRENT_DIR ON)      # For QT
set(CMAKE_AUTOUIC ON)      # For QT
set(CMAKE_AUTOMOC ON)      # For QT
set(CMAKE_AUTORCC ON)      # For QT
set(CMAKE_CXX_STANDARD 11)     # For QT
set(CMAKE_CXX_STANDARD_REQUIRED ON)         # For QT

find_package(Qt5 COMPONENTS Widgets LinguistTools REQUIRED)         # For QT

set(TS_FILES Youth_learning_autoload_GUI_zh_CN.ts)        # For QT

add_executable(Dxx_tool_v0.2 main.cpp mainwindow.cpp mainwindow.h mainwindow.ui ${TS_FILES})

target_link_libraries(Dxx_tool_v0.2 OpenXLSX::OpenXLSX Qt5::Widgets)

qt5_create_translation(QM_FILES ${CMAKE_SOURCE_DIR}/src/Dxx_tool_v0.2 ${TS_FILES})         # For QT

target_include_directories(Dxx_tool_v0.2 PUBLIC ${PROJECT_SOURCE_DIR}/src/Dxx_tool_v0.2)