# Hazel_engine's CMakeLists.txt 作者(Author)：丁利(github: DLee0102)
# 四川大学软件学院(Software Engineering College of Sichuan University)


# 关于多编译器混合编译的说明：
# 首先，不同编译器编译的包处于同于项目中可能会报错
# 其次，由于系统默认编译器路径和cmake配置中使用的编译器不同，可能导致最后生成的exe文件在运行时找不到程序入口
# 原因就是exe文件找不到libstdc++-6.dll这个动态库文件（默认通过系统环境变量路径找dll，所以如果默认编译器和build用的
# 编译器不一样，就会运行时报错），像QT这种使用自带编译器的库解决办法是将我们build所用的编译器的libstdc++-6.dll贴到
# 和exe文件同级目录下。若遇到其他编译器冲突的问题，也可以采用这种解决思路

# 若不用QT自带的编译器编译，则会出现无法定位程序输入点__gxx_personality_sj0于动态链接库Dxx_tool_v0.2.exe上 的问题（即使打包程序也会出现这个错误）。
# 环境变量中两个QT的编译器路径一定要放在最上面，其中Tool在最上面，这样才能在本地调试运行。但是用户就不必如此操作，使用打包好的程序即可（打包好的程序已经包含了
# 程序运行所需的动态链接库）。

# QT编译的项目最后一定要打包!

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
