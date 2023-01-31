#include "Core.h"
#include <iostream>

/**
 * @brief 程序入口
 * 
 * @return int 
 */
int main() {
    /**
     * @brief 界面
     * 
     */
    std::cout << "                    ---------------------------------------青年大学习自动录入工具v0.1---------------------------------------------" << std::endl;
    std::cout << "                    使用说明：" << std::endl;
    std::cout << "                        1. 在文件资源管理器中Dxx_tool_v0.1.exe同级目录下，按住shift的同时单机鼠标右键，在菜单栏中选择打开powershell。" << std::endl;
    std::cout << "                        2. 输入chcp 65001，按下回车。" << std::endl;
    std::cout << "                        3. 输入.\\Dxx_tool_v0.1.exe，按下回车，进入软件界面。" << std::endl;
    std::cout << "                        4. 将需要统计的那一期的 学习用户明细.xlsx 文件放在和 Dxx_tool_v0.1.exe 同级目录下。" << std::endl;
    std::cout << "                        5. 打开 学习用户明细.xlsx 文件，按下ctrl + s保存文件后关闭 学习用户明细.xlsx。" << std::endl;
    std::cout << "                        6. 从在线文档下载 软件学院2021级青年大学习学习情况记录.xlsx ，并放在和Dxx_tool_v0.1.exe同级目录下。" << std::endl;
    std::cout << "                        7. 特别注意：在使用本工具过程中需关闭需要读入和写入的文件。" << std::endl;
    std::cout << "                        8. 后续根据提示进行操作。" << std::endl;
    std::cout << "                    =============================================================================================================" << std::endl;
    std::cout << "                    +++++++++++++++++++++++++++++++++++++++++ 欢 迎 使 用 本 工 具 +++++++++++++++++++++++++++++++++++++++++++++++" << std::endl;
    std::cout << "                    =============================================================================================================" << std::endl;

    /**
     * @brief 处理输入
     * 
     */
    std::string class_;
    std::string grade_;
    std::string obj_;

    std::cout << "                    提示：请输入要处理的年级（输入示例：21）：" << " ";
    std::cin >> grade_;
    std::cout << "                    提示：请输入要处理的班级（输入示例：1）：" << " ";
    std::cin >> class_;
    obj_ = grade_ + "级" + class_ + "班团支部";

    /**
     * @brief 调用内核
     * 
     */
    DXXTL::Excelops* item = new DXXTL::Excelops(obj_, 2, 3);    // // 数据从第几行开始
    item->fillWriteValues(0, 2);    // 需要检查的姓名分别在两个文件中所在的列数
    item->loadData();       // 从第三行开始记录数据
    item->~Excelops();

}