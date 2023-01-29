#include "Core.h"
#include <iostream>


int main() {
    std::cout << "                    ---------------------------------------青年大学习自动录入工具v0.1--------------------------------------" << std::endl;
    std::cout << "                    使用说明：" << std::endl;
    std::cout << "                        1. 将需要统计的那一期的 学习用户明细.xlsx 文件放在和 青年大学习自动录入工具.exe 同级目录下。" << std::endl;
    std::cout << "                        2. 打开 学习用户明细.xlsx 文件，按下ctrl + s保存文件后关闭 学习用户明细.xlsx。" << std::endl;
    std::cout << "                        3. 从在线文档下载 软件学院2021级青年大学习学习情况记录.xlsx ，并放在和青年大学习自动录入工具同级目录下。" << std::endl;
    std::cout << "                        4. 后续根据提示进行操作。" << std::endl;
    std::cout << "                    ======================================================================================================" << std::endl;
    std::cout << "                    +++++++++++++++++++++++++++++++++++++++++ 欢 迎 使 用 本 工 具 ++++++++++++++++++++++++++++++++++++++++" << std::endl;
    std::cout << "                    ======================================================================================================" << std::endl;


    std::string class_;
    std::string grade_;
    std::string obj_;

    std::cout << "                    提示：请输入要处理的年级（输入示例：21）：" << " ";
    std::cin >> grade_;
    std::cout << "                    请输入要处理的班级（输入示例：1）：" << " ";
    std::cin >> class_;
    obj_ = grade_ + "级" + class_ + "班团支部";

    DXXTL::Excelops* item = new DXXTL::Excelops(obj_);
    item->fillWriteValues(0, 2);
    item->loadData(3);
    item->~Excelops();

}