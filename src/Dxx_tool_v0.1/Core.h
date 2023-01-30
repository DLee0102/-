#pragma once

#include <string>
#include <OpenXLSX.hpp>
#include <iostream>

namespace DXXTL{
    class Excelops
    {
    public:
        Excelops(std::string obj_, int temp_start_row_, int total_start_row_);    // 初始化框架
        int getRowlength(OpenXLSX::XLWorksheet wks, int start_row);    // 获取行长度
        int getCollength(OpenXLSX::XLWorksheet wks, int start_row);    // 获取列长度
        bool initWritevector(std::vector<OpenXLSX::XLCellValue>& writeValues, int length);    // 初始化写入向量
        bool init(std::string obj_, int temp_start_row_, int total_start_row_);    // 初始化框架
        void fillWriteValues(int checkedNamecol, int loadedNamecol);    // 将数据读取到写入向量中
        void loadData();    // 将写入向量中的数据写到指定表中
        ~Excelops();    // 关闭表格文件
    
    private:
        /**
         * @brief 待写入的文件
         * 
         */
        std::string total_path;
        OpenXLSX::XLDocument doc_total;
        OpenXLSX::XLWorksheet wks_total;
        std::vector<OpenXLSX::XLCellValue> readValues_total;
        std::vector<OpenXLSX::XLCellValue> writeValues_total;

        /**
         * @brief 待读取的文件
         * 
         */
        std::string temp_path;
        OpenXLSX::XLDocument doc_temp;
        OpenXLSX::XLWorksheet wks_temp;
        std::vector<OpenXLSX::XLCellValue> readValues_temp;

        int temp_length;
        int total_length;
        int col_length;

        // 数据从第几行开始
        int temp_start_row;
        int total_start_row;
    };
}
