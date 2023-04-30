#pragma once

#include <string>
#include <OpenXLSX.hpp>
#include <iostream>
#include <vector>
#define STARTROW_TOTAL 3
#define STARTROW_TEMP 2
#define STARTCOL_TOTAL 5
#define CHECKEDNAMECOL 1
#define LOADEDNAMECOL 3

namespace DXXTL{
    struct FileNamesplit
    {
        std::string fgrade;
        std::string fclass;
        std::string fterms;
    };

    class Excelops
    {
    public:
        Excelops(std::string obj_, std::vector<std::string>& files_, int temp_start_row_, int total_start_row_);    // 初始化框架
        void traverseFiles();
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

        // 待检索的文件路径数组
        std::vector<std::string> files;
        // 文件名中包含的各个属性
        FileNamesplit* temp_fns;

        int temp_length;
        int total_length;
        int col_length;

        // 数据从第几行开始
        int temp_start_row;
        int total_start_row;

        int getrowCount(OpenXLSX::XLWorksheet wks);    // 获取行长度
        int getcolumnCount(OpenXLSX::XLWorksheet wks);    // 获取列长度
        bool initWritevector(std::vector<OpenXLSX::XLCellValue>& writeValues, int length);    // 初始化写入向量
        bool init(std::string obj_, std::vector<std::string>& files_, int temp_start_row_, int total_start_row_);    // 初始化框架
        void fillWriteValues(int checkedNamecol, int loadedNamecol);    // 将数据读取到写入向量中
        void loadData(int termnumber);    // 将写入向量中的数据写到指定表中    待重写
        void retrieveFile(std::string file_);   // 检索一个文件
        std::string getNameproperty(std::string file_);      // 由文件路径获取文件名
        void assignFilenamepro(std::string file_);   // 给文件名属性赋值
    };
}
