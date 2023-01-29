#pragma once

#include <string>
#include <OpenXLSX.hpp>
#include <iostream>

namespace DXXTL{
    class Excelops
    {
    public:
        Excelops(std::string obj_);
        int getRowlength(OpenXLSX::XLWorksheet wks, int start_row);
        int getCollength(OpenXLSX::XLWorksheet wks);
        bool initWritevector(std::vector<OpenXLSX::XLCellValue>& writeValues, int length);
        bool init(std::string obj_);
        void fillWriteValues(int checkedNamecol, int loadedNamecol);
        void loadData(int start_row);
        ~Excelops();
    
    private:
        std::string total_path;
        OpenXLSX::XLDocument doc_total;
        OpenXLSX::XLWorksheet wks_total;
        std::vector<OpenXLSX::XLCellValue> readValues_total;
        std::vector<OpenXLSX::XLCellValue> writeValues_total;

        std::string temp_path;
        OpenXLSX::XLDocument doc_temp;
        OpenXLSX::XLWorksheet wks_temp;
        std::vector<OpenXLSX::XLCellValue> readValues_temp;

        int temp_length;
        int total_length;
        int col_length;
    };
}
