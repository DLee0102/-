#include "Core.h"

namespace DXXTL {
    Excelops::Excelops(std::string obj_, std::vector<std::string>& files_, int temp_start_row_, int total_start_row_)
    {
        if (init(obj_, files, temp_start_row_, total_start_row_)) {
            std::cout << "                    初始化完成..." << std::endl;
        }
    }

    int Excelops::getRowlength(OpenXLSX::XLWorksheet wks, int start_row)
    {
        start_row--;
        std::vector<OpenXLSX::XLCellValue> readValues;
        for (auto& row : wks.rows()) {
            readValues = row.values();
            if (readValues[0] == "") {
                break;
            }
            start_row++;
        }

        return start_row;
    }

    /**
     * @brief 以除表头外第一行内容为参照确定列数
     * 
     * @param wks 
     * @return int 
     */
    int Excelops::getCollength(OpenXLSX::XLWorksheet wks, int start_row)
    {
        int col_length = 0;
        std::vector<OpenXLSX::XLCellValue> readValues;
        for (auto& row : wks.rows(start_row, total_length)) {
            readValues = row.values();
            for (int i = 0; i < readValues.size(); i++) {
                if (readValues[i] == "") {break;}
                col_length++;
            }
            break;
        }

        return col_length;
    }

    bool Excelops::initWritevector(std::vector<OpenXLSX::XLCellValue>& writeValues, int length)
    {
        for (int i = 0; i < length; i++) {
            writeValues[i] = 0;
        }
    }

    void Excelops::traverseFiles()
    {
        std::vector<std::string>::iterator pd;
        for (pd = files.begin(); pd != files.end(); pd ++)
        {
            retrieveFile(*pd);
        }
    }

    void Excelops::retrieveFile(std::string file_)
    {
        doc_total.open(total_path);
        wks_total = doc_total.workbook().worksheet("21级2班团支部");
    }

    void Excelops::retriveSheet()
    {

    }

    bool Excelops::init(std::string obj_, std::vector<std::string>& files_, int temp_start_row_, int total_start_row_)
    {
        total_path = obj_;
        files.assign(files_.begin(), files_.end());

        temp_start_row = temp_start_row_;
        total_start_row = total_start_row_;

        // doc_total.open(total_path);
        // wks_total = doc_total.workbook().worksheet("21级2班团支部");

        // doc_temp.open(temp_path);
        // wks_temp = doc_temp.workbook().sheet(1).get<OpenXLSX::XLWorksheet>();

        // temp_length = getRowlength(wks_temp, 1);
        // total_length = getRowlength(wks_total, 1);
        // col_length = getCollength(wks_total, total_start_row);
        // std::cout << "                    OpenXLSX框架初始化完成..." << std::endl;

        // writeValues_total.resize(total_length - 2);
        // if (initWritevector(writeValues_total, total_length - 2)) {std::cout << "                    写入程序初始化完成..." << std::endl;}

        return true;
    }

    void Excelops::fillWriteValues(int checkedNamecol, int loadedNamecol)
    {
        for (auto& row : wks_temp.rows(2, temp_length)) {
            readValues_temp = row.values();
            if (readValues_temp.size() != 0) {
                int index = 0;
                for (auto &row_m : wks_total.rows(3, total_length)) {
                    readValues_total = row_m.values();
                    // std::cout << col_length << std::endl;
                    
                    if (readValues_total.size() != 0) {
                        // std::cout << readValues_temp[checkedNamecol] << " " << readValues_total[loadedNamecol] << std::endl;
                        if (readValues_temp[checkedNamecol].get<std::string>().compare(readValues_total[loadedNamecol].get<std::string>()) == 0) {
                            writeValues_total[index] = 1; 
                        }
                    }
                    index++;
                }
            }
        }
    }

    // void Excelops::loadData()
    // {
    //     std::cout << "                    正在填入数据，请稍等..." << std::endl;
    //     std::cout << "                    填入数据量为：" << " ";
    //     std::cout << writeValues_total.size() << std::endl;
    //     for (int i = 0; i < writeValues_total.size(); i++) {
    //         wks_total.cell(total_start_row + i, col_length + 1).value() = writeValues_total[i];
    //         // std::cout << writeValues_total[i] << std::endl;
    //     }

    //     std::cout << "                    完成!" << std::endl;
    // }

    Excelops::~Excelops()
    {
        doc_temp.save();
        doc_temp.close();

        doc_total.save();
        doc_total.close();

        readValues_total.clear();
        writeValues_total.clear();
        readValues_temp.clear();

    }
}