#include "Core.h"

namespace DXXTL {
    Excelops::Excelops(std::string obj_)
    {
        if (init(obj_)) {
            std::cout << "                    ��ʼ�����..." << std::endl;
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

    int Excelops::getCollength(OpenXLSX::XLWorksheet wks)
    {
        int col_length = 0;
        std::vector<OpenXLSX::XLCellValue> readValues;
        for (auto& row : wks.rows(3, total_length)) {
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

    bool Excelops::init(std::string obj_)
    {
        total_path = "./���ѧԺ2021�������ѧϰѧϰ�����¼.xlsx";
        temp_path = "./ѧϰ�û���ϸ.xlsx";

        doc_total.open(total_path);
        wks_total = doc_total.workbook().worksheet(obj_);

        doc_temp.open(temp_path);
        wks_temp = doc_temp.workbook().sheet(1).get<OpenXLSX::XLWorksheet>();

        temp_length = getRowlength(wks_temp, 1);
        total_length = getRowlength(wks_total, 1);
        col_length = getCollength(wks_total);
        std::cout << "                    OpenXLSX��ܳ�ʼ�����..." << std::endl;

        writeValues_total.resize(total_length - 2);
        if (initWritevector(writeValues_total, total_length - 2)) {std::cout << "                    д������ʼ�����..." << std::endl;}

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

    void Excelops::loadData(int start_row)
    {
        std::cout << "                    �����������ݣ����Ե�..." << std::endl;
        std::cout << "                    ����������Ϊ��" << " ";
        std::cout << writeValues_total.size() << std::endl;
        for (int i = 0; i < writeValues_total.size(); i++) {
            wks_total.cell(start_row + i, col_length + 1).value() = writeValues_total[i];
            std::cout << writeValues_total[i] << std::endl;
        }

        std::cout << "                    ���!" << std::endl;
    }

    Excelops::~Excelops()
    {
        doc_temp.save();
        doc_temp.close();

        doc_total.save();
        doc_total.close();
    }
}