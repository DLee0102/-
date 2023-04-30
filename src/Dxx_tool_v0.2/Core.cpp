#include "Core.h"

namespace DXXTL {
    /**
     * @brief 构造一个新的 Excelops:: Excelops 对象
     * 
     * @param obj_ 
     * @param files_ 
     * @param temp_start_row_ 
     * @param total_start_row_ 
     */
    Excelops::Excelops(std::string obj_, std::vector<std::string>& files_, int temp_start_row_, int total_start_row_)
    {
        if (init(obj_, files_, temp_start_row_, total_start_row_)) {
            std::cout << "                    初始化完成..." << std::endl;
        }
    }

    /**
     * @brief 获取行数
     * 
     * @param wks 
     * @param start_row 
     * @return int 
     */
    int Excelops::getrowCount(OpenXLSX::XLWorksheet wks)
    {
        // start_row--;
        // std::vector<OpenXLSX::XLCellValue> readValues;
        
        // for (auto itr = wks.rows().begin(); itr != wks.rows().end(); itr ++)
        // {
        //     std::cout << itr->cells().size() << std::endl;
        //     start_row++;
        // }

        // std::cout << wks.rowCount() << std::endl;

        return wks.rowCount();
    }

    /**
     * @brief 以除表头外第一行内容为参照确定列数
     * 
     * @param wks 
     * @return int 
     */
    int Excelops::getcolumnCount(OpenXLSX::XLWorksheet wks)
    {
        // int col_length = 0;
        // std::vector<OpenXLSX::XLCellValue> readValues;
        // for (auto& row : wks.rows(start_row, total_length)) {
        //     readValues = row.values();
        //     for (int i = 0; i < readValues.size(); i++) {
        //         if (readValues[i] == "") {break;}
        //         col_length++;
        //     }
        //     break;
        // }

        return wks.columnCount();
    }

    /**
     * @brief 初始化写入Vector
     * 
     * @param writeValues 
     * @param length 
     * @return true 
     * @return false 
     */
    bool Excelops::initWritevector(std::vector<OpenXLSX::XLCellValue>& writeValues, int length)
    {
        for (int i = 0; i < length; i++) {
            writeValues[i] = 0;
        }
    }

    /**
     * @brief 遍历文件夹下的所有文件
     * 
     */
    void Excelops::traverseFiles()
    {
        std::vector<std::string>::iterator pd;
        for (pd = files.begin(); pd != files.end(); pd ++)
        {
            retrieveFile(*pd);
        }
    }

    std::string Excelops::getNameproperty(std::string file_)
    {
        std::string filename_;
        int namepos = 0;
        // std::cout << file_ << std::endl;
        namepos = file_.rfind('/');
        filename_.assign(file_.begin() + namepos + 1, file_.end());
        // std::cout << filename_ << std::endl;

        return filename_;
    }

    void Excelops::assignFilenamepro(std::string file_)
    {
        temp_fns->fgrade.assign(file_.begin(), file_.begin() + file_.find('_'));
        temp_fns->fclass.assign(file_.begin() + file_.find('_') + 1, file_.begin() + file_.find('_', file_.find('_') + 1));
        temp_fns->fterms.assign(file_.begin() + file_.find('_', file_.find('_') + 1) + 1, file_.begin() + file_.find('_', file_.find('_', file_.find('_') + 1) + 1));
        std::cout << temp_fns->fgrade << " " << temp_fns->fclass << " " << temp_fns->fterms << std::endl;
        // std::cout << temp_fns->fgrade << std::endl;
    }

    // 匹配文件的思路：用string的rfind()或find_last_one()函数找到'/'出现的最后位置（详细说明见C++Primer）
    void Excelops::retrieveFile(std::string file_)
    {
        std::string filename_;
        filename_ = getNameproperty(file_);
        assignFilenamepro(filename_);

        doc_temp.open(file_);
        wks_temp = doc_temp.workbook().sheet(1).get<OpenXLSX::XLWorksheet>();

        wks_total = doc_total.workbook().sheet(temp_fns->fgrade + "_" + temp_fns->fclass).get<OpenXLSX::XLWorksheet>();

        std::cout << temp_fns->fclass << "班总表行数: "<< getrowCount(wks_total) << std::endl;
        std::cout << temp_fns->fclass << "班第" << temp_fns->fterms << "期待检索表行数: "<< getrowCount(wks_temp) << std::endl;
        /**
         * @brief 此处添加exception：若待检索表行数大于对应的总表行数则抛出一个错误
         * 
         */

        writeValues_total.resize(getrowCount(wks_total) - STARTROW_TOTAL + 1);
        if (initWritevector(writeValues_total, getrowCount(wks_total) - STARTROW_TOTAL + 1))
            std::cout << "初始化写入向量成功！向量长度为：" << getrowCount(wks_total) - STARTROW_TOTAL + 1 << std::endl;

        fillWriteValues(CHECKEDNAMECOL, LOADEDNAMECOL);

        int termnumber = 0;
        std::istringstream ss(temp_fns->fterms);
        ss >> termnumber;
        std::cout << "期数：" << termnumber << std::endl;

        loadData(termnumber);

        doc_temp.close();
    }

    /**
     * @brief 初始化属性值
     * 
     * @param obj_ 
     * @param files_ 
     * @param temp_start_row_ 
     * @param total_start_row_ 
     * @return true 
     * @return false 
     */
    bool Excelops::init(std::string obj_, std::vector<std::string>& files_, int temp_start_row_, int total_start_row_)
    {
        total_path = obj_;
        files.assign(files_.begin(), files_.end());

        temp_start_row = temp_start_row_;
        total_start_row = total_start_row_;

        // 一定要初始化
        temp_fns = new FileNamesplit;

        doc_total.open(total_path);

        return true;
    }

    void Excelops::fillWriteValues(int checkedNamecol, int loadedNamecol)
    {
        // c++11新特性：foreach
        for (auto& row : wks_temp.rows(STARTROW_TEMP, getrowCount(wks_temp))) {
            readValues_temp = row.values();
            if (readValues_temp.size() != 0) {
                int index = 0;
                for (auto& row_m : wks_total.rows(STARTROW_TOTAL, getrowCount(wks_total))) {
                    readValues_total = row_m.values();
                    
                    if (readValues_total.size() != 0) {
                        // std::cout << readValues_temp[checkedNamecol- 1] << " " << readValues_total[loadedNamecol- 1] << std::endl;
                        if (readValues_temp[checkedNamecol - 1].get<std::string>().compare(readValues_total[loadedNamecol - 1].get<std::string>()) == 0) {
                            writeValues_total[index] = 1; 
                        }
                    }
                    index++;
                }
            }
        }
    }

    void Excelops::loadData(int termnumber)
    {
        std::cout << "                    正在填入数据，请稍等..." << std::endl;
        std::cout << "                    填入数据量为：" << " ";
        std::cout << writeValues_total.size() << std::endl;
        wks_total.cell("H3").value() = 1;
        for (int i = 0; i < writeValues_total.size(); i++) {
            wks_total.cell(STARTROW_TOTAL + i, STARTCOL_TOTAL + termnumber - 1).value() = writeValues_total[i];
            // std::cout << writeValues_total[i] << std::endl;
        }

        std::cout << "                    完成!" << std::endl;
    }

    Excelops::~Excelops()
    {
        doc_total.save();
        doc_total.close();

        readValues_total.clear();
        writeValues_total.clear();
        readValues_temp.clear();

        delete temp_fns;

    }
}