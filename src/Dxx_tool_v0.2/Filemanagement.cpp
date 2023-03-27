#include "Filemanagement.h"

namespace DXXTL {
    FileManagement::FileManagement() {}
    FileManagement::~FileManagement()
    {
        files_.clear();
    }
    void FileManagement::setPath(std::string path_) { path = path_; }
    void FileManagement::setFilepath(std::string path_) { totalfile_path = path_; }
    std::string FileManagement::getFilepath() { return totalfile_path; }
    void FileManagement::getFileNames(std::string path_, std::vector<std::string>& files)
    {
        files.clear();
        intptr_t hFile = 0;
        struct _finddata_t fileinfo;
        setPath(path_);
        std::string p;
        if ((hFile = _findfirst(p.assign(path_).append("/*").c_str(), &fileinfo)) != -1)
	    {
		    do
		    {
                //如果是目录,递归查找
                //如果不是,把文件绝对路径存入vector中
                if ((fileinfo.attrib & _A_SUBDIR))
                {
                    if (strcmp(fileinfo.name, ".") != 0 && strcmp(fileinfo.name, "..") != 0)
                        getFileNames(p.assign(path_).append("/").append(fileinfo.name), files);
                }
                else
                {
                    files.push_back(p.assign(path_).append("/").append(fileinfo.name));
                }
            } while (_findnext(hFile, &fileinfo) == 0);
            _findclose(hFile);
	    } else {
            std::cout << "此目录下无文件！" << std::endl;
        }
    }
    std::vector<std::string>& FileManagement::searchFiles(std::string path_)
    {
        getFileNames(path_, files_);
        return files_;
    }
    void FileManagement::printFiles()
    {
        std::vector<std::string>::iterator it;
        for(it = files_.begin(); it != files_.end(); it++)
        {
            std::cout << *it << std::endl;
        }
    }

    std::vector<std::string>& FileManagement::getFiles()
    {
        return files_;
    }

}