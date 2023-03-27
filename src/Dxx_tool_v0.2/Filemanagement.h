#include <iostream>
#include <io.h>
#include <string>
#include <vector>

namespace DXXTL {
    class FileManagement 
    {
    public:
        FileManagement();
        ~FileManagement();
        std::vector<std::string>& searchFiles(std::string path_);
        std::vector<std::string>& getFiles();
        void printFiles();
        void setFilepath(std::string path_);   // 设置总文件路径
        std::string getFilepath();
    private:
        std::string path;   
        std::string totalfile_path;     // 总文件路径
        std::vector<std::string> files_;    // 待检索文件夹路径

        void getFileNames(std::string path_, std::vector<std::string>& files);
        void setPath(std::string path_);    // 被getFileNames调用
    };
}