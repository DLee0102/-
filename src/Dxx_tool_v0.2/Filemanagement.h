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
        void setPath(std::string path_);
        void getFileNames(std::string path_, std::vector<std::string>& files);
        std::vector<std::string>& searchFiles(std::string path_);
        std::vector<std::string>& getFiles();
        void printFiles();
        void settFilepath(std::string path_);
        std::string gettFilepath();
    private:
        std::string path;
        std::string totalfile_path;
        std::vector<std::string> files_;
    };
}