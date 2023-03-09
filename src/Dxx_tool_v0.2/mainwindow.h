#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QFileDialog>
#include <QMessageBox>
#include <QDebug>
#include "Core.h"
#include "Filemanagement.h"

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void openfolderSlot();
    void openfileSlot();
    void startOpenxlsxSlot();

private:
    Ui::MainWindow *ui;
    DXXTL::FileManagement* filemanager;
    DXXTL::Excelops* excleops;
};
#endif // MAINWINDOW_H



