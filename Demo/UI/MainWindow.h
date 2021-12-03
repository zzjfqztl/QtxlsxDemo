/**
 * 主窗体部分
 */

#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

class ComboBoxFrame;
class QTextEdit;
class QCheckBox;
class QPushButton;
class FileSelectFrame;
class QLineEdit;
class MainWindow : public QMainWindow
{
    Q_OBJECT
public:
    explicit MainWindow(QWidget *parent = nullptr);

public slots:
    void onImportExcelClick();
    void onSaveExcelClick();

private:
    /**
     * @brief 主界面初始化
    */
    void InitUI();

private:
    QPushButton* btnImportExcel_ = nullptr;
    FileSelectFrame* fileDirFrame_ = nullptr;
    QPushButton* btnSaveExcel_ = nullptr;
    QTextEdit* textEdit_ = nullptr;
    
};

#endif // MAINWINDOW_H
