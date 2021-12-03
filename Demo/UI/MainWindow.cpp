#include "MainWindow.h"
#include "QBoxLayout"
#include "QPushButton"
#include "Common/CustomControls.h"
#include <QTextEdit>
#include <QMessageBox>

#include "Common/MaskWidget.h"
#include "QDateTime"
#include "QLabel"
#include "QLineEdit"
#include "QFileInfo"
#include "QFileDialog"
#include "Dialog/MessageBoxDialog.h"
#include "Xlsx/XlsxOperator.h"


MainWindow::MainWindow(QWidget *parent) : QMainWindow(parent)
{
	InitUI();
	showMaximized();
}

void MainWindow::onImportExcelClick()
{
	QFileDialog dlg;
	QString filter = QString("Excel(*.xlsx *.xls)");
	QString strExcelName = dlg.getOpenFileName(nullptr, QObject::tr("Open File"), QObject::tr("./"), filter);
	if (strExcelName.isEmpty())
	{
		return;
	}

	
	//excel 中第一个工作簿
	/*{
		QStringList headers;
		QList<QStringList> datas;

		XlsxOperator xlsxOpt;
		int nError = xlsxOpt.readExcel(strExcelName, headers, datas);
		textEdit_->append(QString("%1 read %2").arg(strExcelName).arg((nError == 0 ? QString("success") : QString("filed"))));
		textEdit_->append(QString("error code = %1").arg(nError));

		if (nError == 0)
		{
			textEdit_->append(headers.join("       "));
			for (auto& itr : datas)
			{
				textEdit_->append(itr.join("       "));
			}
		}
	}*/

	//excel 所有工作簿内容
	QList<QStringList> headers;
	QList < QList<QStringList>> datas;

	XlsxOperator xlsxOpt;
	int nError = xlsxOpt.readExcel(strExcelName, headers, datas);
	textEdit_->append(QString("%1 read %2").arg(strExcelName).arg((nError == 0 ? QString("success") : QString("filed"))));
	textEdit_->append(QString("error code = %1").arg(nError));

	if (nError == 0)
	{
		int nCount = 0;
		for (auto &itr : headers)
		{
			textEdit_->append(itr.join("       "));
			if (datas.count() > nCount)
			{
				for (auto& itr : datas[nCount])
				{
					textEdit_->append(itr.join("       "));
				}
			}
			
			nCount++;
		}
		if (nCount  < datas.count())
		{
			for (int iLoop = nCount; iLoop < datas.count(); iLoop++)
			{
				for (auto& itr : datas[iLoop])
				{
					textEdit_->append(itr.join("       "));
				}
			}
		}
		
		
	}

}

void MainWindow::onSaveExcelClick()
{
	if (fileDirFrame_->GetFileList().count() < 0)
	{
		MessageBoxDialog* msgDlg = new MessageBoxDialog(QObject::tr("Information"), QObject::tr("Please choose save contents"));
		msgDlg->exec();
		delete msgDlg;
		msgDlg = nullptr;
		return;
	}
	QFileDialog dlg;
	QString filter = QString("Excel(*.xlsx)");
	QString strSaveName = dlg.getSaveFileName(nullptr, QObject::tr("Save File"), QObject::tr("./"), filter);
	if (strSaveName.isEmpty())
	{
		return;
	}

	QDir dir = fileDirFrame_->GetFileList()[0];

	//保存单个sheet
	/*
	{
		QStringList headers;
		QList<QStringList> datas;

		headers << QString("File Name") << QString("Suffix") << QString("Size") << QString("Last modify Time") << QString("File Path");
		QStringList strList;
		if (dir.exists())
		{
			QFileInfoList fileList = dir.entryInfoList(QDir::AllEntries | QDir::NoDotAndDotDot);
			for (auto& itr : fileList)
			{
				strList.clear();
				strList << itr.fileName() << itr.suffix() << QString("%1").arg(itr.size()) << itr.lastModified().toString("yyyy-MM-dd hh:mm:ss") << itr.absoluteFilePath();
				datas.append(strList);
			}
		}
		
		XlsxOperator xlsxOpt;
		bool b = xlsxOpt.writeExcel(strSaveName, headers, datas);
		textEdit_->append(QString("%1 save %2").arg(strSaveName).arg((b ? QString("success") : QString("filed"))));
	}
	*/
	//保存多个sheet
	QList<QStringList> headers;
	QList < QList<QStringList>> datas;

	QStringList header;
	header << QString("File Name") << QString("Suffix") << QString("Size") << QString("Last modify Time") << QString("File Path");
	headers.append(header);
	QList<QStringList> data;
	QStringList strList;
	if (dir.exists())
	{
		QFileInfoList fileList = dir.entryInfoList(QDir::AllEntries | QDir::NoDotAndDotDot);
		for (auto &itr : fileList)
		{
			strList.clear();
			strList << itr.fileName() << itr.suffix() << QString("%1").arg(itr.size()) << itr.lastModified().toString("yyyy-MM-dd hh:mm:ss") << itr.absoluteFilePath();
			data.append(strList);
		}
	}
	datas.append(data);
	if (dir.exists())
	{
		QFileInfoList dirList = dir.entryInfoList(QDir::Dirs | QDir::NoDotAndDotDot);
		for (auto& itr : dirList)
		{
			if (itr.exists() && itr.isDir())
			{
				headers.append(header);
				data.clear();
				strList.clear();
				QFileInfoList fileList = QDir(itr.absoluteFilePath()).entryInfoList(QDir::AllEntries | QDir::NoDotAndDotDot);
				for (auto& citr : fileList)
				{
					strList.clear();
					strList << citr.fileName() << citr.suffix() << QString("%1").arg(citr.size()) << citr.lastModified().toString("yyyy-MM-dd hh:mm:ss") << citr.absoluteFilePath();
					data.append(strList);
				}
				datas.append(data);
			}
			
		}
	}
	XlsxOperator xlsxOpt;
	bool b = xlsxOpt.writeExcel(strSaveName, headers, datas);
	textEdit_->append(QString("%1 save %2").arg(strSaveName).arg((b ? QString("success") : QString("filed"))));
}


void MainWindow::InitUI()
{
	QWidget* centerWidget_ = new QWidget(this);
	centerWidget_->setObjectName("centerWidget");
	setCentralWidget(centerWidget_);

	QVBoxLayout* vLayout = new QVBoxLayout();
	vLayout->setSpacing(10);
	vLayout->setMargin(0);
	vLayout->setContentsMargins(20, 10, 20, 10);
	vLayout->setAlignment(Qt::AlignTop);

	centerWidget_->setLayout(vLayout);

	QHBoxLayout* hlayout = new QHBoxLayout();
	hlayout->setSpacing(20);
	hlayout->setMargin(0);
	vLayout->addLayout(hlayout);
	//导出文件存储路径选择
	fileDirFrame_ = new FileSelectFrame(QObject::tr("Save Path"), this, true);
	hlayout->addWidget(fileDirFrame_);

	btnSaveExcel_ = new QPushButton(this);
	btnSaveExcel_->setObjectName("btnSaveExcel");
	btnSaveExcel_->setText("Save Excel");
	btnSaveExcel_->setFocusPolicy(Qt::NoFocus);
	connect(btnSaveExcel_, &QPushButton::clicked, this, &MainWindow::onSaveExcelClick);
	hlayout->addWidget(btnSaveExcel_);

	btnImportExcel_ = new QPushButton(this);
	btnImportExcel_->setObjectName("btnImportExcel");
	btnImportExcel_->setText("Import Excel");
	btnImportExcel_->setFocusPolicy(Qt::NoFocus);
	connect(btnImportExcel_, &QPushButton::clicked, this, &MainWindow::onImportExcelClick);
	hlayout->addWidget(btnImportExcel_);
	hlayout->addStretch();

	textEdit_ = new QTextEdit(this);
	textEdit_->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Expanding);
	textEdit_->setObjectName("textEdit");
	vLayout->addWidget(textEdit_);
}
