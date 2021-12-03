#include "XlsxOperator.h"

#include "QtXlsx/xlsxdocument.h"
#include "QtXlsx/xlsxformat.h"
#include "QtXlsx/xlsxcellrange.h"
#include "QtXlsx/xlsxworksheet.h"
#include "QFile"

QTXLSX_USE_NAMESPACE


/**
 * @brief 设置excel表中某单元格文本水平对齐方式
 * @param xlsx Document 对象
 * @param cell 某单元格名称
 * @param text 单元格显示的文本
 * @param align 文本水平对齐方式
*/
void writeHorizontalAlignCell(Document& xlsx, const QString& cell, const QString& text,
	Format::HorizontalAlignment align)
{
	Format format;
	format.setHorizontalAlignment(align);
	format.setBorderStyle(Format::BorderThin);
	xlsx.write(cell, text, format);
}

/**
 * @brief 设置excel表中某些单元格文本垂直对齐方式
 * @param xlsx Document 对象
 * @param range 单元格组，多个单元格以"D3:D7"格式传入
 * @param text 显示的文本
 * @param align 垂直对齐方式
*/
void writeVerticalAlignCell(Document& xlsx, const QString& range, const QString& text,
	Format::VerticalAlignment align)
{
	Format format;
	format.setVerticalAlignment(align);
	format.setBorderStyle(Format::BorderThin);
	CellRange r(range);
	xlsx.write(r.firstRow(), r.firstColumn(), text);
	xlsx.mergeCells(r, format);//合并
}

/**
 * @brief 设置单元格边框样式
 * @param xlsx Document 对象
 * @param cell 指定单元格名称
 * @param text 单元格显示的文本
 * @param bs 边框样式
*/
void writeBorderStyleCell(Document& xlsx, const QString& cell, const QString& text,
	Format::BorderStyle bs)
{
	Format format;
	format.setBorderStyle(bs);
	xlsx.write(cell, text, format);
}

/**
 * @brief 设置单元格填充的颜色
 * @param xlsx Document 对象
 * @param cell 指定单元格名称
 * @param color 填充的颜色值
*/
void writeSolidFillCell(Document& xlsx, const QString& cell, const QColor& color)
{
	Format format;
	format.setPatternBackgroundColor(color);
	xlsx.write(cell, QVariant(), format);
}

/**
 * @brief 设置指定单元格的图案及颜色
 * @param xlsx Document 对象
 * @param cell 指定单元格名称
 * @param pattern 图案样式
 * @param color 图案填充的颜色值
*/
void writePatternFillCell(Document& xlsx, const QString& cell, Format::FillPattern pattern,
	const QColor& color)
{
	Format format;
	format.setPatternForegroundColor(color);
	format.setFillPattern(pattern);
	xlsx.write(cell, QVariant(), format);
}

/**
 * @brief 设置指定单元格边框颜色及字体颜色
 * @param xlsx Document 对象
 * @param cell 指定单元格名称
 * @param text 单元格显示文本
 * @param color 边框及字体颜色
*/
void writeBorderAndFontColorCell(Document& xlsx, const QString& cell, const QString& text,
	const QColor& color)
{
	Format format;
	format.setBorderStyle(Format::BorderThin);
	format.setBorderColor(color);
	format.setFontColor(color);
	xlsx.write(cell, text, format);
}

/**
 * @brief 设置单元格使用的字体
 * @param xlsx Document 对象
 * @param cell 指定单元格
 * @param text 字体名称
*/
void writeFontNameCell(Document& xlsx, const QString& cell, const QString& text)
{
	Format format;
	format.setFontName(text);
	format.setFontSize(16);
	xlsx.write(cell, text, format);
}

/**
 * @brief 设置单元格字体大小
 * @param xlsx Document 对象
 * @param cell 指定单元格
 * @param size 字体大小
*/
void writeFontSizeCell(Document& xlsx, const QString& cell, int size)
{
	Format format;
	format.setFontSize(size);
	xlsx.write(cell, "Qt Xlsx", format);
}

/**
 * @brief 设置指定单元格数字内容格式化信息
 * @param xlsx Document 对象
 * @param row 指定单元格
 * @param value 需要格式化的值
 * @param numFmt 格式化类型
 * @param numFmt 格式化类型
*/
void writeInternalNumFormatsCell(Document& xlsx, int row, double value, int numFmt)
{
	Format format;
	format.setNumberFormatIndex(numFmt);
	xlsx.write(row, 1, value);
	xlsx.write(row, 2, QString("Builtin NumFmt %1").arg(numFmt));
	xlsx.write(row, 3, value, format);
}

/**
 * @brief 设置指定单元格数字内容格式化信息
 * @param xlsx Document 对象
 * @param row 指定单元格
 * @param value 需要格式化的值
 * @param numFmt 格式化字段
*/
void writeCustomNumFormatsCell(Document& xlsx, int row, double value, const QString& numFmt)
{
	Format format;
	format.setNumberFormat(numFmt);
	xlsx.write(row, 1, value);
	xlsx.write(row, 2, numFmt);
	xlsx.write(row, 3, value, format);
}

XlsxOperator::XlsxOperator(QObject* parent /*= nullptr*/)
{

}

bool XlsxOperator::writeExcel(const QString& excelFile, const QStringList& headers, const QList<QStringList>& datas, const QString& sheetName)
{
	QXlsx::Document xlsx;//创建xlsx文件
	QStringList list = xlsx.sheetNames();

	for (int i = 0; i < list.size(); ++i)//删除所有sheet
	{
		xlsx.deleteSheet(list[i]);
	}
	
	QString strSheet = sheetName;
	if (strSheet.isEmpty())
	{
		strSheet = QString("Sheet1");
	}
	
	xlsx.addSheet(strSheet);
	xlsx.selectSheet(strSheet);
	int nRow = 1;

	for (int iLoop = 0; iLoop < headers.count(); iLoop++)
	{
		xlsx.write(nRow, iLoop + 1, headers[iLoop]);
	}
	nRow++;
	int nIndex = 1;
	for (auto &itr : datas)
	{
		nIndex = 1;
		for (auto& citr : itr)
		{
			xlsx.write(nRow, nIndex++, citr);
		}
		nRow++;
	}
	
	xlsx.saveAs(excelFile);
	return true;
}


bool XlsxOperator::writeExcel(const QString& excelFile, const QList<QStringList>& headers, const QList<QList<QStringList>>& datas, const QStringList& sheetNames /*= QStringList()*/)
{
	QXlsx::Document xlsx;//创建xlsx文件
	QStringList list = xlsx.sheetNames();

	for (int i = 0; i < list.size(); ++i)//删除所有sheet
	{
		xlsx.deleteSheet(list[i]);
	}
	int nSheetIndex = 1;
	for (int iLoop = 0; iLoop < headers.count();iLoop++)
	{
		QString strSheet;
		if (sheetNames.count() > iLoop)
		{
			strSheet = sheetNames[iLoop];
		}
		if (strSheet.isEmpty())
		{
			strSheet = QString("Sheet%1").arg(nSheetIndex);
			nSheetIndex++;
		}

		xlsx.addSheet(strSheet);
		xlsx.selectSheet(strSheet);
		int nRow = 1;

		for (int hLoop = 0; hLoop < headers[iLoop].count(); hLoop++)
		{
			xlsx.write(nRow, hLoop + 1, headers[iLoop][hLoop]);
		}
		nRow++;
		int nIndex = 1;
		if (datas.count() > iLoop)
		{
			for (auto& itr : datas[iLoop])
			{
				nIndex = 1;
				for (auto& citr : itr)
				{
					xlsx.write(nRow, nIndex++, citr);
				}
				nRow++;
			}
		}
		

	}
	
	xlsx.saveAs(excelFile);
	return true;
}

int XlsxOperator::readExcel(const QString& excelFile, QStringList& headers, QList<QStringList>& datas)
{
	if (!QFile::exists(excelFile))
	{
		return -1; //文件不存在
	}
	Document xlsx(excelFile);
	QStringList list = xlsx.sheetNames();
	if (list.count() <= 0)
	{
		return -2;
	}
	xlsx.selectSheet(list[0]);
	Worksheet* sheet = xlsx.currentWorksheet();
	int rowTotal = sheet->dimension().rowCount();
	int colTotal = sheet->dimension().columnCount();
	if (rowTotal <1 || colTotal < 1)
	{
		return -3;
	}
	int nRow = 1;
	for (int col = 1; col <= colTotal; col++)
	{
		headers.append(sheet->read(nRow, col).toString());
	}
	nRow++;
	QStringList strList;
	for (int row = nRow; row <= rowTotal; row++)
	{
		strList.clear();
		for (int col = 1; col <= colTotal; col++)
		{
			strList.append(sheet->read(row, col).toString());
		}
		datas.append(strList);
	}
	return 0;
}

int XlsxOperator::readExcel(const QString& excelFile, QList<QStringList>& headers, QList<QList<QStringList>>& datas)
{
	if (!QFile::exists(excelFile))
	{
		return -1; //文件不存在
	}
	Document xlsx(excelFile);
	QStringList list = xlsx.sheetNames();
	if (list.count() <= 0)
	{
		return -2;
	}
	QStringList strHeaders;
	QStringList strList;
	QList<QStringList> strDatas;
	for (auto &itr : list)
	{
		strHeaders.clear();
		strList.clear();
		strDatas.clear();
		xlsx.selectSheet(itr);
		Worksheet* sheet = xlsx.currentWorksheet();
		int rowTotal = sheet->dimension().rowCount();
		int colTotal = sheet->dimension().columnCount();
		if (rowTotal < 1 || colTotal < 1)
		{
			return -3;
		}
		int nRow = 1;
		for (int col = 1; col <= colTotal; col++)
		{
			strHeaders.append(sheet->read(nRow, col).toString());
		}
		headers.append(strHeaders);
		nRow++;
		for (int row = nRow; row <= rowTotal; row++)
		{
			strList.clear();
			for (int col = 1; col <= colTotal; col++)
			{
				strList.append(sheet->read(row, col).toString());
			}
			strDatas.append(strList);
		}
		datas.append(strDatas);
	}
	
	return 0;
}
