#include "XlsxOperator.h"

#include "QtXlsx/xlsxdocument.h"
#include "QtXlsx/xlsxformat.h"
#include "QtXlsx/xlsxcellrange.h"
#include "QtXlsx/xlsxworksheet.h"
#include "QFile"

QTXLSX_USE_NAMESPACE


/**
 * @brief ����excel����ĳ��Ԫ���ı�ˮƽ���뷽ʽ
 * @param xlsx Document ����
 * @param cell ĳ��Ԫ������
 * @param text ��Ԫ����ʾ���ı�
 * @param align �ı�ˮƽ���뷽ʽ
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
 * @brief ����excel����ĳЩ��Ԫ���ı���ֱ���뷽ʽ
 * @param xlsx Document ����
 * @param range ��Ԫ���飬�����Ԫ����"D3:D7"��ʽ����
 * @param text ��ʾ���ı�
 * @param align ��ֱ���뷽ʽ
*/
void writeVerticalAlignCell(Document& xlsx, const QString& range, const QString& text,
	Format::VerticalAlignment align)
{
	Format format;
	format.setVerticalAlignment(align);
	format.setBorderStyle(Format::BorderThin);
	CellRange r(range);
	xlsx.write(r.firstRow(), r.firstColumn(), text);
	xlsx.mergeCells(r, format);//�ϲ�
}

/**
 * @brief ���õ�Ԫ��߿���ʽ
 * @param xlsx Document ����
 * @param cell ָ����Ԫ������
 * @param text ��Ԫ����ʾ���ı�
 * @param bs �߿���ʽ
*/
void writeBorderStyleCell(Document& xlsx, const QString& cell, const QString& text,
	Format::BorderStyle bs)
{
	Format format;
	format.setBorderStyle(bs);
	xlsx.write(cell, text, format);
}

/**
 * @brief ���õ�Ԫ��������ɫ
 * @param xlsx Document ����
 * @param cell ָ����Ԫ������
 * @param color ������ɫֵ
*/
void writeSolidFillCell(Document& xlsx, const QString& cell, const QColor& color)
{
	Format format;
	format.setPatternBackgroundColor(color);
	xlsx.write(cell, QVariant(), format);
}

/**
 * @brief ����ָ����Ԫ���ͼ������ɫ
 * @param xlsx Document ����
 * @param cell ָ����Ԫ������
 * @param pattern ͼ����ʽ
 * @param color ͼ��������ɫֵ
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
 * @brief ����ָ����Ԫ��߿���ɫ��������ɫ
 * @param xlsx Document ����
 * @param cell ָ����Ԫ������
 * @param text ��Ԫ����ʾ�ı�
 * @param color �߿�������ɫ
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
 * @brief ���õ�Ԫ��ʹ�õ�����
 * @param xlsx Document ����
 * @param cell ָ����Ԫ��
 * @param text ��������
*/
void writeFontNameCell(Document& xlsx, const QString& cell, const QString& text)
{
	Format format;
	format.setFontName(text);
	format.setFontSize(16);
	xlsx.write(cell, text, format);
}

/**
 * @brief ���õ�Ԫ�������С
 * @param xlsx Document ����
 * @param cell ָ����Ԫ��
 * @param size �����С
*/
void writeFontSizeCell(Document& xlsx, const QString& cell, int size)
{
	Format format;
	format.setFontSize(size);
	xlsx.write(cell, "Qt Xlsx", format);
}

/**
 * @brief ����ָ����Ԫ���������ݸ�ʽ����Ϣ
 * @param xlsx Document ����
 * @param row ָ����Ԫ��
 * @param value ��Ҫ��ʽ����ֵ
 * @param numFmt ��ʽ������
 * @param numFmt ��ʽ������
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
 * @brief ����ָ����Ԫ���������ݸ�ʽ����Ϣ
 * @param xlsx Document ����
 * @param row ָ����Ԫ��
 * @param value ��Ҫ��ʽ����ֵ
 * @param numFmt ��ʽ���ֶ�
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
	QXlsx::Document xlsx;//����xlsx�ļ�
	QStringList list = xlsx.sheetNames();

	for (int i = 0; i < list.size(); ++i)//ɾ������sheet
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
	QXlsx::Document xlsx;//����xlsx�ļ�
	QStringList list = xlsx.sheetNames();

	for (int i = 0; i < list.size(); ++i)//ɾ������sheet
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
		return -1; //�ļ�������
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
		return -1; //�ļ�������
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
