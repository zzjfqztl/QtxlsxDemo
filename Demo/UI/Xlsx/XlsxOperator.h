/**
 * Excel 文件操作类，目前只支持xlsx后缀的excel
 */
#ifndef XLSXOPERATOR_H
#define XLSXOPERATOR_H

#include <QObject>


class XlsxOperator : public QObject
{
    Q_OBJECT
public:

    /**
     * @brief 构造函数
    */
    explicit XlsxOperator(QObject*parent = nullptr);

    /**
     * @brief 保存数据到excel中，并保存到本地
     * @param excelFile 保存到本地的excel文件全路径
     * @param headers 标题列表，也即excel文件第一行显示的内容
     * @param datas excel 数据
     * @param sheetName 工作簿名称，空为默认名称
     * @return true 保存成功 false 失败
    */
    bool writeExcel(const QString& excelFile, const QStringList& headers, const QList<QStringList>& datas, const QString& sheetName = QString());

	/**
	 * @brief 保存数据到excel中，并将excel并保存到本地，适用于保存多个sheet（工作簿）
	 * @param excelFile 保存到本地的excel文件全路径
	 * @param headers 标题列表，也即excel文件每个sheet第一行显示的内容
	 * @param datas excel 文件每个sheet写入的内容
	 * @param sheetNames excel 多个工作簿对应的名称
	 * @return true 保存成功 false 失败
	*/
	bool writeExcel(const QString& excelFile, const QList<QStringList>& headers, const QList<QList<QStringList>>& datas, const QStringList& sheetNames = QStringList());

    /**
     * @brief 读取指定excel文件 第一个sheet中的内容 
     * @param excelFile 指定excel文件
     * @param headers 读取到的头信息，也即标题信息
     * @param datas 读取到的数据信息
     * @return int 0 成功，-1 文件不存在，-2 没有sheet文件 -3 数据不存在
    */
    int readExcel(const QString& excelFile, QStringList& headers, QList<QStringList>& datas);

    /**
     * @brief 读取指定excel文件所有的内容 可能还有多个sheet
     * @param excelFile 指定excel文件
     * @param headers 读取到的头信息，也即标题信息
     * @param datas 读取到的数据信息
     * @return int 0 成功，-1 文件不存在，-2 没有sheet文件 -3 数据不存在
    */
    int readExcel(const QString& excelFile, QList<QStringList>& headers, QList<QList<QStringList>>& datas);
};

#endif // XLSXOPERATOR_H
