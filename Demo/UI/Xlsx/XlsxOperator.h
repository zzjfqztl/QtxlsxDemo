/**
 * Excel �ļ������࣬Ŀǰֻ֧��xlsx��׺��excel
 */
#ifndef XLSXOPERATOR_H
#define XLSXOPERATOR_H

#include <QObject>


class XlsxOperator : public QObject
{
    Q_OBJECT
public:

    /**
     * @brief ���캯��
    */
    explicit XlsxOperator(QObject*parent = nullptr);

    /**
     * @brief �������ݵ�excel�У������浽����
     * @param excelFile ���浽���ص�excel�ļ�ȫ·��
     * @param headers �����б�Ҳ��excel�ļ���һ����ʾ������
     * @param datas excel ����
     * @param sheetName ���������ƣ���ΪĬ������
     * @return true ����ɹ� false ʧ��
    */
    bool writeExcel(const QString& excelFile, const QStringList& headers, const QList<QStringList>& datas, const QString& sheetName = QString());

	/**
	 * @brief �������ݵ�excel�У�����excel�����浽���أ������ڱ�����sheet����������
	 * @param excelFile ���浽���ص�excel�ļ�ȫ·��
	 * @param headers �����б�Ҳ��excel�ļ�ÿ��sheet��һ����ʾ������
	 * @param datas excel �ļ�ÿ��sheetд�������
	 * @param sheetNames excel �����������Ӧ������
	 * @return true ����ɹ� false ʧ��
	*/
	bool writeExcel(const QString& excelFile, const QList<QStringList>& headers, const QList<QList<QStringList>>& datas, const QStringList& sheetNames = QStringList());

    /**
     * @brief ��ȡָ��excel�ļ� ��һ��sheet�е����� 
     * @param excelFile ָ��excel�ļ�
     * @param headers ��ȡ����ͷ��Ϣ��Ҳ��������Ϣ
     * @param datas ��ȡ����������Ϣ
     * @return int 0 �ɹ���-1 �ļ������ڣ�-2 û��sheet�ļ� -3 ���ݲ�����
    */
    int readExcel(const QString& excelFile, QStringList& headers, QList<QStringList>& datas);

    /**
     * @brief ��ȡָ��excel�ļ����е����� ���ܻ��ж��sheet
     * @param excelFile ָ��excel�ļ�
     * @param headers ��ȡ����ͷ��Ϣ��Ҳ��������Ϣ
     * @param datas ��ȡ����������Ϣ
     * @return int 0 �ɹ���-1 �ļ������ڣ�-2 û��sheet�ļ� -3 ���ݲ�����
    */
    int readExcel(const QString& excelFile, QList<QStringList>& headers, QList<QList<QStringList>>& datas);
};

#endif // XLSXOPERATOR_H
