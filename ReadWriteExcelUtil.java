package com.sbt.tool;

import java.io.File;
import java.io.IOException;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ReadWriteExcelUtil {

	/**
	 * @param args
	 */
	public static void implEXL(HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		String fileName = "E:" + File.separator + "OA����.xls";
		System.out.println(ReadWriteExcelUtil.readExcel(fileName));
		String fileName1 = "E:" + File.separator + "abc.xls";
		ReadWriteExcelUtil.writeExcel(fileName1);
	}

	/**
	 * ��excel�ļ����xȡ���еă���
	 * 
	 * @param file
	 *            excel�ļ�
	 * @return excel�ļ��ă���
	 */
	public static String readExcel(String fileName) {
		StringBuffer sb = new StringBuffer();
		Workbook wb = null;
		try {
			// ����Workbook��������������
			wb = Workbook.getWorkbook(new File(fileName));
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		if (wb == null)
			return null;

		// �����Workbook����֮�󣬾Ϳ���ͨ�����õ�Sheet��������������
		Sheet[] sheet = wb.getSheets();

		if (sheet != null && sheet.length > 0) {
			// ��ÿ�����������ѭ��
			for (int i = 0; i < sheet.length; i++) {
				// �õ���ǰ�����������
				int rowNum = sheet[i].getRows();
				for (int j = 0; j < rowNum; j++) {
					// �õ���ǰ�е����е�Ԫ��
					Cell[] cells = sheet[i].getRow(j);
					if (cells != null && cells.length > 0) {
						// ��ÿ����Ԫ�����ѭ��
						for (int k = 0; k < cells.length; k++) {
							// ��ȡ��ǰ��Ԫ���ֵ
							String cellValue = cells[k].getContents();
							sb.append(cellValue + "\t");
						}
					}
					sb.append("\r\n");
				}
				sb.append("\r\n");
			}
		}
		// ���ر���Դ���ͷ��ڴ�
		wb.close();
		return sb.toString();
	}

	/**
	 * �у��݌���excel�ļ���
	 * 
	 * @param fileName
	 *            Ҫ������ļ������Q
	 */
	public static void writeExcel(String fileName) {
		WritableWorkbook wwb = null;
		try {
			// ����Ҫʹ��Workbook��Ĺ�����������һ����д��Ĺ�����(Workbook)����
			wwb = Workbook.createWorkbook(new File(fileName));
		} catch (IOException e) {
			e.printStackTrace();
		}
		if (wwb != null) {
			// ����һ����д��Ĺ�����
			// Workbook��createSheet������������������һ���ǹ���������ƣ��ڶ����ǹ������ڹ������е�λ��
			WritableSheet ws = wwb.createSheet("sheet1", 0);

			// ���濪ʼ��ӵ�Ԫ��
			for (int i = 0; i < 10; i++) {
				for (int j = 0; j < 5; j++) {
					// ������Ҫע����ǣ���Excel�У���һ��������ʾ�У��ڶ�����ʾ��
					Label labelC = new Label(j, i, "���ǵ�" + (i + 1) + "�У���"
							+ (j + 1) + "��");
					try {
						// �����ɵĵ�Ԫ����ӵ���������
						ws.addCell(labelC);
					} catch (RowsExceededException e) {
						e.printStackTrace();
					} catch (WriteException e) {
						e.printStackTrace();
					}

				}
			}

			try {
				// ���ڴ���д���ļ���
				wwb.write();
				// �ر���Դ���ͷ��ڴ�
				wwb.close();
			} catch (IOException e) {
				e.printStackTrace();
			} catch (WriteException e) {
				e.printStackTrace();
			}
		}
	}

}