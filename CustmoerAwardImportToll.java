package com.sbt.tool;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Types;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import com.webbuilder.common.Main;
import com.webbuilder.tool.ExcelObject;
import com.webbuilder.utils.DbUtil;
import com.webbuilder.utils.FileUtil;
import com.webbuilder.utils.StringUtil;
import com.webbuilder.utils.SysUtil;
import com.webbuilder.utils.WebUtil;
import com.webbuilder.utils.ZipUtil;

public class CustmoerAwardImportToll {
	public static void getFile(HttpServletRequest request,
			HttpServletResponse response) throws Exception {

		InputStream in = (InputStream) request.getAttribute("uploadFile");
		String fileName = request.getAttribute("uploadFile__name").toString();
		String fileType = fileName.substring(fileName.lastIndexOf(".") + 1,
				fileName.length());
		String imptype = request.getAttribute("imptype").toString();
		Map<String, String> map = new HashMap<String, String>();
		if ("1".equals(imptype)) { // ����-��������

			map.put("�������", "SECHEMNO");
			map.put("��������", "SECHEMNAME");
			map.put("ս��", "ZHANGQ");
			map.put("Ӫ����", "YXB");
			map.put("�����Ʒ��", "PK_PRODLINE");
			map.put("������", "INVBASDOCID");
			map.put("��ʼ����", "STARTMONTH");
			map.put("��������", "ENDMONTH");
			map.put("��������", "COUNTMONTH");
			map.put("���������", "SALETYPE");
			map.put("������ʼ", "SALESTART");
			map.put("��������", "SALEEND");
			map.put("������", "INVBASDOCID");
			map.put("��������", "MONEY_YWY");
			map.put("�����鳤", "MONEY_ZZ");
		}
		if ("2".equals(imptype)) { // ����-�������ڻ�������

			map.put("�������", "SECHEMNO");
			map.put("��������", "SECHEMNAME");
			map.put("ս��", "ZHANGQ");
			map.put("Ӫ����", "YXB");
			map.put("�����Ʒ��", "PK_PRODLINE");
			map.put("��ʼ�·�", "STARTMONTH");
			map.put("�����·�", "ENDMONTH");
			map.put("������", "COUNTMONTH");
			map.put("�ƻ��˲��ƽ���Ʒ", "PK_INVBASDOCGROUP");
			map.put("���������", "SALETYPE");
			map.put("������������δ����·���", "COUNTMONTH");
			map.put("�����ۻر���", "RATE");
		}
		if ("3".equals(imptype)) { // ����-�����ڴ��»�������
			map.put("�������", "SECHEMNO");
			map.put("��������", "SECHEMNAME");
			map.put("ս��", "ZHANGQ");
			map.put("Ӫ����", "YXB");
			map.put("��ʼ�·�", "STARTMONTH");
			map.put("�����·�", "ENDMONTH");
			map.put("������", "COUNTNUMBER");
			map.put("�����Ʒ��", "PK_PRODLINE");
			map.put("�ƻ��˲��ƽ�����Ʒ��ϣ�", "PK_INVBASDOCGROUP");
			map.put("������N��", "COUNTMONTH");
			map.put("���������", "SALETYPE");
			map.put("������������δ������·�����", "COUNTMONTH");
			map.put("�����ۻأ�P��", "RATE");
			map.put("���ڱ�����M��", "LOWERRATE");
		}
		if ("4".equals(imptype)) { // ����-�����ڴ����»�������
			map.put("�������", "SECHEMNO");
			map.put("��������", "SECHEMNAME");
			map.put("ս��", "ZHANGQ");
			map.put("Ӫ����", "YXB");
			map.put("��ʼ�·�", "STARTMONTH");
			map.put("�����·�", "ENDMONTH");
			map.put("������", "COUNTNUMBER");
			map.put("�����Ʒ��", "PK_PRODLINE");
			map.put("�ƻ��˲��ƽ���Ʒ���", "PK_INVBASDOCGROUP");
			map.put("����", "TYPE");
			map.put("���������", "SALETYPE");
			map.put("���˱��", "KHNUMBER");
			map.put("�����ۻر���", "RATE");
			map.put("��ע", "MEMO");
		}
		if ("5".equals(imptype)) { // �ͻ�-��������1-�ͻ��½�
			map.put("�������", "SCHEMENO");
			map.put("��������", "SCHEMENAME");
			map.put("����", "PK_CORP");
			map.put("��������", "TYPE");
			map.put("��ʼ����", "STARTDATE");
			map.put("��������", "ENDDATE");
			map.put("������Ŀ", "DETAILNAME");
			map.put("�ͻ��������", "CUSTOMERGROUP");
			map.put("������׼������", "SCHEMEGROUP");
			map.put("������׼����", "DJTYPE");
			map.put("��Ʒ������", "INVBADOCGROUP");
			map.put("��Ʒ��1����", "INVBADOCGROUP1");
			map.put("��Ʒ��2����", "INVBADOCGROUP2");
			map.put("��Ʒ��3����", "INVBADOCGROUP3");
			map.put("��Ʒ��4����", "INVBADOCGROUP4");
			map.put("������ʽ", "CONDITIONJLFS");
			map.put("�ͻ�������ʽ", "CONDITION0");
			map.put("�����һ�������Ʒ", "CONDITION1");
			map.put("�����һ�ֱ���Ʒ", "CONDITION2");
			map.put("ֱ�����ͷ�����Ʒ", "CONDITION4");
			map.put("ֱ������ֱ���Ʒ", "CONDITION5");
			map.put("������ʼ", "STARTDATE_JS");
			map.put("��������", "ENDDATE_JS");
			map.put("��������", "RATE");
			map.put("��Ʒ�������", "RATE_CONDITION");
		}

		if ("6".equals(imptype)) { // �ͻ�-��������4-ר����ͬ��
			map.put("��ͬ���", "HTNO");
			map.put("��������", "JSZQ");
			map.put("��˾PK", "PK_CORP");
			map.put("��Ʒ", "INVBADOCGROUP");
			map.put("��׼", "SCHEMEGROUP");
			map.put("���㷽ʽ", "JSFS");
			map.put("��������", "MINNNUMBER");
			map.put("��������", "SALETYPE");
		}

		if ("7".equals(imptype)) { // �ͻ�-��������5-ר����ͬ��ǩ�������
			map.put("��ͬ���", "HTNO");
			map.put("��ͬ����", "HTTYPE");
			map.put("�ͻ�����", "CUSTCODE");
			map.put("�ͻ�����", "CUSTNAME");
			map.put("�ƻ���Ч����", "STARTDATE");
			map.put("�ƻ���ֹ����", "ENDDATE");
			map.put("�����һ�˫��", "TYPE1");
			map.put("�����һ�������", "TYPE2");
			map.put("�����һ�Ԥ����", "TYPE3");
			map.put("ֱ������˫��", "TYPE4");
			map.put("ֱ������������", "TYPE5");
			map.put("ֱ������Ԥ����", "TYPE6");
			map.put("ֱ������Ԥ����", "TYPE6");
			map.put("�ͻ������������ּ�����ʽ", "TYPE7");
			map.put("�ͻ�������ʽ", "TYPE8");
		}

		if ("8".equals(imptype)) { // �ͻ�-��������6-�ͻ����ս�
			map.put("��ͬ���", "HTNO");
			map.put("��ͬ����", "HTTYPE");
			map.put("�ͻ�����", "CUSTCODE");
			map.put("�ͻ�����", "CUSTNAME");
			map.put("�ƻ���Ч����", "STARTDATE");
			map.put("�ƻ���ֹ����", "ENDDATE");
			map.put("��Ʒ", "INVBADOCGROUP");
			map.put("������׼", "SCHEMEGROUP");
			map.put("0�ͻ�������ʽ", "TYPE0");
			map.put("7�ͻ�������������", "TYPE7");
			map.put("6�ͻ�����������̥", "TYPE6");
			map.put("5ֱ��ͻ�������˫���Ʒ", "TYPE5");
			map.put("1�����һ�˫��", "TYPE1");
			map.put("2�����һ�������", "TYPE2");
			map.put("3ֱ������˫��", "TYPE3");
			map.put("4ֱ������������", "TYPE4");
			map.put("�����һ�Ԥ����", "TYPE8");
			map.put("ֱ����������̥", "TYPE9");
			map.put("ֱ������Ԥ����", "TYPE10");
			map.put("��������", "SALETYPE");
		}
		if ("9".equals(imptype)) { // ����ͻ�-������ͻ��б�
			map.put("�ͻ�����", "CUSTCODE");
			map.put("�ͻ�����", "CUSTNAME");
			map.put("��ʼ�·�", "STARTMONTH");
			map.put("�����·�", "ENDMONTH");
			map.put("�������", "RATE");
			map.put("��������", "TYPE");
		}
		if ("10".equals(imptype) || "11".equals(imptype)) {// ��ֵ�ͻ�Ŀ��Ϳͻ�����Ŀ��
			map.put("�ͻ�����", "CUSTCODE");
			map.put("�ͻ�����", "CUSTNAME");
			map.put("�����ڼ�", "MONTHYEAR");
			map.put("Ԥ��Ŀ�꣨�֣�", "TARGET");
			map.put("ս��", "ZHANGQ");
			map.put("Ӫ����", "YXB");
			map.put("�ͻ�����", "CUSTTYPE");
			map.put("�����ѣ�Ԫ��", "MONEYOFPROMTION");
			map.put("��������Ԫ��", "MONEYOFBANK");
			map.put("��Ʒ����", "CPXX");
			map.put("Ƭ��", "SALESTR");
		}
		if ("12".equals(imptype)) { // �ͻ����ܷ�����
			map.put("�������", "SECHNO");
			map.put("��������", "SECHNAME");
			map.put("��������", "JSDATE");
			map.put("ս��", "ZHANGQ");
			map.put("Ӫ����", "YXB");
			map.put("Ƭ��", "SALESTR");
			map.put("��֯", "UNIT");
			map.put("������׼", "STANDARD");
			map.put("��������", "BASENUMBER");
			map.put("Ŀ������", "HIGHERNUMBER");
			map.put("���ڻ������񵥼�", "LOWERNUMBER");
			map.put("��������̶�����", "BASEGDPRICE");
			map.put("�������񵥼�", "BASEPRICE");
			map.put("��Ŀ�굥��ϵ��", "XS");
			map.put("��Ŀ�����񵥼�", "HIGHERPRICE");
			map.put("��Ŀ������ⶥ����", "HIGHERMAXPRICE");
		}

		if ("13".equals(imptype)) {// Ƭ�齱���������
			map.put("�������", "SECHNO");
			map.put("��������", "SECHNAME");
			map.put("��������", "JSMONTH");
			map.put("ս��", "ZHANGQ");
			map.put("Ӫ����", "YXB");
			map.put("Ƭ��", "SALESTR");
			map.put("�ͻ����ܱ���", "PSNCODE");
			map.put("�ͻ���������", "PSNNAME");
			map.put("�ڸ�ϵ��", "STATUS");
			map.put("�������", "XS");
			map.put("Ԥ������", "BL");
		}
		if ("14".equals(imptype)) {// ����ֽ��
			map.put("�������", "SCHEMNO");
			map.put("��������", "SCHEMNAME");
			map.put("ս��", "ZHANGQ");
			map.put("Ӫ����", "YXB");
			map.put("�ͻ�����", "CUSTCODE");
			map.put("�ͻ�����", "CUSTNAME");
			map.put("����", "LEVEL_CUST");
			map.put("������������", "CUSTTYPE");
			map.put("��������ȡ���ھ�", "INVBASDOCID");
			map.put("�������������Ʒȡ���ھ�", "INVBASDOCID_OTHER");
			map.put("����/���", "SALETYPE");
			map.put("���㿼��ȡ���ھ�", "LEVEL_WD");
			map.put("�����·�1", "TARGET_MONTH1");
			map.put("Ƭ��1", "TARGET_SALESTR1");
			map.put("�ͻ����ܱ���1", "TARGET_PSNCODE1");
			map.put("�ͻ���������1", "TARGET_PSNNAME1");
			map.put("�����·�2", "TARGET_MONTH2");
			map.put("Ƭ��2", "TARGET_SALESTR2");
			map.put("�ͻ����ܱ���2", "TARGET_PSNCODE2");
			map.put("�ͻ���������2", "TARGET_PSNNAME2");
			map.put("�����·�3", "TARGET_MONTH3");
			map.put("Ƭ��3", "TARGET_SALESTR3");
			map.put("�ͻ����ܱ���3", "TARGET_PSNCODE3");
			map.put("�ͻ���������3", "TARGET_PSNNAME3");
		}
		if ("15".equals(imptype)) {// �˲�ѡ��
			map.put("�·�", "MONTH");
			map.put("����", "NAME");
			map.put("�Ա�", "SEX");
			map.put("����", "AGE");
			map.put("ѧ��", "SCHOOLLEVEL");
			map.put("����", "ADDRESS");
			map.put("��ҵԺУ", "SCHOOL");
			map.put("רҵ", "MAJOR");
			map.put("��ϵ�绰", "PHONE");
			map.put("���Խ׶�_���Թ�", "INTERVIEWER");
			map.put("���Խ׶�_���Է���", "INTERVIEWERSCORE");
			map.put("���Խ׶�_���Խ��", "INTERVIEWERRESULT");
			map.put("���Խ׶�_��������", "EVALUATION");
			map.put("��¼�ý׶�_����", "COMEPOSTION");
			map.put("��¼�ý׶�_�����·�", "COMEPOSTIONMONTH");
			map.put("��¼�ý׶�_����ʱ��", "COMEPOSTIONTIME");
			map.put("��¼�ý׶�_δ����ԭ��", "REASON");
			map.put("��Ħ�׶�_��ʼ����", "WATCHSTARTDATE");
			map.put("��Ħ�׶�_��������", "WATCHENDDATE");
			map.put("��Ħ�׶�_��Ħ����", "WATCHDEPT");
			map.put("��Ħ�׶�_��Ħʦ��", "WATCHMASTERWORKER");
			map.put("��Ħ�׶�_��Ħ����", "WATCHSCORE");
			map.put("��Ħ�׶�_��Ħ���", "WATCHRESULT");
			map.put("��Ħ�׶�_ʦ������", "WATCHMASTEREVALUATION");
			map.put("��˾��ѵ�׶�_��ʼ����", "TRAINSTARTDATE");
			map.put("��˾��ѵ�׶�_��������", "TRAINENDDATE");
			map.put("��˾��ѵ�׶�_��ѵ����", "TRAINSCORE");
			map.put("��˾��ѵ�׶�_��ѵ���", "TRAINRESULT");
			map.put("��˾��ѵ�׶�_��ѵ����", "TRAINEVALUATION");
			map.put("��˾��ѵ�׶�_��ְ����", "TRAINDEPT");
			map.put("��˾��ѵ�׶�_��ְ��λ", "TRAINPOST");
		}
		if ("16".equals(imptype)) {// �˲�����
			map.put("����/ս��", "ZHANGQ");
			map.put("����/��˾", "YXB");
			map.put("����ҵ������", "TYPE");
			map.put("ͽ����Ϣ_��Ա����", "PSNCODE_TD");
			map.put("ͽ����Ϣ_����", "PSNNAME_TD");
			map.put("ͽ����Ϣ_��λ", "POSTNAME_TD");
			map.put("ʦ����Ϣ_��Ա����", "PSNCODE_SF");
			map.put("ʦ����Ϣ_����", "PSNNAME_SF");
			map.put("ʦ����Ϣ_��λ", "POSTNAME_SF");
			map.put("ʦ����Ϣ_��Ա����1", "PSNCODE_SF1");
			map.put("ʦ����Ϣ_����1", "PSNNAME_SF1");
			map.put("ʦ����Ϣ_��λ1", "POSTNAME_SF1");
			map.put("���̿�ʼʱ��", "BEGIN_TRAIN");
			map.put("���̽���ʱ��", "END_TRAIN");
			map.put("��������", "TRAINGDAYS");
			map.put("����ʱ��", "ACCEPTANCETIME");
			map.put("�ɼ�_ѧϰ��ͼ", "SCORE1");
			map.put("�ɼ�_���ճɼ�", "SCORE2");
			map.put("��������_�ۺϳɼ�", "SCORE3");
			map.put("��������_���ս��", "TRAIN_CONCLUSION");
			map.put("�ɲ����۳ɼ�", "TRAIN_ACHIEVEMENT");
			map.put("�ϸ�֤���", "TRAIN_POSTCARD");
			map.put("����(ʦ��/���ų�)_�Ը�", "CHARACTER");
			map.put("����(ʦ��/���ų�)_������", "CONSCIENTIOUSNESS");
			map.put("����(ʦ��/���ų�)_��ֵ��", "SENSEOFWORTH");
			map.put("����(ʦ��/���ų�)_�ۺ�����", "COMPREHENSIVE");
			map.put("����(ʦ��/���ų�)_����", "CONCLUSION");
			map.put("��Ħ�׶�_���ڲ���", "TOPJOBDEPARTMENT");
			map.put("��Ħ�׶�_���ڸ�λ", "POSTPOSITION");
			map.put("��Ħ�׶�_����", "CONCLUSIONPOST");
			map.put("��˾��ѵ�׶�_�÷�", "REGULAR_SCORE");
			map.put("��˾��ѵ�׶�_���", "REGULAR_CONCLUSION");
		}
		if ("17".equals(imptype)) {// �ɹ�ͷ��
			map.put("����", "BILLDATE");
			map.put("����", "ZHANGQ");
			map.put("ʡ��", "PROVICE");
			map.put("��˾", "UNITNAME");
			map.put("ԭ������", "INVNAME");
			map.put("�����", "INHAND");
			map.put("����", "ORDERNO");
			map.put("�ϼ�", "SUMNO");
			map.put("����", "INHANDPRICE");
			map.put("������", "ORDERPRICE");
			map.put("��Ȩƽ����", "AVGPRICE");
			map.put("�����", "HQPRICE");
			map.put("��ӯ(��)", "YK");
			map.put("ӯ�����/��Ԫ", "YKMONEY");
			map.put("30��ƽ������(��)", "NNUMBER");
			map.put("ͷ������(��)", "DAYS");

		}

		if ("18".equals(imptype)) {// ����ģ��-�ͻ������ֹ���
			map.put("��������", "TYPE");
			map.put("����", "UNITNAME");
			map.put("�ͻ�����", "CUSTCODE");
			map.put("�ͻ�����", "CUSTNAME");
			map.put("��Ʒ���", "INVBASDOCGROUP");
			map.put("��Ʒ���1", "INVBASDOCGROUP1");
			map.put("��Ʒ���2", "INVBASDOCGROUP2");
			map.put("���������·�", "JTMONTH");
			map.put("�����������", "JTMONEY");
			map.put("��������", "ZQ");
			map.put("��ע", "MEMO");

		}
		if ("19".equals(imptype)) {// ����_������׼
			// map.put("��׼ʱ��","STANDID");
			map.put("������׼����", "STANDRADNAME");
			// map.put("����", "TYPE");
			map.put("��ʼ����", "STARTNUM");
			map.put("��������", "ENDNUM");
			map.put("����", "PRICE");
			// map.put("�Ƿ���", "FLAG");
			// map.put("TS", "TS");
		}
		if("20".equals(imptype)) {//Ԥ��Ŀ�꣨ս����
			map.put("ս��", "ZHANQ");
			map.put("�·�", "MONTH");
			map.put("����Ŀ���������֣�", "MBNNUMBER");
		}
		if("21".equals(imptype)) {//Ԥ��Ŀ�꣨������
			map.put("����", "CPX");
			map.put("�·�", "MONTH");
			map.put("����Ŀ���������֣�", "MBNNUMBER");
		}
		if("22".equals(imptype)) {//Ԥ��Ŀ�꣨Ӫ������
			map.put("ս��", "ZHANQ");
			map.put("Ӫ����", "YXB");
			map.put("�·�", "MONTH");
			map.put("����Ŀ���������֣�", "MBNNUMBER");
		}
		read(in, fileType, map, request, response);

	}

	/**
	 * ��ȡ�ļ�
	 * 
	 * @param in
	 * @param fileType
	 * @throws Exception
	 */
	public static void read(InputStream in, String fileType,
			Map<String, String> map, HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		int dqrow = 0;
		try {
			List<JSONObject> list = new ArrayList<JSONObject>();
			//
			if (fileType.equals("xls")) {
				HSSFWorkbook wb = new HSSFWorkbook(in);
				HSSFSheet sheet = wb.getSheetAt(0);
				HSSFRow headRow = sheet.getRow(0);
				HSSFRow row = null;
				HSSFCell cell = null;
				Object cellVal = null;
				SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
				// ��һ������
				if (sheet != null && sheet.getLastRowNum() > 0) {
					for (int i = 1; i <= sheet.getLastRowNum(); i++) {
						dqrow = i;
						JSONObject jsonObject = new JSONObject();
						// �õ���ǰ�����������
						row = sheet.getRow(i);

						if (null != row) {
							for (int j = 0; j < row.getLastCellNum(); j++) {
								if (j > headRow.getLastCellNum() - 1) {
									break;
								}
								cell = row.getCell(j);
								if (cell == null) {
									jsonObject.put(map.get(headRow.getCell(j)
											.getStringCellValue().toString().trim()
											.replaceAll("\r|\n", "")), "");
									continue;
								}
								if (headRow.getCell(j) != null
										&& cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC
										&& HSSFDateUtil
												.isCellDateFormatted(cell)) {
									cellVal = format.format(cell
											.getDateCellValue());
								} else {
									cellVal = ExcelObject.getCellValue(cell);
									if (cellVal == null
											&& "2".equals(request.getAttribute(
													"imptype").toString())
											&& j == 8) {
										cellVal = "";
									}
									if (cellVal == null
											&& "3".equals(request.getAttribute(
													"imptype").toString())
											&& j == 8) {
										cellVal = "";
									}
									if (cellVal == null
											&& "4".equals(request.getAttribute(
													"imptype").toString())
											&& j == 8) {
										cellVal = "";
									}
									if ("5".equals(request.getAttribute(
											"imptype").toString())) {
										if (cellVal == null) {
											cellVal = "";
										}
									}
								}
								if (cellVal == null) {
									cellVal = "";
								}
								/*
								 * else {
								 * 
								 * cellVal = ExcelObject.getCellValue(cell);
								 * 
								 * if (cellVal == null&&
								 * "2".equals(request.getAttribute
								 * ("imptype").toString())&& ((j >= 12 && j <=
								 * 17) || j >= 19)) cellVal = ""; if (cellVal ==
								 * null&&
								 * "15".equals(request.getAttribute("imptype"
								 * ).toString())&&j==9) cellVal = ""; if
								 * (cellVal == null&&
								 * "16".equals(request.getAttribute
								 * ("imptype").toString())&&j==6) cellVal = "";
								 * if (cellVal == null&&
								 * "2".equals(request.getAttribute
								 * ("imptype").toString())) throw new
								 * Exception("��" + i + "��" + j + "��Ϊ�գ�����д"); }
								 */

								/*
								 * System.out.print(cellVal);
								 * System.out.println(i + "��");
								 * System.out.println(map.get(headRow.getCell(j)
								 * .getStringCellValue().toString().trim()
								 * .replaceAll("\r|\n", "")));
								 */

								jsonObject.put(map.get(headRow.getCell(j)
										.getStringCellValue().toString().trim()
										.replaceAll("\r|\n", "")), cellVal);
							}
							list.add(jsonObject);

						}
					}
				}

			} else {
				throw new Exception("ֻ֧��2003�汾��Excel���룡");
			}
			StringBuffer sBuffer = new StringBuffer();
			String imptype = request.getAttribute("imptype").toString();

			if ("1".equals(imptype)) {
				imp_sbt_employee_scheme_kf(list, request, response);
			}
			if ("2".equals(imptype) || "3".equals(imptype)
					|| "4".equals(imptype)) {
				imp_sbt_employee_scheme_kf_check(list, request, response,
						imptype);
			}
			if ("5".equals(imptype)) {
				imp_SBT_CUST_SCHEME(list, request, response);
			}

			if ("6".equals(imptype)) {
				imp_SBT_CUST_HT(list, request, response);
			}
			if ("7".equals(imptype)) {
				imp_sbt_cust_ht_sign(list, request, response);
			}
			if ("8".equals(imptype)) {
				imp_sbt_cust_ht_year(list, request, response);
			}
			if ("9".equals(imptype)) {
				imp_sbt_cust_specialcust(list, request, response);
			}
			if ("10".equals(imptype) || "11".equals(imptype)) {
				imp_saletarget_customer(list, request, response, imptype);
			}
			if ("12".equals(imptype)) {
				imp_SBT_CUST_EMPLOYEE_SCHEME(list, request, response);
			}
			if ("13".equals(imptype)) {
				imp_sbt_cust_employee_standrad(list, request, response);
			}
			if ("14".equals(imptype)) {
				imp_SBT_CUST_EMPLOYEE_TARGET(list, request, response);
			}
			if ("15".equals(imptype)) {
				imp_sbt_hr_personchose(list, request, response);
			}
			if ("16".equals(imptype)) {
				imp_sbt_hr_persontraining(list, request, response);
			}
			if ("17".equals(imptype)) {
				imp_sbt_cg_markettc(list, request, response);
			}
			if ("18".equals(imptype)) {
				imp_sjlr_cust_sum(list, request, response);
			}
			if ("19".equals(imptype)) {
				imp_sbt_cust_standrad(list, request, response);
			}
			if("20".equals(imptype)) {//Ԥ��Ŀ�꣨ս����
				imp_app_main_yjmb_zq(list, request, response);
			}
			if("21".equals(imptype)) {//Ԥ��Ŀ�꣨������
				imp_app_main_yjmb_xx(list, request, response);
			}
			if("22".equals(imptype)) {//Ԥ��Ŀ�꣨Ӫ������
				imp_app_main_yjmb_yxb(list, request, response);
			}
			in.close();
			if (sBuffer.length() > 0) {
				request.setAttribute("msg", "������Աδ����(��������ɹ�)��"
						+ sBuffer.toString());
			}

		} catch (Exception e) {
			dqrow = dqrow + 1;
			throw e;

		}
	}

	/**
	 *�¶�����-�ͻ������ֹ���
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sjlr_cust_sum(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps2 = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// ��ɾ���¶ȵ�����
			sql = "DELETE wb_erp.sjlr_cust_sum where  jtmonth=? and type=? ";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("JTMONTH"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("TYPE"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		int i = 0;
		for (int f = 0; f < voList.size(); f++) {
			i++;
			System.out.println(i);
			JSONObject vo = voList.get(f);

			// У�������
			sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVBASDOCGROUP"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("INVBADOCGROUP").toString()
						+ "�����ϲ����ڣ���������µ��룡");
			}
			DbUtil.closeStatement(ps1);
			if (!vo.get("INVBASDOCGROUP1").toString().equals("")) {
				// У�������1
				sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo
						.get("INVBASDOCGROUP1"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + 1 + "��"
							+ vo.get("INVBADOCGROUP").toString()
							+ "������1�����ڣ���������µ��룡");
				}
				DbUtil.closeStatement(ps1);
			}
			// У�������2
			if (!vo.get("INVBASDOCGROUP2").toString().equals("")) {
				sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo
						.get("INVBASDOCGROUP2"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + 1 + "��"
							+ vo.get("INVBADOCGROUP").toString()
							+ "������2�����ڣ���������µ��룡");
				}
				DbUtil.closeStatement(ps1);
			}

			sql = "insert into wb_erp.sjlr_cust_sum (TYPE, UNITNAME, CUSTCODE, CUSTNAME, INVBASDOCGROUP, JTMONTH, JTMONEY, ZQ, MEMO, ID,INVBASDOCGROUP1,INVBASDOCGROUP2)"
					+ " values (?,?,?,?, (select id  from wb_erp.sbt_cust_invbasdocgroup a  where nvbasdocgroup=? and rownum=1) , ?,?, ?, ?, ?,(select id  from wb_erp.sbt_cust_invbasdocgroup a  where nvbasdocgroup=? and rownum=1),(select id  from wb_erp.sbt_cust_invbasdocgroup a  where nvbasdocgroup=? and rownum=1))";

			ps2 = conn.prepareStatement(sql);
			DbUtil.setObject(ps2, 1, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps2, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps2, 3, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps2, 4, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps2, 5, Types.VARCHAR, vo.opt("INVBASDOCGROUP"));
			DbUtil.setObject(ps2, 6, Types.VARCHAR, vo.opt("JTMONTH"));
			DbUtil.setObject(ps2, 7, Types.VARCHAR, vo.opt("JTMONEY"));
			DbUtil.setObject(ps2, 8, Types.VARCHAR, vo.opt("ZQ"));
			DbUtil.setObject(ps2, 9, Types.VARCHAR, vo.opt("MEMO"));
			String PK_ID = SysUtil.getId();
			ps2.setString(10, PK_ID);
			DbUtil.setObject(ps2, 11, Types.VARCHAR, vo.opt("INVBASDOCGROUP1"));
			DbUtil.setObject(ps2, 12, Types.VARCHAR, vo.opt("INVBASDOCGROUP2"));
			ps2.execute();
			ps2.close();
			// �ر���Դ
			DbUtil.closeStatement(ps2);
		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}

	/**
	 *�ɹ�ͷ��
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_cg_markettc(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps2 = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// ��ɾ���¶ȵ�����
			sql = "DELETE wb_erp.sbt_cg_markettc where billdate=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("BILLDATE"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			sql = "insert into wb_erp.sbt_cg_markettc(billdate  ,zhangq  ,provice  ,unitname  ,invname  ,inhand  ,orderno  ,sumno  ,inhandprice  ,orderprice  ,avgprice  ,hqprice  ,yk  ,ykmoney	,nnumber	,days	) "
					+ " values(?  ,?  ,?  ,?  ,?  ,?  ,?  ,?  ,?  ,?  ,?  ,?  ,?  ,?	,?	,?	)";

			ps2 = conn.prepareStatement(sql);
			DbUtil.setObject(ps2, 1, Types.VARCHAR, vo.opt("BILLDATE"));
			DbUtil.setObject(ps2, 2, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps2, 3, Types.VARCHAR, vo.opt("PROVICE"));
			DbUtil.setObject(ps2, 4, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps2, 5, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps2, 6, Types.VARCHAR, vo.opt("INHAND"));
			DbUtil.setObject(ps2, 7, Types.VARCHAR, vo.opt("ORDERNO"));
			DbUtil.setObject(ps2, 8, Types.VARCHAR, vo.opt("SUMNO"));
			DbUtil.setObject(ps2, 9, Types.VARCHAR, vo.opt("INHANDPRICE"));
			DbUtil.setObject(ps2, 10, Types.VARCHAR, vo.opt("ORDERPRICE"));
			DbUtil.setObject(ps2, 11, Types.VARCHAR, vo.opt("AVGPRICE"));
			DbUtil.setObject(ps2, 12, Types.VARCHAR, vo.opt("HQPRICE"));
			DbUtil.setObject(ps2, 13, Types.VARCHAR, vo.opt("YK"));
			DbUtil.setObject(ps2, 14, Types.VARCHAR, vo.opt("YKMONEY"));
			DbUtil.setObject(ps2, 15, Types.VARCHAR, vo.opt("NNUMBER"));
			DbUtil.setObject(ps2, 16, Types.VARCHAR, vo.opt("DAYS"));
			ps2.execute();
			ps2.close();
			// �ر���Դ

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}

	/**
	 * �˲�����
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_hr_persontraining(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps2 = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// ��ɾ���¶ȵ�����
			sql = "DELETE wb_erp.sbt_hr_persontraining where PSNCODE_TD=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("PSNCODE_TD"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			sql = "insert into wb_erp.sbt_hr_persontraining(ID, ZHANGQ,YXB,TYPE,	 PSNCODE_TD,	 PSNNAME_TD,	 POSTNAME_TD,	 PSNCODE_SF,	 PSNNAME_SF,	 POSTNAME_SF,	 BEGIN_TRAIN,	 END_TRAIN,	 TRAINGDAYS,	 ACCEPTANCETIME,	 SCORE1,	 SCORE2,	 SCORE3,	 TRAIN_CONCLUSION,	 CHARACTER,	 CONSCIENTIOUSNESS,	 SENSEOFWORTH,	 COMPREHENSIVE,	 CONCLUSION,	 TOPJOBDEPARTMENT,	 POSTPOSITION,	 CONCLUSIONPOST,	 REGULAR_SCORE,	 REGULAR_CONCLUSION, TS, COPERATOR,TRAIN_ACHIEVEMENT,TRAIN_POSTCARD,PSNCODE_SF1,PSNNAME_SF1,POSTNAME_SF1) "
					+ " values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?,?,?,?)";

			ps2 = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps2.setString(1, PK_ID);
			DbUtil.setObject(ps2, 2, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps2, 3, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps2, 4, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps2, 5, Types.VARCHAR, vo.opt("PSNCODE_TD"));
			DbUtil.setObject(ps2, 6, Types.VARCHAR, vo.opt("PSNNAME_TD"));
			DbUtil.setObject(ps2, 7, Types.VARCHAR, vo.opt("POSTNAME_TD"));
			DbUtil.setObject(ps2, 8, Types.VARCHAR, vo.opt("PSNCODE_SF"));
			DbUtil.setObject(ps2, 9, Types.VARCHAR, vo.opt("PSNNAME_SF"));
			DbUtil.setObject(ps2, 10, Types.VARCHAR, vo.opt("POSTNAME_SF"));
			DbUtil.setObject(ps2, 11, Types.VARCHAR, vo.opt("BEGIN_TRAIN"));
			DbUtil.setObject(ps2, 12, Types.VARCHAR, vo.opt("END_TRAIN"));
			DbUtil.setObject(ps2, 13, Types.VARCHAR, vo.opt("TRAINGDAYS"));
			DbUtil.setObject(ps2, 14, Types.VARCHAR, vo.opt("ACCEPTANCETIME"));
			DbUtil.setObject(ps2, 15, Types.VARCHAR, vo.opt("SCORE1"));
			DbUtil.setObject(ps2, 16, Types.VARCHAR, vo.opt("SCORE2"));
			DbUtil.setObject(ps2, 17, Types.VARCHAR, vo.opt("SCORE3"));
			DbUtil
					.setObject(ps2, 18, Types.VARCHAR, vo
							.opt("TRAIN_CONCLUSION"));
			DbUtil.setObject(ps2, 19, Types.VARCHAR, vo.opt("CHARACTER"));
			DbUtil.setObject(ps2, 20, Types.VARCHAR, vo
					.opt("CONSCIENTIOUSNESS"));
			DbUtil.setObject(ps2, 21, Types.VARCHAR, vo.opt("SENSEOFWORTH"));
			DbUtil.setObject(ps2, 22, Types.VARCHAR, vo.opt("COMPREHENSIVE"));
			DbUtil.setObject(ps2, 23, Types.VARCHAR, vo.opt("CONCLUSION"));
			DbUtil
					.setObject(ps2, 24, Types.VARCHAR, vo
							.opt("TOPJOBDEPARTMENT"));
			DbUtil.setObject(ps2, 25, Types.VARCHAR, vo.opt("POSTPOSITION"));
			DbUtil.setObject(ps2, 26, Types.VARCHAR, vo.opt("CONCLUSIONPOST"));
			DbUtil.setObject(ps2, 27, Types.VARCHAR, vo.opt("REGULAR_SCORE"));
			DbUtil.setObject(ps2, 28, Types.VARCHAR, vo
					.opt("REGULAR_CONCLUSION"));
			DbUtil.setObject(ps2, 29, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps2, 30, Types.VARCHAR, vo
					.opt("TRAIN_ACHIEVEMENT"));
			DbUtil.setObject(ps2, 31, Types.VARCHAR, vo.opt("TRAIN_POSTCARD"));
			DbUtil.setObject(ps2, 32, Types.VARCHAR, vo.opt("PSNCODE_SF1"));
			DbUtil.setObject(ps2, 33, Types.VARCHAR, vo.opt("PSNNAME_SF1"));
			DbUtil.setObject(ps2, 34, Types.VARCHAR, vo.opt("POSTNAME_SF1"));
			ps2.execute();
			ps2.close();
			// �ر���Դ

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}

	/**
	 * �˲�ѡ��
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_hr_personchose(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps2 = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			sql = "insert into wb_erp.sbt_hr_personchose (ID, MONTH, NAME, SEX, AGE, SCHOOLLEVEL, ADDRESS, SCHOOL, MAJOR, PHONE, INTERVIEWER, INTERVIEWERSCORE, INTERVIEWERRESULT, EVALUATION, COMEPOSTION, COMEPOSTIONMONTH, COMEPOSTIONTIME, REASON, WATCHSTARTDATE, WATCHENDDATE, WATCHDEPT, WATCHMASTERWORKER, WATCHSCORE, WATCHRESULT, WATCHMASTEREVALUATION, TRAINSTARTDATE, TRAINENDDATE, TRAINSCORE, TRAINRESULT, TRAINEVALUATION, TRAINDEPT, TRAINPOST, TS, COPERATOR) "
					+ " values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?)";

			ps2 = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			// ��ֵ���ӱ�������
			String FK_ID = PK_ID;
			ps2.setString(1, PK_ID);
			DbUtil.setObject(ps2, 2, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps2, 3, Types.VARCHAR, vo.opt("NAME"));
			DbUtil.setObject(ps2, 4, Types.VARCHAR, vo.opt("SEX"));
			DbUtil.setObject(ps2, 5, Types.VARCHAR, vo.opt("AGE"));
			DbUtil.setObject(ps2, 6, Types.VARCHAR, vo.opt("SCHOOLLEVEL"));
			DbUtil.setObject(ps2, 7, Types.VARCHAR, vo.opt("ADDRESS"));
			DbUtil.setObject(ps2, 8, Types.VARCHAR, vo.opt("SCHOOL"));
			DbUtil.setObject(ps2, 9, Types.VARCHAR, vo.opt("MAJOR"));
			DbUtil.setObject(ps2, 10, Types.VARCHAR, vo.opt("PHONE"));
			DbUtil.setObject(ps2, 11, Types.VARCHAR, vo.opt("INTERVIEWER"));
			DbUtil
					.setObject(ps2, 12, Types.VARCHAR, vo
							.opt("INTERVIEWERSCORE"));
			DbUtil.setObject(ps2, 13, Types.VARCHAR, vo
					.opt("INTERVIEWERRESULT"));
			DbUtil.setObject(ps2, 14, Types.VARCHAR, vo.opt("EVALUATION"));
			DbUtil.setObject(ps2, 15, Types.VARCHAR, vo.opt("COMEPOSTION"));
			DbUtil
					.setObject(ps2, 16, Types.VARCHAR, vo
							.opt("COMEPOSTIONMONTH"));
			DbUtil.setObject(ps2, 17, Types.VARCHAR, vo.opt("COMEPOSTIONTIME"));
			DbUtil.setObject(ps2, 18, Types.VARCHAR, vo.opt("REASON"));
			DbUtil.setObject(ps2, 19, Types.VARCHAR, vo.opt("WATCHSTARTDATE"));
			DbUtil.setObject(ps2, 20, Types.VARCHAR, vo.opt("WATCHENDDATE"));
			DbUtil.setObject(ps2, 21, Types.VARCHAR, vo.opt("WATCHDEPT"));
			DbUtil.setObject(ps2, 22, Types.VARCHAR, vo
					.opt("WATCHMASTERWORKER"));
			DbUtil.setObject(ps2, 23, Types.VARCHAR, vo.opt("WATCHSCORE"));
			DbUtil.setObject(ps2, 24, Types.VARCHAR, vo.opt("WATCHRESULT"));
			DbUtil.setObject(ps2, 25, Types.VARCHAR, vo
					.opt("WATCHMASTEREVALUATION"));
			DbUtil.setObject(ps2, 26, Types.VARCHAR, vo.opt("TRAINSTARTDATE"));
			DbUtil.setObject(ps2, 27, Types.VARCHAR, vo.opt("TRAINENDDATE"));
			DbUtil.setObject(ps2, 28, Types.VARCHAR, vo.opt("TRAINSCORE"));
			DbUtil.setObject(ps2, 29, Types.VARCHAR, vo.opt("TRAINRESULT"));
			DbUtil.setObject(ps2, 30, Types.VARCHAR, vo.opt("TRAINEVALUATION"));
			DbUtil.setObject(ps2, 31, Types.VARCHAR, vo.opt("TRAINDEPT"));
			DbUtil.setObject(ps2, 32, Types.VARCHAR, vo.opt("TRAINPOST"));
			DbUtil.setObject(ps2, 33, Types.VARCHAR, request
					.getAttribute("sys.userName"));

			ps2.execute();
			ps2.close();
			// �ر���Դ

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}

	/**
	 * ����ֽ��
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_SBT_CUST_EMPLOYEE_TARGET(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;

		// ��ѭ��ɾ��������ͬ�ķ�����

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// ��ɾ���¶ȵ�����
			sql = "DELETE wb_erp.SBT_CUST_EMPLOYEE_TARGET_month where schemno=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SCHEMNO"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);

			// ��ɾ����������
			sql = "DELETE wb_erp.SBT_CUST_EMPLOYEE_TARGET where schemno=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SCHEMNO"));
			ps1.executeUpdate();
			conn.commit();
			DbUtil.closeStatement(ps1);
			ps1.close();

		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PreparedStatement ps2 = null;
			// У������ȡ���ھ�
			sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
			ps2 = conn.prepareStatement(sql);
			DbUtil.setObject(ps2, 1, Types.VARCHAR, vo.get("INVBASDOCID"));
			rSet = ps2.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("INVBASDOCID").toString()
						+ "��������ȡ���ھ������ڣ���������µ��룡");
			}
			ps2.close();
			rSet.close();

			// У������ȡ���ھ�
			if (!vo.get("INVBASDOCID_OTHER").toString().equals("")) {
				System.out.print(vo.get("INVBASDOCID_OTHER"));
				sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
				ps2 = conn.prepareStatement(sql);
				DbUtil.setObject(ps2, 1, Types.VARCHAR, vo
						.get("INVBASDOCID_OTHER"));
				rSet = ps2.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + 1 + "��"
							+ vo.get("INVBASDOCID_OTHER").toString()
							+ "��������ȡ���ھ������ڣ���������µ��룡");
				}
			}
			ps2.close();
			rSet.close();

			sql = "insert into wb_erp.SBT_CUST_EMPLOYEE_TARGET (ID,SCHEMNO,	SCHEMNAME,	ZHANGQ,	YXB,	CUSTCODE,	CUSTNAME,	LEVEL_CUST,	CUSTTYPE,	INVBASDOCID,	INVBASDOCID_OTHER,	SALETYPE,	LEVEL_WD) "
					+ " values (?,?,?,	?,	?,	?,	?,	?,	?,	(select id from wb_erp.sbt_cust_invbasdocgroup a where  a.nvbasdocgroup=?),	(select id from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?),	?,?)";

			ps2 = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			// ��ֵ���ӱ�������
			String FK_ID = PK_ID;
			ps2.setString(1, PK_ID);
			DbUtil.setObject(ps2, 2, Types.VARCHAR, vo.opt("SCHEMNO"));
			DbUtil.setObject(ps2, 3, Types.VARCHAR, vo.opt("SCHEMNAME"));
			DbUtil.setObject(ps2, 4, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps2, 5, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps2, 6, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps2, 7, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps2, 8, Types.VARCHAR, vo.opt("LEVEL_CUST"));
			DbUtil.setObject(ps2, 9, Types.VARCHAR, vo.opt("CUSTTYPE"));
			DbUtil.setObject(ps2, 10, Types.VARCHAR, vo.opt("INVBASDOCID"));
			DbUtil.setObject(ps2, 11, Types.VARCHAR, vo
					.opt("INVBASDOCID_OTHER"));
			DbUtil.setObject(ps2, 12, Types.VARCHAR, vo.opt("SALETYPE"));
			DbUtil.setObject(ps2, 13, Types.VARCHAR, vo.opt("LEVEL_WD"));
			ps2.execute();
			ps2.close();
			// �ر���Դ

			// ѭ���������ɼ��ȵ�һ�¡����¡����µ�����ֽ��
			sql = "insert into wb_erp.SBT_CUST_EMPLOYEE_TARGET_month (ID ,SCHEMNO, SCHEMNAME, ZHANGQ        ,YXB          ,CUSTCODE      , CUSTNAME    ,  FK_ID, TARGET_MONTH  , TARGET_SALESTR , TARGET_PSNCODE , TARGET_PSNNAME) "
					+ " values (? ,?, ?, ? ,?  ,? , ?  ,  ?, ?  , ? , ? , ? )";
			// ��һ����
			ps2 = conn.prepareStatement(sql);
			// ������������
			PK_ID = SysUtil.getId();
			ps2.setString(1, PK_ID);
			DbUtil.setObject(ps2, 2, Types.VARCHAR, vo.opt("SCHEMNO"));
			DbUtil.setObject(ps2, 3, Types.VARCHAR, vo.opt("SCHEMNAME"));
			DbUtil.setObject(ps2, 4, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps2, 5, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps2, 6, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps2, 7, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps2, 8, Types.VARCHAR, FK_ID);
			DbUtil.setObject(ps2, 9, Types.VARCHAR, vo.opt("TARGET_MONTH1"));
			DbUtil.setObject(ps2, 10, Types.VARCHAR, vo.opt("TARGET_SALESTR1"));
			DbUtil.setObject(ps2, 11, Types.VARCHAR, vo.opt("TARGET_PSNCODE1"));
			DbUtil.setObject(ps2, 12, Types.VARCHAR, vo.opt("TARGET_PSNNAME1"));
			ps2.execute();
			ps2.close();
			// �ر���Դ

			// �ڶ�����
			ps2 = conn.prepareStatement(sql);
			// ������������
			PK_ID = SysUtil.getId();
			ps2.setString(1, PK_ID);
			DbUtil.setObject(ps2, 2, Types.VARCHAR, vo.opt("SCHEMNO"));
			DbUtil.setObject(ps2, 3, Types.VARCHAR, vo.opt("SCHEMNAME"));
			DbUtil.setObject(ps2, 4, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps2, 5, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps2, 6, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps2, 7, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps2, 8, Types.VARCHAR, FK_ID);
			DbUtil.setObject(ps2, 9, Types.VARCHAR, vo.opt("TARGET_MONTH2"));
			DbUtil.setObject(ps2, 10, Types.VARCHAR, vo.opt("TARGET_SALESTR2"));
			DbUtil.setObject(ps2, 11, Types.VARCHAR, vo.opt("TARGET_PSNCODE2"));
			DbUtil.setObject(ps2, 12, Types.VARCHAR, vo.opt("TARGET_PSNNAME2"));
			ps2.execute();
			ps2.close();
			// ��������
			ps2 = conn.prepareStatement(sql);
			// ������������
			PK_ID = SysUtil.getId();
			ps2.setString(1, PK_ID);
			DbUtil.setObject(ps2, 2, Types.VARCHAR, vo.opt("SCHEMNO"));
			DbUtil.setObject(ps2, 3, Types.VARCHAR, vo.opt("SCHEMNAME"));
			DbUtil.setObject(ps2, 4, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps2, 5, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps2, 6, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps2, 7, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps2, 8, Types.VARCHAR, FK_ID);
			DbUtil.setObject(ps2, 9, Types.VARCHAR, vo.opt("TARGET_MONTH3"));
			DbUtil.setObject(ps2, 10, Types.VARCHAR, vo.opt("TARGET_SALESTR3"));
			DbUtil.setObject(ps2, 11, Types.VARCHAR, vo.opt("TARGET_PSNCODE3"));
			DbUtil.setObject(ps2, 12, Types.VARCHAR, vo.opt("TARGET_PSNNAME3"));
			ps2.execute();
			ps2.close();
			// �ر���Դ
			// �ύ����
			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			conn.commit();
			System.out.println(f);

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}

	/**
	 * Ƭ�齱���������
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_cust_employee_standrad(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;
		// ��ѭ��ɾ��������ͬ�ķ�����

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// ��ɾ��

			sql = "DELETE WB_ERP.sbt_cust_employee_standrad where SECHNO=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SECHNO"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		conn.commit();

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			if (Double.parseDouble(vo.get("XS").toString()) < 0) {
				throw new Exception("�������������С��0");
			}
			if (Double.parseDouble(vo.get("STATUS").toString()) < 0) {
				throw new Exception("�ڸ�ϵ��������С��0");
			}
			if (Double.parseDouble(vo.get("BL").toString()) < 0) {
				throw new Exception("Ԥ������������С��0");
			}

			if (Double.parseDouble(vo.get("XS").toString()) > 1) {
				throw new Exception("����������������1");
			}
			if (Double.parseDouble(vo.get("STATUS").toString()) > 1) {
				throw new Exception("�ڸ�ϵ�����������1");
			}
			if (Double.parseDouble(vo.get("BL").toString()) > 1) {
				throw new Exception("Ԥ���������������1");
			}

			sql = "insert into wb_erp.sbt_cust_employee_standrad (ID,SECHNO,	SECHNAME,	JSMONTH,	ZHANGQ,	YXB,	SALESTR,	PSNCODE,	PSNNAME,	STATUS,	XS,	BL) "
					+ " values (?,?,?,	?,	?,	?,	?,	?,	?,	?,	?,	?)";

			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SECHNO"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("SECHNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("JSMONTH"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("SALESTR"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PSNCODE"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("PSNNAME"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("STATUS"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("XS"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("BL"));

			ps.execute();
			// �ύ����

			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * �ͻ�����-��������1
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_SBT_CUST_EMPLOYEE_SCHEME(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;
		// ��ѭ��ɾ��������ͬ�ķ�����

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// ��ɾ��

			sql = "DELETE from  WB_ERP.SBT_CUST_EMPLOYEE_SCHEME  where SECHNO=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SECHNO"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		conn.commit();
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			System.out.println(f);
			sql = "insert into WB_ERP.SBT_CUST_EMPLOYEE_SCHEME (ID,SECHNO,	SECHNAME,	JSDATE,	ZHANGQ,	YXB,	SALESTR,	UNIT,	STANDARD,	BASENUMBER,	HIGHERNUMBER,	LOWERNUMBER,	BASEGDPRICE,	BASEPRICE,	XS,	HIGHERPRICE,	HIGHERMAXPRICE) "
					+ " values (?,?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?)";

			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SECHNO"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("SECHNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("JSDATE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("SALESTR"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("UNIT"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("STANDARD"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("BASENUMBER"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("HIGHERNUMBER"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("LOWERNUMBER"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("BASEGDPRICE"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("BASEPRICE"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("XS"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("HIGHERPRICE"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("HIGHERMAXPRICE"));

			ps.execute();
			// �ύ����

			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * �ͻ�����������Ŀ��
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_saletarget_customer(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		int result2 = 0;
		String strcusttype = "";
		String strcpxx = "";
		String strcustcode = "";
		String strzhangq = "";
		String stryxb = "";
		ResultSet rSet = null;
		PreparedStatement ps1 = null;

		// �ж��Ƿ�����ˣ���˵����ݾ��������޸�
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			strcusttype = vo.get("CUSTTYPE").toString().trim();

			sql = "select status as CT from wb_erp.saletarget_customer a  where monthyear=? and rownum=1";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTHYEAR"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 1) {
				throw new Exception("���������" + f + "�����������ڼ�����ˣ��������������޸ģ�");
			}
			DbUtil.closeStatement(ps1);
		}
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			strcusttype = vo.get("CUSTTYPE").toString().trim();
			strcustcode = vo.get("CUSTCODE").toString().trim();
			strzhangq = vo.get("ZHANGQ").toString().trim();
			stryxb = vo.get("YXB").toString().trim();
			strcpxx = vo.get("CPXX").toString().trim();

			if (strcusttype.equals("Ǳ���ͼ�ֵ�ͻ�") || strcusttype.equals("ά���ͼ�ֵ�ͻ�")
					|| strcusttype.equals("һ��ͻ�")
					|| strcusttype.equals("�¿���ֵ�ͻ�")) {
				System.out.println(f + "У��ͨ��");
				if (strcusttype.equals("Ǳ���ͼ�ֵ�ͻ�")
						|| strcusttype.equals("ά���ͼ�ֵ�ͻ�")) {
					if (strcustcode.equals("")) {
						throw new Exception("��" + f + "�пͻ����벻����Ϊ��");
					} else {
						sql = "select count(1) as CT from wb_erp.bd_custcode_view  where custcode1=?";
						ps1 = conn.prepareStatement(sql);
						DbUtil.setObject(ps1, 1, Types.VARCHAR, vo
								.get("CUSTCODE"));
						rSet = ps1.executeQuery();
						if (rSet.next()) {
							result2 = rSet.getInt("CT");
						}
						if (result2 == 0) {
							throw new Exception("���������" + f + 1 + "��"
									+ vo.get("CUSTCODE").toString()
									+ "һ���ͻ����벻���ڣ���������µ��룡");
						}
						DbUtil.closeStatement(ps1);
					}
				}
			}

			else {
				System.out.print(strcusttype);
				throw new Exception("��" + f + "�пͻ�����У�鲻ͨ��");
			}
			if (!strzhangq.equals("") && !stryxb.equals("")) {
				sql = "select count(1) as CT from WB_ERP.org_orgs a  where name  in (?,?)";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ZHANGQ"));
				DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("YXB"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 != 2) {
					throw new Exception("���������" + f + 1
							+ "��ս����Ӫ���������ڣ���������µ��룡");
				}
				DbUtil.closeStatement(ps1);
			} else {
				throw new Exception("��" + f + "��ս����Ӫ����������Ϊ��");
			}

			if (strcpxx.equals("����") || strcpxx.equals("ֱ��")
					|| strcpxx.equals("����") || strcpxx.equals("����")
					|| strcpxx.equals("OEM") || strcpxx.equals("Ԥ����"))
				System.out.println("У��ͨ��");
			else
				throw new Exception("��" + f + "�в�Ʒ����У�鲻ͨ��");

			// ��ɾ��
			if ("10".equals(imptype))
				sql = "DELETE WB_ERP.saletarget_customer WHERE custcode = ? and MONTHYEAR =? and cpxx=? and salestr=?";
			if ("11".equals(imptype))
				sql = "DELETE WB_ERP.costtarget_customer WHERE custcode = ? and MONTHYEAR =? and cpxx=? and salestr=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("CUSTCODE"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("MONTHYEAR"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("CPXX"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("SALESTR"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
			conn.commit();
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			if (vo.get("CUSTTYPE").toString().trim().equals("�¿���ֵ�ͻ�")) {
				if (vo.get("SALESTR").toString().trim().equals("")) {
					throw new Exception("��" + f + "�У��¿���ֵ�ͻ���Ƭ��Ϊ���������д�����µ��룡");
				}
			}
			if ("10".equals(imptype))
				sql = "insert into WB_ERP.saletarget_customer (PK_SALETARGET, CUSTCODE, MONTHYEAR, TARGET , COPERATOR,ZHANGQ,YXB,CUSTTYPE,MONEYOFPROMTION,MONEYOFBANK,CPXX,TS,SALESTR) "
						+ " values (?, ?, ?, ?, ?,?,?,?,?,?,?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?)";
			if ("11".equals(imptype))
				sql = "insert into WB_ERP.costtarget_customer (PK_SALETARGET, CUSTCODE, MONTHYEAR, TARGET , COPERATOR,ZHANGQ,YXB,CUSTTYPE,MONEYOFPROMTION,MONEYOFBANK,CPXX,ts) "
						+ " values (?, ?, ?, ?, ?,?,?,?,?,?,?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'))";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("MONTHYEAR"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("TARGET"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("CUSTTYPE"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("MONEYOFPROMTION"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("MONEYOFBANK"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("CPXX"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("SALESTR"));
			ps.execute();
			// �ύ����
			System.out.println(f);
			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * ����ͻ�-�������б�
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_cust_specialcust(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;
		// ��ѭ��ɾ��������ͬ�ķ�����
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// ��ɾ��
			sql = "DELETE WB_ERP.sbt_cust_specialcust WHERE custcode = ? and startmonth =? and endmonth=? and type=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("CUSTCODE"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("STARTMONTH"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ENDMONTH"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("TYPE"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		conn.commit();
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			sql = "insert into WB_ERP.sbt_cust_specialcust (ID, CUSTCODE, CUSTNAME, STARTMONTH, ENDMONTH, RATE, TYPE, TS, COPERATOR) "
					+ " values (?, ?, ?, ?, ?, ?, ?, TO_CHAR(SYSDATE,'yyyy-mm-dd hh24:mi:ss'), ?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("STARTMONTH"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ENDMONTH"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("RATE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, request
					.getAttribute("sys.userName"));

			ps.execute();
			// �ύ����

			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * �ͻ����ս�
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_cust_ht_year(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;
		// ��ѭ��ɾ��������ͬ�ķ�����
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// ��ɾ��
			sql = "DELETE WB_ERP.sbt_cust_ht_year WHERE htno = ?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("HTNO"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		conn.commit();
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// У�������
			sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVBADOCGROUP"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("INVBADOCGROUP").toString()
						+ "�����ϲ����ڣ���������µ��룡");
			}
			DbUtil.closeStatement(ps1);
			// У����������Ƿ����
			if (!vo.get("SALETYPE").equals("����")
					&& !vo.get("SALETYPE").equals("���")) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("SALETYPE").toString() + "�������Ͳ����ڣ���������µ��룡");
			}

			// System.out.println(f);
			// У���׼�Ƿ����

			sql = "select  count(1) as CT FROM WB_ERP.SBT_CUST_STANDRAD B WHERE B.STANDRADNAME=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SCHEMEGROUP"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("SCHEMEGROUP").toString()
						+ "������׼�����ڣ���������µ��룡");
			}
			DbUtil.closeStatement(ps1);

			sql = "INSERT INTO  WB_ERP.sbt_cust_ht_year(id,HTNO,  HTTYPE,  CUSTCODE,  CUSTNAME,  STARTDATE,  ENDDATE,  INVBADOCGROUP,  SCHEMEGROUP,  TYPE0,  TYPE7,  TYPE6,  TYPE5,  TYPE1,  TYPE2,  TYPE3,  TYPE4,	TYPE8,	TYPE9,	TYPE10,	SALETYPE) "
					+ " VALUES(?,?,	?,	?,	?,	?,	?,	(select id from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?),	(select  id FROM WB_ERP.SBT_CUST_STANDRAD B WHERE B.STANDRADNAME=? and rownum=1),	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("HTNO"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("HTTYPE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("STARTDATE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ENDDATE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("INVBADOCGROUP"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("SCHEMEGROUP"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("TYPE0"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("TYPE7"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("TYPE6"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("TYPE5"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("TYPE1"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("TYPE2"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("TYPE3"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("TYPE4"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("TYPE8"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("TYPE9"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("TYPE10"));
			DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("SALETYPE"));

			/*
			 * DbUtil.setObject(ps, 7, Types.VARCHAR, request
			 * .getAttribute("sys.userName"));
			 */

			ps.execute();
			// �ύ����

			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * ר����ͬǩ�������
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_cust_ht_sign(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;
		// ��ѭ��ɾ��������ͬ�ķ�����
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// ��ɾ��
			sql = "DELETE wb_erp.sbt_cust_ht_sign WHERE htno = ?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("HTNO"));
			ps1.executeUpdate();
			conn.commit();
			DbUtil.closeStatement(ps1);
			ps1.close();
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// У���ͬ�Ƿ����

			sql = "select count(1) as CT from wb_erp.sbt_cust_ht where HTNO=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("HTNO"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("HTNO").toString() + "��ͬ��Ų����ڣ���������µ��룡");
			}
			ps1.close();
			rSet.close();

			sql = "insert into  WB_ERP.sbt_cust_ht_sign(id,HTNO,  HTTYPE,  CUSTCODE,  CUSTNAME,  STARTDATE,  ENDDATE,  TYPE1,  TYPE2,	TYPE3,	TYPE4,	TYPE5,	TYPE6,TYPE7,TYPE8) "
					+ " values(?,?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("HTNO"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("HTTYPE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("STARTDATE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ENDDATE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("TYPE1"));
			/*
			 * DbUtil.setObject(ps, 7, Types.VARCHAR, request
			 * .getAttribute("sys.userName"));
			 */
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("TYPE2"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("TYPE3"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("TYPE4"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("TYPE5"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("TYPE6"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("TYPE7"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("TYPE8"));
			ps.execute();
			// �ύ����
			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			ps.close();
			DbUtil.closeStatement(ps);
			System.out.println(f);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * ר����ͬ
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_SBT_CUST_HT(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;
		// ��ѭ��ɾ��������ͬ�ķ�����
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// ��ɾ��
			sql = "DELETE WB_ERP.SBT_CUST_HT WHERE HTNO = ?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("HTNO"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// У�������
			sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVBADOCGROUP"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("PK_PRODLINE").toString()
						+ "�����ϲ����ڣ���������µ��룡");
			}
			// У���׼�Ƿ����

			sql = "select  count(1) as CT FROM WB_ERP.SBT_CUST_STANDRAD B WHERE B.STANDRADNAME=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SCHEMEGROUP"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("SCHEMEGROUP").toString()
						+ "������׼�����ڣ���������µ��룡");
			}

			// У�鹫˾�Ƿ����
			if (!vo.get("PK_CORP").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("PK_CORP"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�в����ڣ���������µ��룡");
				}
			}

			sql = "insert into wb_erp.sbt_cust_ht (ID, HTNO, JSZQ, PK_CORP, INVBADOCGROUP, SCHEMEGROUP, JSFS, TS, COPERATOR,SALETYPE) "
					+ " values (?,?, ?, (select pk_corp FROM wb_erp.bd_corp p where p.memo=?), (select id from WB_ERP.sbt_cust_invbasdocgroup b  where b.nvbasdocgroup=?), (select id from WB_ERP.sbt_cust_standrad b  where b.standradname=? and rownum=1), ?, to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'), ?,?)";

			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("HTNO"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("JSZQ"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("PK_CORP"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("INVBADOCGROUP"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("SCHEMEGROUP"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("JSFS"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("SALETYPE"));
			ps.execute();
			// �ύ����

			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * �ͻ���������
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_SBT_CUST_SCHEME(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		String strtmp = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;
		// ��ѭ��ɾ��������ͬ�ķ�����
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// ��ɾ��
			sql = "DELETE wb_erp.sbt_cust_scheme WHERE schemeno = ?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SCHEMENO"));
			ps1.executeUpdate();
			conn.commit();
			DbUtil.closeStatement(ps1);
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// �ͻ����

			sql = "select count(1) as CT from wb_erp.SBT_CUST_CUSTGROUP A WHERE A.CUSTGROUPNAME =?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("CUSTOMERGROUP"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("CUSTOMERGROUP").toString()
						+ "�ͻ���ϲ����ڣ���������µ��룡");
			}
			// У�������
			sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVBADOCGROUP"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("INVBADOCGROUP").toString()
						+ "�����ϲ����ڣ���������µ��룡");
			}

			// У�������1
			if (!vo.get("INVBADOCGROUP1").equals("")) {
				sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo
						.get("INVBADOCGROUP1"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + 1 + "��"
							+ vo.get("INVBADOCGROUP1").toString()
							+ "��Ʒ��1���Ʋ����ڣ���������µ��룡");
				}
			}

			// У�������2
			if (!vo.get("INVBADOCGROUP2").equals("")) {
				sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo
						.get("INVBADOCGROUP2"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + 1 + "��"
							+ vo.get("INVBADOCGROUP2").toString()
							+ "��Ʒ��2���Ʋ����ڣ���������µ��룡");
				}
			}
			// У�������3
			if (!vo.get("INVBADOCGROUP3").equals("")) {
				sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo
						.get("INVBADOCGROUP3"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + 1 + "��"
							+ vo.get("INVBADOCGROUP3").toString()
							+ "��Ʒ��3���Ʋ����ڣ���������µ��룡");
				}
			}
			// У�������4
			if (!vo.get("INVBADOCGROUP4").equals("")) {
				sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo
						.get("INVBADOCGROUP4"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + 1 + "��"
							+ vo.get("INVBADOCGROUP4").toString()
							+ "��Ʒ��4���Ʋ����ڣ���������µ��룡");
				}
			}

			// У����������Ƿ����
			if (!vo.get("DJTYPE").equals("����")
					&& !vo.get("DJTYPE").equals("���")) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("DJTYPE").toString() + "������׼���Ͳ����ڣ���������µ��룡");
			}

			// У��ͻ�������ʽ
			if (!vo.get("CONDITIONJLFS").equals("")
					&& !vo.get("CONDITIONJLFS").equals("��������")
					&& !vo.get("CONDITIONJLFS").equals("����Ʒ���ִ��")
					&& !vo.get("CONDITIONJLFS").equals("��������")
					&& !vo.get("CONDITIONJLFS").equals("����������Ԥ����")
					&& !vo.get("CONDITIONJLFS").equals("����������S4011")
					&& !vo.get("CONDITIONJLFS").equals("ֱ������")
					&& !vo.get("CONDITIONJLFS").equals("ֱ������")

			) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("CONDITIONJLFS").toString()
						+ "������ʽ�����ڣ���������µ��룡");
			}

			// У���׼�Ƿ����

			sql = "select  count(1) as CT FROM WB_ERP.SBT_CUST_STANDRAD B WHERE B.STANDRADNAME=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SCHEMEGROUP"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("SCHEMEGROUP").toString()
						+ "������׼�����ڣ���������µ��룡");
			}

			/*
			 * У���ĸ�������ʽ
			 */
			for (int i = 0; i < 5; i++) {
				if (i != 3 && i != 6) {
					strtmp = vo.get("CONDITION" + i).toString();
					if (!strtmp.equals("")) {
						if (!strtmp.equals("�����ƽ�") && !strtmp.equals("�������ƽ�")
								&& !strtmp.equals("���������ƽ�")) {
							throw new Exception("���������" + f + 1 + "��"
									+ vo.get("CONDITION" + i).toString()
									+ "������");

						}
					}
				}
			}

			// У�鹫˾�Ƿ����
			if (!vo.get("PK_CORP").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("PK_CORP"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�й�˾�����ڣ���������µ��룡");
				}
			}

			sql = "insert into wb_erp.sbt_cust_scheme(ID,SCHEMENO,  SCHEMENAME,  PK_CORP,  TYPE,  STARTDATE,  ENDDATE,  DETAILNAME,  CUSTOMERGROUP,  SCHEMEGROUP,  DJTYPE,  INVBADOCGROUP,  INVBADOCGROUP1,  INVBADOCGROUP2,  INVBADOCGROUP3,  INVBADOCGROUP4,  CONDITIONJLFS,  CONDITION0,  CONDITION1,  CONDITION2,  CONDITION4,	CONDITION5,TS,COPERATOR,flag,STARTDATE_JS,ENDDATE_JS,RATE,RATE_CONDITION) "
					+ "values(?,?,  ?,  (SELECT PK_CORP FROM WB_ERP.BD_CORP WHERE MEMO=? ),  ?,  ?,  ?,  ?,  (SELECT ID FROM WB_ERP.SBT_CUST_CUSTGROUP A WHERE A.CUSTGROUPNAME =?),  (SELECT ID FROM WB_ERP.SBT_CUST_STANDRAD B WHERE B.STANDRADNAME=? and rownum=1),  ?,  (SELECT ID FROM WB_ERP.SBT_CUST_INVBASDOCGROUP B WHERE B.NVBASDOCGROUP=?),  (SELECT ID FROM WB_ERP.SBT_CUST_INVBASDOCGROUP B WHERE B.NVBASDOCGROUP=?),  (SELECT ID FROM WB_ERP.SBT_CUST_INVBASDOCGROUP B WHERE B.NVBASDOCGROUP=?),  (SELECT ID FROM WB_ERP.SBT_CUST_INVBASDOCGROUP B WHERE B.NVBASDOCGROUP=?),  (SELECT ID FROM WB_ERP.SBT_CUST_INVBASDOCGROUP B WHERE B.NVBASDOCGROUP=?),  ?,  ?,  ?,  ?,  ?,	?,TO_CHAR(SYSDATE,'YYYY-MM-DD HH24:MI:SS'),?,0,?,?,?,?)";

			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SCHEMENO"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("SCHEMENAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("PK_CORP"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("STARTDATE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ENDDATE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("DETAILNAME"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("CUSTOMERGROUP"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("SCHEMEGROUP"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("DJTYPE"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("INVBADOCGROUP"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("INVBADOCGROUP1"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("INVBADOCGROUP2"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("INVBADOCGROUP3"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("INVBADOCGROUP4"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("CONDITIONJLFS"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("CONDITION0"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("CONDITION1"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("CONDITION2"));
			DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("CONDITION4"));
			DbUtil.setObject(ps, 22, Types.VARCHAR, vo.opt("CONDITION5"));
			DbUtil.setObject(ps, 23, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 24, Types.VARCHAR, vo.opt("STARTDATE_JS"));
			DbUtil.setObject(ps, 25, Types.VARCHAR, vo.opt("ENDDATE_JS"));
			DbUtil.setObject(ps, 26, Types.VARCHAR, vo.opt("RATE"));
			DbUtil.setObject(ps, 27, Types.VARCHAR, vo.opt("RATE_CONDITION"));
			ps.execute();
			conn.commit();
			// �ύ����
			System.out.println(f);
			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * ������Ա��������
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_employee_scheme_kf(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;
		// ��ѭ��ɾ��������ͬ�ķ�����
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// ��ɾ��
			sql = "DELETE wb_erp.sbt_employee_scheme_kf WHERE SECHEMNO = ?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SECHEMNO"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// У������Ʒ��

			sql = "select count(1) as CT from wb_erp.bd_prodline a where a.prodlinename=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("PK_PRODLINE"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("PK_PRODLINE").toString()
						+ "�����Ʒ�߲����ڣ���������µ��룡");
			}
			// У�������
			sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVBASDOCID"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("PK_PRODLINE").toString()
						+ "�����ϲ����ڣ���������µ��룡");
			}
			// У����������Ƿ����
			if (!vo.get("SALETYPE").equals("����")
					&& !vo.get("SALETYPE").equals("���")) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("SALETYPE").toString() + "�������Ͳ����ڣ���������µ��룡");
			}
			sql = "insert into wb_erp.sbt_employee_scheme_kf (ID, SECHEMNAME, ZHANGQ, YXB, PK_PRODLINE, STARTMONTH, ENDMONTH, COUNTMONTH, SALETYPE, SALESTART, SALEEND, INVBASDOCID, MONEY_YWY, MONEY_ZZ, SECHEMNO)"
					+ " values (?, ?, ?, ?, (select PK_PRODLINE from  wb_erp.bd_prodline where prodlinename=?), ?, ?, ?, ?, ?,?, (select id from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?), ?, ?, ?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SECHEMNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PK_PRODLINE"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("STARTMONTH"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ENDMONTH"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("COUNTMONTH"));
			/*
			 * DbUtil.setObject(ps, 7, Types.VARCHAR, request
			 * .getAttribute("sys.userName"));
			 */
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("SALETYPE"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("SALESTART"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("SALEEND"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("INVBASDOCID"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("MONEY_YWY"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("MONEY_ZZ"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("SECHEMNO"));
			ps.execute();
			// �ύ����

			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * ������Ա�������������������������ͬһ�����
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_employee_scheme_kf_check(
			List<JSONObject> voList, HttpServletRequest request,
			HttpServletResponse response, String imptype)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;

		// ��ѭ��ɾ��������ͬ�ķ�����
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// ��ɾ��
			if (imptype.equals("2"))
				sql = "DELETE wb_erp.sbt_employee_scheme_kf_check1 where SECHEMNO=?";
			if (imptype.equals("3"))
				sql = "DELETE wb_erp.sbt_employee_scheme_kf_check2 where SECHEMNO=?";
			if (imptype.equals("4"))
				sql = "DELETE wb_erp.sbt_employee_scheme_kf_check3 where SECHEMNO=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SECHEMNO"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// У������Ʒ��

			sql = "select count(1) as CT from wb_erp.bd_prodline a where a.prodlinename=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("PK_PRODLINE"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("PK_PRODLINE").toString()
						+ "�����Ʒ�߲����ڣ���������µ��룡");
			}
			// У�������
			if (vo.get("PK_INVBASDOCGROUP") != null) {
				if (!vo.get("PK_INVBASDOCGROUP").toString().equals("")) {
					sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
					ps1 = conn.prepareStatement(sql);
					DbUtil.setObject(ps1, 1, Types.VARCHAR, vo
							.get("PK_INVBASDOCGROUP"));
					rSet = ps1.executeQuery();
					if (rSet.next()) {
						result2 = rSet.getInt("CT");
					}
					if (result2 == 0) {
						throw new Exception("���������" + f + 1 + "��"
								+ vo.get("PK_PRODLINE").toString()
								+ "�����ϲ����ڣ���������µ��룡");
					}
				}
			}
			// У����������Ƿ����
			if (!vo.get("SALETYPE").equals("����")
					&& !vo.get("SALETYPE").equals("���")) {
				throw new Exception("���������" + f + 1 + "��"
						+ vo.get("SALETYPE").toString() + "�������Ͳ����ڣ���������µ��룡");
			}
			// �Ƿ���ڷ���
			sql = "select count(1) as CN from wb_erp.sbt_employee_scheme_kf  t where t.zhangq=? and  pk_prodline=(select pk_prodline from wb_erp.bd_prodline where prodlinename=?) and t.startmonth=? and t.endmonth=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ZHANGQ"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("PK_PRODLINE"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("STARTMONTH"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ENDMONTH"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CN");
			}
			if (result2 == 0) {
				throw new Exception("���������" + f + 1
						+ "���Ҳ�����Ӧ�Ľ��������������ȵ��뿪�������������ٵ������������ ");
			}

			if (imptype.equals("2")) {
				// sbt_employee_scheme_kf_check1 �������ڻ�������д�����ݿ� by wenshixian
				// 2017-11-20
				sql = "insert into wb_erp.sbt_employee_scheme_kf_check1 (ID,  SCHEMNAME,ZHANGQ, YXB, PK_PRODLINE, STARTMONTH, ENDMONTH, COUNTNUMBER, PK_INVBASDOCGROUP, SALETYPE, COUNTMONTH, RATE, SECHEMNO)"
						+ " values (?, ?, ?, ?,  (select a.pk_prodline from wb_erp.bd_prodline a where prodlinename=?),?, ?,?, (select id from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?), ?, ?, ?, ?)";
				ps = conn.prepareStatement(sql);
				String PK_ID = SysUtil.getId();
				ps.setString(1, PK_ID);
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SECHEMNAME"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ZHANGQ"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("YXB"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PK_PRODLINE"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("STARTMONTH"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ENDMONTH"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("COUNTMONTH"));
				/*
				 * DbUtil.setObject(ps, 7, Types.VARCHAR, request
				 * .getAttribute("sys.userName"));
				 */
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo
						.opt("PK_INVBASDOCGROUP"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("SALETYPE"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("COUNTMONTH"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("RATE"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("SECHEMNO"));
				ps.execute();
			}
			if (imptype.equals("3")) {
				// sbt_employee_scheme_kf_check2 �����ڴ��»�������д�����ݿ� by wenshixian
				// 2017-11-20
				sql = "insert into wb_erp.sbt_employee_scheme_kf_check2 (ID,  SCHEMNAME,ZHANGQ, YXB, PK_PRODLINE, STARTMONTH, ENDMONTH, COUNTNUMBER, PK_INVBASDOCGROUP, SALETYPE, COUNTMONTH, RATE, SECHEMNO,LOWERRATE)"
						+ " values (?, ?, ?, ?,  (select a.pk_prodline from wb_erp.bd_prodline a where a.prodlinename=?),?, ?,?, (select id from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?), ?, ?, ?, ?,?)";
				ps = conn.prepareStatement(sql);
				String PK_ID = SysUtil.getId();
				ps.setString(1, PK_ID);
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SECHEMNAME"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ZHANGQ"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("YXB"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PK_PRODLINE"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("STARTMONTH"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ENDMONTH"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("COUNTMONTH"));
				/*
				 * DbUtil.setObject(ps, 7, Types.VARCHAR, request
				 * .getAttribute("sys.userName"));
				 */
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo
						.opt("PK_INVBASDOCGROUP"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("SALETYPE"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("COUNTMONTH"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("RATE"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("SECHEMNO"));
				DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("LOWERRATE"));
				ps.execute();
			}
			if (imptype.equals("4")) {
				// sbt_employee_scheme_kf_check3 �����ڴ����»�������д�����ݿ� by wenshixian
				// 2017-11-20
				sql = "insert into wb_erp.sbt_employee_scheme_kf_check3 (ID,  SCHEMNAME,ZHANGQ, YXB, PK_PRODLINE, STARTMONTH, ENDMONTH, COUNTNUMBER, PK_INVBASDOCGROUP, SALETYPE, TYPE, RATE, SECHEMNO,KHNUMBER,MEMO)"
						+ " values (?, ?, ?, ?,  (select a.pk_prodline from wb_erp.bd_prodline a where a.prodlinename=?),?, ?,?, (select id from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?), ?, ?, ?, ?,?,?)";
				ps = conn.prepareStatement(sql);
				String PK_ID = SysUtil.getId();
				ps.setString(1, PK_ID);
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SECHEMNAME"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ZHANGQ"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("YXB"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PK_PRODLINE"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("STARTMONTH"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ENDMONTH"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("COUNTMONTH"));
				/*
				 * DbUtil.setObject(ps, 7, Types.VARCHAR, request
				 * .getAttribute("sys.userName"));
				 */
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo
						.opt("PK_INVBASDOCGROUP"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("SALETYPE"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("TYPE"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("RATE"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("SECHEMNO"));
				DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("KHNUMBER"));
				DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("MEMO"));
				ps.execute();
			}
			// �ύ����

			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * ����_������׼
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_cust_standrad(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;
		// ��ѭ������׼����Ψһֵ�ҳ���
		List<String> STANDRADNAME = new ArrayList();
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			if (!STANDRADNAME.contains(vo.get("STANDRADNAME").toString())) {
				STANDRADNAME.add(vo.get("STANDRADNAME").toString());
			}
		}
		// ѭ����׼���ƿ�ϵͳ���Ƿ��Ѵ��ڣ��Ѵ�����ɾ��
		for (int i = 0; i < STANDRADNAME.size(); i++) {
			sql = "DELETE WB_ERP.sbt_cust_standrad WHERE STANDRADNAME ='"
					+ STANDRADNAME.get(i) + "'";
			ps1 = conn.prepareStatement(sql);
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		/*
		 * for (int f = 0; f < voList.size(); f++) { JSONObject vo =
		 * voList.get(f); // ��ɾ�� sql =
		 * "DELETE WB_ERP.sbt_cust_specialcust WHERE custcode = ? and startmonth =? and endmonth=? and type=?"
		 * ; ps1 = conn.prepareStatement(sql); DbUtil.setObject(ps1, 1,
		 * Types.VARCHAR, vo.get("CUSTCODE")); DbUtil.setObject(ps1, 2,
		 * Types.VARCHAR, vo.get("STARTMONTH")); DbUtil.setObject(ps1, 3,
		 * Types.VARCHAR, vo.get("ENDMONTH")); DbUtil.setObject(ps1, 4,
		 * Types.VARCHAR, vo.get("TYPE")); ps1.executeUpdate();
		 * DbUtil.closeStatement(ps1); } conn.commit();
		 */
		for (int i = 0; i < STANDRADNAME.size(); i++) {
			String PK_ID = SysUtil.getId();
			for (int f = 0; f < voList.size(); f++) {
				JSONObject vo = voList.get(f);
				if (vo.opt("STANDRADNAME").equals(STANDRADNAME.get(i))) {
					sql = "insert into WB_ERP.sbt_cust_standrad (ID, STANDRADNAME, STARTNUM, ENDNUM, PRICE, TS, COPERATOR) "
							+ " values (?, ?, ?, ?, ?, TO_CHAR(SYSDATE,'yyyy-mm-dd hh24:mi:ss'), ?)";
					ps = conn.prepareStatement(sql);
					ps.setString(1, PK_ID);
					DbUtil.setObject(ps, 2, Types.VARCHAR, vo
							.opt("STANDRADNAME"));
					// DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("TYPE"));
					DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("STARTNUM"));
					DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ENDNUM"));
					DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PRICE"));
					// DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("FLAG"));
					DbUtil.setObject(ps, 6, Types.VARCHAR, request
							.getAttribute("sys.userName"));

					ps.execute();
					// �ύ����

					// �ر���Դ
					// DbUtil.closeStatement(ps1);
					DbUtil.closeStatement(ps);
				}
			}
		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}
	
	/**
	 * Ԥ��Ŀ�꣨ս����
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_app_main_yjmb_zq(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;
		ResultSet rSet = null;
		int result2 = 0;
		String Times = "(";
		for(int f = 0; f < voList.size(); f++) {
			JSONObject obj = voList.get(f);
			if(Times.indexOf(obj.get("MONTH").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("MONTH").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("MONTH").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_yjmb_zq WHERE MONTH in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_yjmb_zq (zhanq,month,mbnnumber)"
					+ "values (?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("ZHANQ"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("MBNNUMBER"));
			ps.execute();
			// �ύ����
			System.out.println(f);
			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}
	
	/**
	 * Ԥ��Ŀ�꣨������
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_app_main_yjmb_xx(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;
		ResultSet rSet = null;
		int result2 = 0;
		String Times = "(";
		for(int f = 0; f < voList.size(); f++) {
			JSONObject obj = voList.get(f);
			if(Times.indexOf(obj.get("MONTH").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("MONTH").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("MONTH").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_yjmb_xx WHERE MONTH in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_yjmb_xx (cpx,month,mbnnumber)"
					+ "values (?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("MBNNUMBER"));
			ps.execute();
			// �ύ����
			System.out.println(f);
			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}
	
	/**
	 * Ԥ��Ŀ�꣨ս����
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_app_main_yjmb_yxb(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;
		ResultSet rSet = null;
		int result2 = 0;
		String Times = "(";
		for(int f = 0; f < voList.size(); f++) {
			JSONObject obj = voList.get(f);
			if(Times.indexOf(obj.get("MONTH").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("MONTH").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("MONTH").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_yjmb_yxb WHERE MONTH in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_yjmb_yxb (zhanq,YXB,month,mbnnumber)"
					+ "values (?,?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("ZHANQ"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("MBNNUMBER"));
			ps.execute();
			// �ύ����
			System.out.println(f);
			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}

	/**
	 * ����ģ��
	 * 
	 * @param request
	 * @param response
	 * @throws Exception
	 */
	public static void exportFiles(HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		JSONArray ja = new JSONArray(request.getParameter("files"));
		int i, j = ja.length();
		File[] files = new File[j];
		boolean useZip;
		String fileName;

		for (i = 0; i < j; i++)
			files[i] = new File(Main.path, ja.optString(i));
		fileName = files[0].getName();
		useZip = StringUtil.isEqual(request.getParameter("type"), "1") || j > 1
				|| files[0].isDirectory();
		if (j == 1) {
			if (useZip)
				fileName = FileUtil.extractFilenameNoExt(fileName) + ".zip";
		} else {
			File parentFile = files[0].getParentFile();
			if (parentFile == null)
				fileName = "file.zip";
			else
				fileName = parentFile.getName() + ".zip";
		}
		if (fileName.equals(".zip") || fileName.equals("/.zip"))
			fileName = "file.zip";
		response.reset();
		if (!useZip)
			response.setHeader("content-length", Long.toString(files[0]
					.length()));
		response.setHeader("content-type", "application/force-download");
		response.setHeader("content-disposition", "attachment;"
				+ WebUtil.encodeFilename(request, fileName));
		if (useZip) {
			ZipUtil.zip(files, response.getOutputStream());
			response.flushBuffer();
		} else
			WebUtil.response(response, new FileInputStream(files[0]));
	}
}
