package com.sbt.tool;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Types;
import java.text.DecimalFormat;
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

/**
 * EXcel ����
 * 
 * @author Administrator
 * 
 */
public class ExcelImportTool {
	public static void getFile(HttpServletRequest request,
			HttpServletResponse response) throws Exception {

		InputStream in = (InputStream) request.getAttribute("uploadFile");
		String fileName = request.getAttribute("uploadFile__name").toString();
		String fileType = fileName.substring(fileName.lastIndexOf(".") + 1,
				fileName.length());
		String imptype = request.getAttribute("imptype").toString();
		Map<String, String> map = new HashMap<String, String>();
		if ("1".equals(imptype)) { // �г���Χά�� ����

			map.put("��Ա����", "USERCODE");
			map.put("����", "USERNAME");
			map.put("ʡ", "PROVINCE");
			map.put("��", "CITY");
			map.put("��/��", "AREA");
			map.put("��Ʒ��", "PRODUCTLINE");
			map.put("ʧЧ����", "SXRQ"); //���ǿ���� ����ƽ 2018-03-26
			map.put("�Ƿ�ʧЧ", "SFSX"); //���ǿ���� ����ƽ 2018-03-26
		} 
		else if ("2".equals(imptype)) { // ҵ��Ա��ͻ���Ӧ��ϵά�� ����
			map.put("�ͻ�����", "CUSTCODE");
			map.put("�ͻ�����", "CUSTNAME");
			map.put("�����Ʒ��", "RELATIONTYPE");
			map.put("Ӫ����", "YXB");
			map.put("С��", "XQ");
			map.put("Ƭ��", "SALESTR");
			
			map.put("�ͻ��������", "PSNCODE_CUSTMANAGER");
			map.put("�ͻ�������Ա����", "PSNNAME_CUSTMANAGER");
			
			map.put("�����鳤����", "MANAGERCODE");
			map.put("�����鳤", "MANAGERNAME");

			map.put("���濪������1����", "PSNCODE");
			map.put("���濪������1", "PSNNAME");
			map.put("��������1�������", "MEMO1");

			map.put("���濪������2����", "PSNCODE2");
			map.put("���濪������2", "PSNNAME2");
			map.put("��������2�������", "MEMO2");

			map.put("���濪������3����", "PSNCODE3");
			map.put("���濪������3", "PSNNAME3");
			map.put("��������3�������", "MEMO3");

			map.put("�����������", "PSNCODEMANA_TECHNICAL");
			map.put("������������", "PSNNAMEMANA_TECHNICAL");
			
			map.put("������Ա1����", "PSNCODE1_TECHNICAL");
			map.put("������Ա1����", "PSNNAME1_TECHNICAL");
			map.put("������Ա1�������", "TECHNICAL1_RATE");

			map.put("������Ա2����", "PSNCODE2_TECHNICAL");
			map.put("������Ա2����", "PSNNAME2_TECHNICAL");
			map.put("������Ա2�������", "TECHNICAL2_RATE");

			map.put("�Ƿ�ҵ��", "TYPE");
			map.put("��ע", "MEMO");

			//map.put("�Ͽͻ���������", "MONTH1");
			map.put("�Ͽͻ�����·�", "MONTH2");
			map.put("ĸ��ͷ��", "MPIG");
			map.put("����ͷ��", "RPIG");
			//
			Connection conn = DbUtil.getConnection();
			PreparedStatement ps1 = null;
			ResultSet rSet = null;
			int result = 0;
			// ��ѯ�Ƿ����
			String sql = "select count(*) CT from wb_erp.APP_ZBXX_USER where MANGERCODE = ?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result = rSet.getInt("CT");
			}
			// �ر���Դ
			DbUtil.closeResultSet(rSet);
			DbUtil.closeStatement(ps1);
			DbUtil.closeConnection(conn);
			if (result == 0) {
				throw new Exception("��ǰ��¼�ˣ���Ȩ�޵��룡");
			}

		} else if ("3".equals(imptype)) {
			map.put("�ͻ�����", "KHBM");
			map.put("Ӫ����", "YXB");
			map.put("�ͻ�����", "KHXM");
			map.put("�ͻ��绰", "KHDH");
			map.put("ע��ʱ��", "ZCSJ");
			map.put("��3�¾����", "JBD3");
			map.put("�����±��", "JBD2");
			map.put("�����������", "FXDKYE");
			map.put("ֱ���������", "ZXDKYE");
			map.put("���ÿ����", "XYKYE");
			map.put("�������Ŷ��", "FXSXED");
			map.put("��ǰ�ɴ����", "DQKDED");
		} else if ("4".equals(imptype)) {
			map.put("�ͻ���ʶ", "KHBS");
			map.put("��ǩ����", "MQYH");
			map.put("����", "DQ");
			map.put("ʡ��", "SQ");
			map.put("Ӫ����", "YXB");
			map.put("��������", "GCMC");
			map.put("�ͻ�����", "KHLX");
			map.put("�ͻ�����", "KHBM");
			map.put("�ͻ�����", "KHXM");
			map.put("��������", "SZXS");
			map.put("�����ܶ��Ԫ��", "JYZE");
			map.put("�Ƽ���ȣ���Ԫ��", "TJED");
			map.put("���׶��ڼ�", "JYEQJ");
			map.put("�绰", "DH");
			map.put("���֤��", "SFZH");
			map.put("�������ޣ�ϵͳʱ�䣩", "HZNX");
			map.put("Ӫҵִ�պ�", "YYZZH");
			map.put("�ֳ���ǩ���", "XCMQJG");
			map.put("ԭ��˵��", "YYSM");
			map.put("�ֳ���ǩ��", "XCMQR");
			map.put("��ǩ����", "MQRQ");
			map.put("������ȫ����", "CLQQRQ");
			map.put("�Ŵ�����", "XDSX");
			map.put("�������", "WTLB");
			map.put("�Ŵ����˵��", "XDSHSM");
			map.put("�Ƽ�����", "TJRQ");
			map.put("�Ƽ������", "TJHBH");
			map.put("�����", "HKKH");
			map.put("�ſ�����", "FKRQ");
			map.put("�ſ���", "FKJE");
			map.put("�ſ�����", "FKLL");
			map.put("�ſ�����/��", "FKQS");
			map.put("NC�տ����", "NCSKDRQ");
			map.put("�Ƿ�绰֪ͨ���", "SFDHTZTH");
			map.put("�����", "WHR");
			map.put("�������", "WHRQ");
			map.put("��ע", "BZ");
			map.put("Ӧ�ſ�ʱ��", "YFKSJ");
			map.put("��������", "CCTS");
			map.put("�ſ��ʱ", "FKHS");
			map.put("��ǰ����ʱ��", "TQHKSJ");
			map.put("Ӧ����ʱ��", "YHKSJ");
			map.put("��ǰ���", "DQYE");
		} else if ("5".equals(imptype)) {
			map.put("�ͻ�����", "KHLX");
			map.put("��������", "DKYH");
			map.put("�����г�", "SSSC");
			map.put("����", "DQ");
			map.put("����", "JD");
			map.put("Ӫ����", "YXB");
			map.put("��������", "GCMC");
			map.put("�ͻ�����", "KHBM");
			map.put("�ͻ�����", "KHXM");
			map.put("�ͻ��绰", "KHDH");
			map.put("��������", "SZXS");
			map.put("�״��������", "SCTHRQ");
			map.put("���֤��", "SFZH");
			map.put("����ĸ��ͷ��", "CLMZTS");
			map.put("����ɼ�/ͷ", "YZCJ");
			map.put("��������/��", "YZNX");
			map.put("�Ƽ���ȣ���Ԫ��", "TJED");
			map.put("ϵͳע��ʱ��", "XTZCSJ");
			map.put("��һ�꽻�׶��Ԫ��", "JYNJYE");
			map.put("����ĸ��ͷ��", "JCMZTS");
			map.put("��������ͷ��", "CLRZTS");
			map.put("��ǰ���и�ծ", "DQYHFZ");
			map.put("��ǰ�����", "DQMJFZ");
			map.put("Э��������������/�꣩", "XYGLDS");
			map.put("����OA����", "JSOARQ");
			map.put("���������", "DSHRQ");
			map.put("��ǩ��Ա", "MQRY");
			map.put("��ʵ��ǩ����", "HSMQRQ");
			map.put("�Ŵ�����", "XDSX");
			map.put("�������", "WTLB");
			map.put("�Ŵ����˵��", "XDSHSM");
			map.put("�Ƽ�����", "TJRQ");
			map.put("�Ƽ������", "TJHBH");
			map.put("�Ƿ񲹲���", "SFBQCL");
			map.put("������ȫ����", "CLQQRQ");
			map.put("�Ƿ�����", "SFTP");
			map.put("��������", "TPFL");
			map.put("����˵��", "TPSM");
			map.put("�ۼ�����", "KJTS");
			map.put("������������", "XQMZTS");
			map.put("����Ԥ��", "JDYJ");
			map.put("�ſ���", "FKJE");
			map.put("�ſ�����", "FKRQ");
			map.put("NC��������", "NCXZRQ");
			map.put("�Ƿ񽻱�֤��", "SFJBZJ");
			map.put("�ſ�����/��", "FKQS");
			map.put("�����̱���", "WLSBM");
			map.put("����������", "WLSMC");
			map.put("�绰", "DH");
			map.put("���������֤��", "WLSSFZH");
			map.put("���и�ծ��", "YHFZE");
			map.put("����", "JZ");
			map.put("����", "BZ");
			map.put("ҵ��Ա", "YWY");
			map.put("Ӧ����������", "YHBJRQ");
			map.put("ʵ�ʽ��屾������", "SJJQBJRQ");
			map.put("�����", "HKKH");
			map.put("��������", "DKLL");
			map.put("�Ŵ��ƽ��·�", "XDTJYF");

		} else if ("6".equals(imptype)) {
			map.put("��������", "BLYH");
			map.put("Ӫ����", "YXB");
			map.put("�ͻ�����", "KHBM");
			map.put("�ͻ�����", "KHXM");
			map.put("�ſ���", "FKJE");
			map.put("������������", "XLJSQJ");
			map.put("��������", "XLJS");
			map.put("����Ŀ��", "ZLMB");
			map.put("��ǰʵ������", "DQSJXL");
			map.put("ȫ�Ϣ����", "QEBXXL");

		} else if ("7".equals(imptype)) {
			map.put("��Ϣ����", "BXFA");
			map.put("��Ϣ�·�", "BXYF");
			map.put("��������", "BLYH");
			map.put("Ӫ����", "YXB");
			map.put("��������", "GCMC");
			map.put("�ͻ�����", "KHBM");
			map.put("�ͻ�����", "KHXM");
			map.put("�绰", "DH");
			map.put("���֤��", "SFZH");
			map.put("��ǩ����", "MQRQ");
			map.put("�ſ�����", "FKRQ");
			map.put("�ſ���", "FKJE");
			map.put("������Ϣ", "DKLX");
			map.put("�������ʱ��", "HKWBSJ");
			map.put("��ǩǰһ������", "MQQYNXL");
			map.put("��ǩ��һ������", "MQHYNXL");
			map.put("��Ϣ��", "BXL");
			map.put("��Ϣ���", "BXJE");
			map.put("��Ϣ˵��", "BXSM");

		} else if ("8".equals(imptype)) {
			map.put("�ͻ�����", "KHLX");
			map.put("Ӫ����", "YXB");
			map.put("��������", "DKYH");
			map.put("�ͻ�����", "KHBM");
			map.put("�ͻ�����", "KHXM");
			map.put("�ͻ��绰", "KHDH");
			map.put("Ӧ��ʱ��", "YHSJ");
			map.put("Ӧ�����", "YHJE");
			map.put("�����", "HKKH");
		} else if ("9".equals(imptype)) {
			map.put("Ӫ����", "YXB");
			map.put("�ͻ�����", "KHBM");
			map.put("�ͻ�����", "KHXM");
			map.put("ǰ������", "QYXL");
			map.put("��������", "SYXL");
			map.put("��������", "DYXL");
			map.put("������", "DXB");
			map.put("Ԥ��ԭ��", "YJYY");
			map.put("���Ԥ��Ŀ��", "JCYJMB");
			map.put("Ԥ������", "YJCS");
		} else if ("10".equals(imptype)) {
			map.put("����", "DQ");
			map.put("����", "JD");
			map.put("Ӫ����", "YXB");
			map.put("ֱ���г�", "ZXSC");
			map.put("��������", "JRQY");
			map.put("����ϸ���г�", "JRXFSC");
			map.put("��Ա����", "USERCODE");
			map.put("����", "USERNAME");
			map.put("�ҵ��쵼����", "GDLDBM");
			map.put("�ҵ��쵼", "GDLDXM");
			map.put("��������רԱ����", "JRZYBM");
			map.put("��������רԱ", "JRZYXM");
		} else if ("11".equals(imptype)) {// ����������Ϣ����
			map.put("�ͻ�����", "KHLX");
			map.put("�ͻ�����", "KHBM");
			map.put("ս��", "ZQ");
			map.put("Ӫ����", "YXB");
			map.put("�ͻ�����", "KHXM");
			map.put("�ͻ��绰", "KHDH");
			map.put("��������", "SZXS");
			map.put("��������", "BLYH");
			map.put("�����ȣ���Ԫ��", "DKED");
			map.put("�����", "HKKH");
			map.put("�������", "YQLB");
			map.put("���ڿ�ʼʱ��", "YQKSSJ");
			map.put("��������", "YQTS");
			map.put("���ڽ��", "YQJE");
			map.put("�������", "GJQK");
			map.put("Ȩ�滺����ʵ���", "QYHFQK");
			map.put("�Ƿ�׷��", "SFZH");
			map.put("��������", "JQRQ");
			map.put("׷�ؽ��", "ZHJE");
			map.put("�������", "YQYE");
		} else if ("12".equals(imptype)) {// ���ڷŴ�����
			map.put("�Ƽ�����", "TJHBH");
			map.put("�ͻ�����", "KHBM");
			map.put("�ͻ�����", "KHXM");
			map.put("�ſ�����", "FKRQ");
			map.put("�ſ���", "FKJE");
			map.put("�����", "HKKH");
			map.put("NC��������", "NCXZRQ");

		} else if ("13".equals(imptype)) {// ֱ���ͻ����ܿͻ��ƽ�
			map.put("ս��", "DQ");
			map.put("Ӫ����", "YXB");
			map.put("�ͻ�����", "CUSTCODE");
			map.put("�ͻ�����", "CUSTNAME");
			map.put("�ƽ�/�¿�", "TRANSFERTYPE");
			map.put("�ƽ��·�/�¿��·�", "MONTH");
			map.put("�ƽ�ǰ������ԱƬ��", "SALESTRUBEFORE");
			map.put("�ƽ�ǰ������Ա����", "PSNCODEBEFORE");
			map.put("�ƽ�ǰ������Ա����", "PSNNAMEBEFORE");
			map.put("�ƽ���������ԱƬ��", "SALESTRUAFTER");
			map.put("�ƽ���������Ա����", "PSNCODEAFTER");
			map.put("�ƽ���������Ա����", "PSNNAMEAFTER");

		} else if ("14".equals(imptype))// �ͻ��ҿ���ϵά��
		{
			map.put("ģʽ����", "TYPE");
			map.put("�ͻ�����", "CUSTCODE");
			map.put("�ͻ�����", "CUSTNAME");
			map.put("ʵ�����������̱���", "FACTCUSTCODE");
			map.put("ʵ����������������", "FACTCUSTNAME");
		} else if ("15".equals(imptype))// ��Ա����/��Ա��֯��ϵ
		{
			map.put("����", "MONTH");
			// map.put("ϵͳ", "SYS");
			// map.put("������֯", "ORGNAME");
			// map.put("��֯���", "ORGSHORTNAME");
			map.put("��Ա����", "PSNCODE");
			map.put("����", "PSNNAME");
			// map.put("���뼯��ʱ��", "ONJOBTIME");
			map.put("���ڲ���", "DEPTNAME");
			map.put("��λ", "POSTNAME");
			map.put("��ְ����", "OUTTIME");
			map.put("��Ʒ��", "CPX");
			map.put("һ������", "ORG1");
			map.put("��������", "ORG2");
			map.put("��������", "ORG3");
			map.put("�ļ�����Ƭ", "ORG4");
			map.put("��Ա״̬", "STATUS");
			map.put("����ҵ��", "BUSSTYPE");
		} else if ("16".equals(imptype) || "161".equals(imptype))// ��������ϸ-�����籣
																	// ������-н��-����
		{
			map.put("���ù����·�", "MONTH");
			map.put("��Ʒ��", "CPX");
			map.put("һ������", "ORG1");
			map.put("��������", "ORG2");
			map.put("��������", "ORG3");
			map.put("�ļ�����", "ORG4");
			map.put("��Ա����", "PSNCODE");
			map.put("��Ա����", "PSNNAME");
			map.put("����һ����Ŀ", "KM1");
			map.put("���������Ŀ", "KM2");
			map.put("����������Ŀ", "KM3");
			map.put("�����ļ���Ŀ", "KM4");
			map.put("���ý��", "MONEY");
			map.put("��������", "TYPE");
			map.put("����ҵ��", "BUSSTYPE");

		} else if ("17".equals(imptype))// ��������ϸ-н��
		{
			map.put("���ù����·�", "MONTH");
			map.put("��Ʒ��", "CPX");
			map.put("һ������", "ORG1");
			map.put("��������", "ORG2");
			map.put("��������", "ORG3");
			map.put("�ļ�����", "ORG4");
			map.put("��Ա����", "PSNCODE");
			map.put("����", "PSNNAME");
			map.put("�̶�����", "MONEYBASE");
			map.put("��Ч����", "MONEYKPI");
			map.put("���÷�", "MONEYTRAVEL");
			map.put("�г��д���", "MONEYSERVE1");
			map.put("ҵ���д���", "MONEYSERVE2");
			map.put("��Ŀ����", "MONEYPROJECT");
			map.put("ս�Բ���", "MONEYFILLPOST");
			map.put("����", "MONEYKK");
			map.put("����ҵ��", "BUSSTYPE");
		} else if ("18".equals(imptype))// ҵ�����ϸ-�ͻ�
		{
			map.put("���ù����·�", "MONTH");
			map.put("��Ʒ��", "CPX");
			map.put("һ������", "ORG1");
			map.put("��������", "ORG2");
			map.put("��������", "ORG3");
			map.put("�ļ�����", "ORG4");
			map.put("�ͻ�����", "CUSTCODE");
			map.put("�ͻ�����", "CUSTNAME");
			map.put("����һ����Ŀ", "KM1");
			map.put("���������Ŀ", "KM2");
			map.put("����������Ŀ", "KM3");
			map.put("�����ļ���Ŀ", "KM4");
			map.put("���ý��", "MONEY");
			map.put("����˵��", "MEMO");
			map.put("��������", "TYPE");
			map.put("��˾", "UNITSHORTNAME");
			map.put("�������", "INVNAME");
			map.put("��������", "NNUMBER");
			map.put("ϵͳ��Ʊ����", "SYSTYPE");
			map.put("ϵͳ��Ʊ��ע", "SYSMEMO");
			map.put("������������", "HDMB");
			map.put("����ҵ��", "BUSSTYPE");
		}

		else if ("19".equals(imptype))// ��������ϸ-Ӫ������1-��������
		{
			map.put("��������·�", "MONTH");
			map.put("�����·�", "XSMONTH");
			map.put("��Ʒ��", "CPX");
			map.put("һ������", "ORG1");
			map.put("��������", "ORG2");
			map.put("��������", "ORG3");
			map.put("�ļ�����", "ORG4");
			map.put("�ͻ�����", "CUSTCODE");
			map.put("�ͻ�����", "CUSTNAME");
			map.put("�����·�", "KFMONTH");
			map.put("��������", "MONEY");
			map.put("�����ں���˿ۻ�", "MONEYKH");
			map.put("��ע", "MEMO");
			map.put("����ҵ��", "BUSSTYPE");
		} else if ("20".equals(imptype))// ��������ϸ-Ӫ������2-�ͻ�����
		{
			map.put("��������·�", "MONTH");
			map.put("��Ʒ��", "CPX");
			map.put("һ������", "ORG1");
			map.put("��������", "ORG2");
			map.put("��������", "ORG3");
			map.put("�ļ�����", "ORG4");
			map.put("����Ƭ", "SALESTR");
			map.put("��Ա����", "PSNCODE");
			map.put("��Ա����", "PSNNAME");
			map.put("�ͻ����ܽ���", "MONEY");
			map.put("����", "TYPE");
			map.put("����ҵ��", "BUSSTYPE");
		} else if ("21".equals(imptype))// ��������ϸ-Ӫ������3-����Ӫ������
		{
			map.put("��������·�", "MONTH");
			map.put("��Ʒ��", "CPX");
			map.put("һ������", "ORG1");
			map.put("��������", "ORG2");
			map.put("��������", "ORG3");
			map.put("�ļ�����", "ORG4");
			map.put("��Ա����", "PSNCODE");
			map.put("��Ա����", "PSNNAME");
			map.put("�����·�", "KFMONTH");
			map.put("Ӫ��Ա����ϼ�", "SUMMONEY");
			map.put("������", "MONEYGL");
			map.put("Ԥ���Ͻ���", "MONEYYHL");
			map.put("OEM����", "MONEYOEM");
			map.put("פ������Ա����", "MONEYZCJS");
			map.put("��������", "MONEYQT");
			map.put("��������", "MONEYBF");
			map.put("����ҵ��", "BUSSTYPE");
		} else if ("22".equals(imptype))// ��������ϸ
		{
			map.put("���ù����·�", "MONTH");
			map.put("OA��", "REQUESTID");
			map.put("ID��", "WORKFLOWID");
			map.put("��������", "REQUESTNAME");
			map.put("�����˱���", "PSNCODE");
			map.put("������", "PSNNAME");
			map.put("���", "MONEY");
			map.put("֧������", "MEMO");
			map.put("ƾ֤��", "ISTRUE");
			map.put("���˹�˾", "UNITSHORTNAME");
			map.put("һ����Ŀ", "KM1");
			map.put("������Ŀ", "KM2");
			map.put("������Ŀ", "KM3");
			map.put("�ļ���Ŀ", "KM4");
			map.put("��Ʒ��", "CPX");
			map.put("һ������", "ORG1");
			map.put("��������", "ORG2");
			map.put("��������", "ORG3");
			map.put("�ļ�����", "ORG4");
			map.put("�˵����Ƿ�ȡ��", "ISBILL");
			map.put("����ҵ��", "BUSSTYPE");

		} else if ("23".equals(imptype))// Ա����ѵ��
		{
			map.put("���ù����·�", "MONTH");
			map.put("��Ʒ��", "CPX");
			map.put("һ������", "ORG1");
			map.put("��������", "ORG2");
			map.put("��������", "ORG3");
			map.put("�ļ�����", "ORG4");
			map.put("����һ����Ŀ", "KM1");
			map.put("���������Ŀ", "KM2");
			map.put("����������Ŀ", "KM3");
			map.put("�����ļ���Ŀ", "KM4");
			map.put("���ý��", "MONEY");
			map.put("��������", "TYPE");
			map.put("����ҵ��", "BUSSTYPE");

		} else if ("24".equals(imptype))// �������-OA����
		{
			map.put("�����·�", "MONTH");
			map.put("��Ʒ��", "CPX");
			map.put("һ������", "ORG1");
			map.put("��������", "ORG2");
			map.put("��������", "ORG3");
			map.put("�ļ�����", "ORG4");
			map.put("����һ����Ŀ", "KM1");
			map.put("���������Ŀ", "KM2");
			map.put("����������Ŀ", "KM3");
			map.put("�����ļ���Ŀ", "KM4");
			map.put("���ý��", "MONEY");
			map.put("����˵��", "MEMO");
			map.put("����ҵ��", "BUSSTYPE");
			map.put("���˹�˾", "UNITNAME");
			map.put("ƾ֤��", "PZH");
		} else if ("25".equals(imptype))// ���鲹��
		{
			map.put("���ù����·�", "MONTH");
			map.put("��Ʒ��", "CPX");
			map.put("һ������", "ORG1");
			map.put("��������", "ORG2");
			map.put("��������", "ORG3");
			map.put("�ļ�����", "ORG4");
			map.put("����һ����Ŀ", "KM1");
			map.put("���������Ŀ", "KM2");
			map.put("����������Ŀ", "KM3");
			map.put("�����ļ���Ŀ", "KM4");
			map.put("���ý��", "MONEY");
			map.put("��������", "TYPE");
			map.put("����ҵ��", "BUSSTYPE");

		} else if ("26".equals(imptype))// ��������
		{
			map.put("���ù����·�", "MONTH");
			map.put("��������", "DBILLDATE");
			map.put("ժҪ", "MEMO");
			map.put("��˾����", "DEPT");
			map.put("̨��һ��", "KM1TZ");
			map.put("̨�ʶ���", "KM2TZ");
			map.put("�跽", "JF");
			map.put("����", "DF");
			map.put("����ҵ��", "BUSSTYPE");
			map.put("��Ʒ��", "CPX");
			map.put("һ������", "ORG1");
			map.put("��������", "ORG2");
			map.put("��������", "ORG3");
			map.put("�ļ�����", "ORG4");
			map.put("һ����Ŀ", "KM1");
			map.put("������Ŀ", "KM2");
			map.put("������Ŀ", "KM3");
			map.put("�ļ���Ŀ", "KM4");
			map.put("���", "MONEY");

		} else if ("27".equals(imptype))// ����У��
		{
			map.put("���ù����·�", "MONTH");
			map.put("��Ʒ��", "CPX");
			map.put("����ҵ��", "TYPE");
			map.put("һ������", "ORG1");
			map.put("��������", "ORG2");
			map.put("��������", "ORG3");
			map.put("�ļ�����", "ORG4");
		} else if ("28".equals(imptype))// ��������ϸ-н�����
		{
			map.put("��������", "HSMONTH");
			map.put("���ù����·�", "MONTH");
			map.put("��Ʒ��", "CPX");
			map.put("һ������", "ORG1");
			map.put("��������", "ORG2");
			map.put("��������", "ORG3");
			map.put("�ļ�����", "ORG4");
			map.put("�̶�����", "MONEYBASE");
			map.put("��Ч����", "MONEYKPI");
			map.put("���÷�", "MONEYTRAVEL");
			map.put("ҵ���д���", "MONEYSERVE2");
			map.put("��Ŀ����", "MONEYPROJECT");
			map.put("ս�Բ���", "MONEYFILLPOST");
			map.put("����", "MONEYKK");
			map.put("����ҵ��", "BUSSTYPE");
		} else if ("29".equals(imptype))//���õ���
		{
			map.put("���ù����·�", "MONTH");
			map.put("��Ա����", "PSNNAME");
			map.put("��Ա����", "PSNCODE");
			map.put("��Ʒ����", "CPXX");
			map.put("һ������", "ORG1");
			map.put("��������", "ORG2");
			map.put("��������", "ORG3");
			map.put("�ļ�����", "ORG4");
			map.put("���", "Price");
			map.put("����һ����Ŀ", "KM1");
			map.put("���������Ŀ", "KM2");
			map.put("����������Ŀ", "KM3");
			map.put("�����ļ���Ŀ", "KM4");
		}else if("30".equals(imptype))
		{
			map.put("��֯", "ZZNAME");
			map.put("Ŀ���·�","MONTH");
			map.put("Ŀ������", "TYPE");
			map.put("Ŀ��", "TARGET");
			map.put("���۷���", "MEASURE");
			map.put("ҵ������", "PRODLINENAME");
		}
		 else if ("31".equals(imptype))//�̶�
			{
				map.put("���ù����·�", "MONTH");
				map.put("��Ʒ��", "CPX");
				map.put("һ������", "ORG1");
				map.put("��������", "ORG2");
				map.put("��������", "ORG3");
				map.put("�ļ�����", "ORG4");
				map.put("���", "PRICE");
				map.put("����һ����Ŀ", "KM1");
				map.put("���������Ŀ", "KM2");
				map.put("����������Ŀ", "KM3");
				map.put("�����ļ���Ŀ", "KM4");
				map.put("ƾ֤��", "PZH");
			}
		 else if("32".equals(imptype)) {
			 map.put("ʡ", "PROVICE");
			 map.put("��", "CITY");
			 map.put("����", "COUNTY");
			 map.put("��ֳ��������", "NAME");
			 map.put("����", "TYPE");
			 map.put("��ϸ��ַ", "ADDRESS");
			 map.put("��ϵ��", "LINKMAN");
			 map.put("�ֻ�", "PHONE");
			 map.put("�ܴ���", "TOTALCL");
			 map.put("�ܷ�ĸ�����", "MCCL");
			 map.put("����", "CL");
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
								if (cell == null&&!"30".equals(request
										.getAttribute("imptype")
										.toString())) {
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
									/*
									 * �ͻ���Ӧ��ϵ���������ֶ�Ϊ�� by wensx 2016-07-20
									 */
									if (cellVal == null
											&& "2".equals(request.getAttribute(
													"imptype").toString())
											&& (j >= 6 && j!= 30))
										cellVal = "";
									if (cellVal == null
											&& "15".equals(request
													.getAttribute("imptype")
													.toString()) && j == 9)
										cellVal = "";
									if (cellVal == null
											&& "16".equals(request
													.getAttribute("imptype")
													.toString()) && j == 6)
										cellVal = "";
									if (cellVal == null
											&& "30".equals(request
													.getAttribute("imptype")
													.toString()) )
										cellVal = "";
									if (cellVal == null
											&& "2".equals(request.getAttribute(
													"imptype").toString()))
										throw new Exception("��" + i + "��" + j
												+ "��Ϊ�գ�����д");
								}
						/*	System.out.print(cellVal);
								System.out.println(j + "��");
							System.out.println(i + "��");*/

//								System.out.println(map.get(headRow.getCell(j)
//										.getStringCellValue().toString().trim()
//										.replaceAll("\r|\n", "")));
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
			if ("3".equals(imptype)) {
				imp_app_jr_kde(list, request, response);
			} else if ("4".equals(imptype)) {
				imp_APP_JR_FX(list, request, response);
			} else if ("5".equals(imptype)) {
				imp_APP_JR_ZX(list, request, response);
			} else if ("6".equals(imptype)) {
				imp_APP_JR_TXXLMB(list, request, response);
			} else if ("7".equals(imptype)) {
				imp_APP_JR_TXQK(list, request, response);
			} else if ("8".equals(imptype)) {
				imp_APP_JR_HB(list, request, response);
			} else if ("9".equals(imptype)) {
				imp_APP_JR_YJ(list, request, response);
			} else if ("10".equals(imptype)) {
				imp_APP_JR_GDLDorJRZY(list, request, response);
			} else if ("11".equals(imptype)) {
				imp_APP_JR_YQXX(list, request, response);
			} else if ("12".equals(imptype)) {
				imp_APP_JR_FKXX(list, request, response);
			} else if ("13".equals(imptype)) {
				imp_Customertransfer(list, request, response);
			} else if ("14".equals(imptype)) {
				imp_SuperiorMerchants(list, request, response);
			} else if ("15".equals(imptype)) {
				imp_PersonAndOrg(list, request, response);
			} else if ("16".equals(imptype) || "161".equals(imptype)) {
				//�����籣��������-н���������ƣ�
				imp_DetailsCharges(list, request, response,imptype);
			} else if ("17".equals(imptype) || "28".equals(imptype)) {
				imp_Salarydetails(list, request, response, imptype);
			} else if ("18".equals(imptype)) {
				imp_DetailsChargesCust(list, request, response);
			} else if ("19".equals(imptype)) {
				imp_detailschargekfjl(list, request, response);
			} else if ("20".equals(imptype)) {
				imp_detailschargekhjl(list, request, response);
			} else if ("21".equals(imptype)) {
				imp_detailschargeqtjl(list, request, response);
			} else if ("22".equals(imptype)) {
				imp_TotalCost(list, request, response);
			} else if ("23".equals(imptype)) {
				imp_StaffTraining(list, request, response);
			} else if ("24".equals(imptype)) {
				imp_specialexpenses(list, request, response);
			} else if ("25".equals(imptype)) {
				imp_Testsubsidy(list, request, response);
			} else if ("26".equals(imptype)) {
				imp_DetailsChargestz(list, request, response);
			} else if ("27".equals(imptype)) {
				imp_sbt_deptcheck(list, request, response);
			}else if ("29".equals(imptype)) {
				imp_sbt_costimport(list, request, response);
			}else if("30".equals(imptype)) {
				imp_saletarget(list,request,response);
			}else if("31".equals(imptype)) {
				imp_sbt_guding(list,request,response);
			}else if("32".equals(imptype)) {
				imp_sbt_chda(list,request,response);
			}
			else {
				for (int f = 0; f < list.size(); f++) {
					if ("1".equals(imptype)) {
						imp_scfwwh(list.get(f), request, response);
					} else if ("2".equals(imptype)) {

						imp_CUSTMOERRELATSALESMAN(list.get(f), request,
								response, sBuffer, f + 2);
						System.out.println(f);
					}
				}
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
	
	
	private static void imp_sbt_chda(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		 PreparedStatement ps1 = null;
		int result2 = 0;
		ResultSet rSet = null;
		/*for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// ��ɾ��
				sql = "DELETE wb_erp.sbt_deptcheck WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);	
		}*/
		

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
	


			// System.out.println(f);
			sql = "INSERT INTO wb_erp.app_main_archives "
					+ "(ID, Provice,City, County, Name, Type, Address, Linkman, Phone, Totalcl, MCCL,CL) "
					+ "VALUES " + "(?, ?,?, ?, ?, ?, ?, ?, ?, ?, ?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("PROVICE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("CITY"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("COUNTY"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("NAME"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ADDRESS"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("LINKMAN"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("PHONE"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("TOTALCL"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MCCL"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("CL"));
			ps.execute();
			// �ύ����

			// �ر���Դ
	
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 27����У��
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_deptcheck(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		 PreparedStatement ps1 = null;
		int result2 = 0;
		ResultSet rSet = null;
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// ��ɾ��
				sql = "DELETE wb_erp.sbt_deptcheck WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);	
		}
		

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
	


			// System.out.println(f);
			sql = "insert into wb_erp.sbt_deptcheck (MONTH,  CPX,  ORG1,  ORG2,  ORG3,	ORG4,		TS,	COPERATOR,ID,TYPE) "
					+ "values (?,	?,	?,	?,	?,	?,	to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(8, PK_ID);
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("TYPE"));
			ps.execute();
			// �ύ����

			// �ر���Դ
	
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 26��������
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_DetailsChargestz(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		PreparedStatement ps1 = null;
		int result2 = 0;
		ResultSet rSet = null;
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// ��ɾ��
				sql = "DELETE wb_erp.SBT_DetailsChargestz WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("��" + f + "��" + vo.get("ORG4").toString()
						+ "���Ų����ڣ���������µ��룡ע����0�����һ��");

			}

		
			// System.out.println(f);
			sql = "insert into wb_erp.SBT_DetailsChargestz (MONTH,  CPX,  ORG1,  ORG2,  ORG3,	ORG4,	KM1TZ,	KM2TZ,	KM1,	KM2,	KM3,	KM4,	MONEY,	DEPT,	TS,	COPERATOR,ID,BUSSTYPE,JF,DF,MEMO,DBILLDATE) "
					+ "values (?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("KM1TZ"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("KM2TZ"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("DEPT"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(16, PK_ID);
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("BUSSTYPE"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("JF"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("DF"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("MEMO"));
			DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("DBILLDATE"));

			ps.execute();
			// �ύ����

			// �ر���Դ
			 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 25���鲹��
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_Testsubsidy(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		 PreparedStatement ps1 = null;
		ResultSet rSet = null;
		int result2 = 0;
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// ��ɾ��
				sql = "DELETE wb_erp.sbt_detailsTestsubsidy WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("��" + f + "��"+ "���Ų����ڣ���������µ��룡");
			}
	

			// System.out.println(f);
			sql = "insert into wb_erp.sbt_detailsTestsubsidy (MONTH,  CPX,  ORG1,  ORG2,  ORG3,	ORG4,		KM1,	KM2,	KM3,	KM4,	MONEY,	TYPE,	TS,	COPERATOR,ID,BUSSTYPE) "
					+ "values (?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(14, PK_ID);
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("BUSSTYPE"));

			ps.execute();
			// �ύ����

			// �ر���Դ
			 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 24�������-OA����
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_specialexpenses(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		 PreparedStatement ps1 = null;
		ResultSet rSet = null;
		int result2 = 0;
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// ��ɾ��
				sql = "DELETE wb_erp.sbt_detailsspecialexpenses WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);	
		}
	
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("��" + f + "��" + vo.get("ORG4").toString()
						+ "���Ų����ڣ���������µ��룡");
			}
		
			// System.out.println(f);
			sql = "insert into wb_erp.sbt_detailsspecialexpenses (MONTH,  CPX,  ORG1,  ORG2,  ORG3,	ORG4,		KM1,	KM2,	KM3,	KM4,	MONEY,	MEMO,	TS,	COPERATOR,ID,BUSSTYPE,UNITNAME,PZH) "
					+ "values (?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("MEMO"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(14, PK_ID);
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("BUSSTYPE"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("PZH"));
			ps.execute();
			// �ύ����

			// �ر���Դ
			 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 23Ա����ѵ��
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_StaffTraining(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		 PreparedStatement ps1 = null;
		ResultSet rSet = null;
		int result2 = 0;
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// ��ɾ��
				sql = "DELETE wb_erp.sbt_detailsStaffTraining WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("��" + f + "��" + vo.get("ORG4").toString()
						+ "���Ų����ڣ���������µ��룡");
			}
		

			// System.out.println(f);
			sql = "insert into wb_erp.sbt_detailsStaffTraining (MONTH,  CPX,  ORG1,  ORG2,  ORG3,	ORG4,		KM1,	KM2,	KM3,	KM4,	MONEY,	TYPE,	TS,	COPERATOR,ID,BUSSTYPE) "
					+ "values (?,	?,	?,		?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(14, PK_ID);
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("BUSSTYPE"));

			ps.execute();
			// �ύ����

			// �ر���Դ
			 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 22�����˵���ϸ
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_TotalCost(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		 PreparedStatement ps1 = null;
		ResultSet rSet = null;
		int result2 = 0;
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// ��ɾ��
				sql = "DELETE wb_erp.SBT_DETAILSCHARGEGLZ WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);	
		}
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("��" + f + "��" + vo.get("ORG4").toString()
						+ "���Ų����ڣ���������µ��룡");
			}
			DbUtil.closeStatement(ps1);
		
			sql = "insert into WB_ERP.SBT_DETAILSCHARGEGLZ (MONTH, REQUESTID, WORKFLOWID, REQUESTNAME, PSNCODE, PSNNAME, MONEY, MEMO, ISTRUE, UNITSHORTNAME, KM1, KM2, KM3, KM4, CPX, ORG1, ORG2, ORG3, ORG4, ISBILL, TS, ID,COPERATOR,BUSSTYPE) "
					+ "values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,  to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?, ?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("REQUESTID"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("WORKFLOWID"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("REQUESTNAME"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PSNCODE"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("PSNNAME"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("MEMO"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("ISTRUE"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("UNITSHORTNAME"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("ISBILL"));

			String PK_ID = SysUtil.getId();
			ps.setString(21, PK_ID);
			DbUtil.setObject(ps, 22, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 23, Types.VARCHAR, vo.opt("BUSSTYPE"));

			ps.execute();
			// �ύ����

			// �ر���Դ
			DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);
			System.out.println(f);
		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 21����Ӫ������
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_detailschargeqtjl(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		 PreparedStatement ps1 = null;
		ResultSet rSet = null;
		int result2 = 0;
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// ��ɾ��
				sql = "DELETE wb_erp.sbt_detailschargeqtjl WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);
		}
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("��" + f + "��" + vo.get("ORG4").toString()
						+ "���Ų����ڣ���������µ��룡");
			}
		

			sql = "insert into wb_erp.sbt_detailschargeqtjl (MONTH, CPX, ORG1, ORG2, ORG3, ORG4, PSNCODE, PSNNAME, SUMMONEY, MONEYGL, MONEYYHL, MONEYZCJS, MONEYOEM, MONEYQT, MONEYBF, TS,COPERATER , ID, MEMO,BUSSTYPE) "
					+ "values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,  to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'), ?,?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PSNCODE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PSNNAME"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("SUMMONEY"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("MONEYGL"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MONEYYHL"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("MONEYZCJS"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("MONEYOEM"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("MONEYQT"));

			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("MONEYBF"));
			// DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("TS"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(17, PK_ID);
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("MEMO"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("BUSSTYPE"));
			ps.execute();
			// �ύ����
			/*System.out.println( vo.opt("PSNCODE").toString());
			if(vo.opt("PSNCODE").toString().equals("008422"))
			{
				System.out.println(vo.opt("PSNCODE").toString());
			}*/
			// �ر���Դ
		 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 20�ͻ����ܽ���
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_detailschargekhjl(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		 PreparedStatement ps1 = null;
		ResultSet rSet = null;
		int result2 = 0;
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// ��ɾ��
				sql = "DELETE wb_erp.SBT_DETAILSCHARGEKHZG WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);
		}
	
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=?";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("��" + f + "��" + vo.get("ORG4").toString()
						+ "���Ų����ڣ���������µ��룡");
			}
		

			sql = "insert into wb_erp.SBT_DETAILSCHARGEKHZG (MONTH, CPX, ORG1, ORG2, ORG3, ORG4,  PSNCODE, PSNNAME,SALESTR, TS, COPERATOR, ID, MEMO, TYPE, MONEY,BUSSTYPE) "
					+ "values (?, ?, ?, ?, ?, ?, ?, ?, ?, to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'), ?, ?, ?, ?, ?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PSNCODE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PSNNAME"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("SALESTR"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(11, PK_ID);
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("MEMO"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("BUSSTYPE"));
			ps.execute();
			// �ύ����

			// �ر���Դ
			 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 19��������
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_detailschargekfjl(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		 PreparedStatement ps1 = null;
		ResultSet rSet = null;
		int result2 = 0;
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// ��ɾ��
				sql = "DELETE wb_erp.SBT_DETAILSCHARGEkfjl WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);
				
		}
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("��" + f + "��" + vo.get("ORG4").toString()
						+ "���Ų����ڣ���������µ��룡");
			}
		
			sql = "insert into wb_erp.SBT_DETAILSCHARGEkfjl (MONTH, CPX, ORG1, ORG2, ORG3, ORG4, CUSTCODE, CUSTNAME, KFMONTH, MONEY, MONEYKH, TS, COPERATOR, ID, MEMO,XSMONTH,BUSSTYPE) "
					+ "values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'), ?, ?, ?, ?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KFMONTH"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MONEYKH"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(13, PK_ID);
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("MEMO"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("XSMONTH"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("BUSSTYPE"));

			ps.execute();
			// �ύ����

			// �ر���Դ
			 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 18ҵ�����ϸ-�ͻ�
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_DetailsChargesCust(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		 PreparedStatement ps1 = null;
		ResultSet rSet = null;
		int result2 = 0;
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// ��ɾ��
				sql = "DELETE wb_erp.sbt_detailschargesCust WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);	
		}
	
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;

			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("��" + f + "��" + vo.get("ORG4").toString()
						+ "���Ų����ڣ���������µ��룡");
			}

		

			sql = "insert into wb_erp.sbt_detailschargesCust (MONTH, CPX, ORG1, ORG2, ORG3, ORG4, CUSTCODE, CUSTNAME, KM1, KM2, KM3, KM4, MONEY, TYPE, TS, COPERATOR, ID, MEMO, UNITSHORTNAME, INVNAME, NNUMBER, SYSTYPE, SYSMEMO, HDMB,BUSSTYPE)"
					+ "values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'), ?, ?, ?, ?, ?, ?, ?, ?, ?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(16, PK_ID);
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("MEMO"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("UNITSHORTNAME"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("NNUMBER"));
			DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("SYSTYPE"));
			DbUtil.setObject(ps, 22, Types.VARCHAR, vo.opt("SYSMEMO"));
			DbUtil.setObject(ps, 23, Types.VARCHAR, vo.opt("HDBM"));
			DbUtil.setObject(ps, 24, Types.VARCHAR, vo.opt("BUSSTYPE"));

			ps.execute();
			// �ύ����

			// �ر���Դ
			 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 17н����ϸ
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_Salarydetails(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		 PreparedStatement ps1 = null;
		ResultSet rSet = null;
		int result2 = 0;

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			if ( imptype.equals("17")) {
				// ��ɾ��
				sql = "DELETE wb_erp.sbt_Salarydetails WHERE month = ?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
			}
			if (imptype.equals("28")) {
				// ��ɾ��
				sql = "DELETE wb_erp.sbt_Salarydetails_total WHERE month = ?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
			}
			conn.commit();
			DbUtil.closeStatement(ps1);
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("��" + f + "��" + vo.get("ORG4").toString()
						+ "���Ų����ڣ���������µ��룡");
			}
			
			DbUtil.closeStatement(ps1);

			if ("28".equals(imptype)) {
				sql = "insert into wb_erp.sbt_Salarydetails_total (HSMONTH,MONTH,	CPX,	ORG1,	ORG2,	ORG3,	ORG4,	MONEYBASE,	MONEYKPI,	MONEYTRAVEL,	MONEYSERVE2,	MONEYPROJECT,	MONEYFILLPOST,TS,COPERATOR,ID,MONEYKK,BUSSTYPE)"
						+ "values (?,	?,	?,	?,	?,	?,	?,	?,		?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?,?)";
				ps = conn.prepareStatement(sql);
				DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("HSMONTH"));
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("MONTH"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("CPX"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG1"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG2"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG3"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ORG4"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("MONEYBASE"));
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("MONEYKPI"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("MONEYTRAVEL"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MONEYSERVE2"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("MONEYPROJECT"));
				DbUtil
						.setObject(ps, 13, Types.VARCHAR, vo
								.opt("MONEYFILLPOST"));
				DbUtil.setObject(ps, 14, Types.VARCHAR, request
						.getAttribute("sys.userName"));
				String PK_ID = SysUtil.getId();
				ps.setString(15, PK_ID);
				DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("MONEYKK"));
				DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("BUSSTYPE"));

				ps.execute();
			}

			if ("17".equals(imptype)) {
				sql = "insert into wb_erp.sbt_Salarydetails (MONTH,	CPX,	ORG1,	ORG2,	ORG3,	ORG4,	PSNCODE,	PSNNAME,	MONEYBASE,	MONEYKPI,	MONEYTRAVEL,	MONEYSERVE1,	MONEYSERVE2,	MONEYPROJECT,	MONEYFILLPOST,TS,COPERATOR,ID,MONEYKK,BUSSTYPE)"
						+ "values (?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?,?)";
				ps = conn.prepareStatement(sql);
				DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PSNCODE"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PSNNAME"));
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("MONEYBASE"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("MONEYKPI"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MONEYTRAVEL"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("MONEYSERVE1"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("MONEYSERVE2"));
				DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("MONEYPROJECT"));
				DbUtil
						.setObject(ps, 15, Types.VARCHAR, vo
								.opt("MONEYFILLPOST"));
				DbUtil.setObject(ps, 16, Types.VARCHAR, request
						.getAttribute("sys.userName"));
				String PK_ID = SysUtil.getId();
				ps.setString(17, PK_ID);
				DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("MONEYKK"));
				DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("BUSSTYPE"));

				ps.execute();
			}
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));

			// �ύ����
			System.out.println(f);
			// �ر���Դ
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 16�籣����
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_DetailsCharges(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		PreparedStatement ps1 = null;
		// PreparedStatement ps1 = null;
		int result2 = 0;
		ResultSet rSet = null;
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
		if ("16".equals(imptype)) {
			sql = "DELETE wb_erp.SBT_DetailsCharges WHERE month in "+Times;
		}
		if ("161".equals(imptype)) {
			sql = "DELETE wb_erp.SBT_Salarydetails_other WHERE month in "+Times;;
		}
		ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		conn.commit();
		DbUtil.closeStatement(ps1);
	

		for ( int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("��" + f + "��" + vo.get("ORG4").toString()
						+ "���Ų����ڣ���������µ��룡ע����0�����һ��");

			}

			// System.out.println(f);
			if ("16".equals(imptype)) {
				sql = "insert into wb_erp.SBT_DetailsCharges (MONTH,  CPX,  ORG1,  ORG2,  ORG3,	ORG4,	PSNCODE,	PSNNAME,	KM1,	KM2,	KM3,	KM4,	MONEY,	TYPE,	TS,	COPERATOR,ID,BUSSTYPE) "
						+ "values (?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?)";
			}
			if ("161".equals(imptype)) {
				sql = "insert into wb_erp.SBT_Salarydetails_other (MONTH,  CPX,  ORG1,  ORG2,  ORG3,	ORG4,	PSNCODE,	PSNNAME,	KM1,	KM2,	KM3,	KM4,	MONEY,	TYPE,	TS,	COPERATOR,ID,BUSSTYPE) "
						+ "values (?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?)";
			}
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PSNCODE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PSNNAME"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(16, PK_ID);
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("BUSSTYPE"));
			System.out.println(f);
			ps.execute();
			// �ύ����
//System.out.println(f);
			// �ر���Դ
		  DbUtil.closeStatement(ps1);
			conn.commit();
			DbUtil.closeStatement(ps);

		}

		DbUtil.closeConnection(conn);

	}

	/**
	 * 15��Ա��֯��ϵ
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_PersonAndOrg(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		PreparedStatement ps1 = null;
		ResultSet rSet = null;
		int result2 = 0;


		for ( int f = 0; f < voList.size(); f++) {
			//
			JSONObject vo = voList.get(f);
			sql = "DELETE wb_erp.SBT_PERSONANDORG WHERE month = ? ";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
			ps1.executeUpdate();
			conn.commit();
			DbUtil.closeStatement(ps1);
		}
		

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("��" + f + "��" + vo.get("ORG4").toString()
						+ "���Ų����ڣ���������µ��룡ע����0�����һ��");

			}

		
			sql = "insert into wb_erp.SBT_PERSONANDORG (MONTH, PSNCODE, PSNNAME,  DEPTNAME, POSTNAME, OUTTIME, CPX, ORG1, ORG2, ORG3, ORG4, STATUS,TS,COPERATOR,id,BUSSTYPE)"
					+ "values (?,  ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			// DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SYS"));
			// DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORGNAME"));
			// DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORGSHORTNAME"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("PSNCODE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("PSNNAME"));
			// DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ONJOBTIME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("DEPTNAME"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("POSTNAME"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("OUTTIME"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("STATUS"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(14, PK_ID);
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("BUSSTYPE"));

			ps.execute();
			// �ύ����

			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		// �ر���Դ
		DbUtil.closeConnection(conn);

	}

	/**
	 * 14�ͻ��ҿ���ϵ����
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_SuperiorMerchants(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result = 0;
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		PreparedStatement ps1 = null;
		ResultSet rSet = null;

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			int result2 = 0;
			// ����Ƿ���ڿͻ���Ϣ
			sql = "select count(1) as CT from wb_erp.bd_cubasdoc b where  b.custcode in (?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("CUSTCODE"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("FACTCUSTCODE"));

			rSet = ps.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result2 < 2) {
				throw new Exception("��" + f
						+ "�пͻ������ʵ�����������̱��벻���ڻ�ʵ�����������̲�����ͻ�������ͬ��������");
			}

			// У���Ƿ���ڴ˿ͻ��������������£���������������

			sql = "select count(1) as CT from wb_erp.SuperiorMerchants where custcode=?";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("CUSTCODE"));
			rSet = ps.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result2 > 0) {

				sql = "update wb_erp.SuperiorMerchants "
						+ "  set CUSTNAME = ?, FACTCUSTCODE = ?, FACTCUSTNAME = ?, TYPE = ?,operator=?,ts=to_char(sysdate,'yyyy-mm-dd hh24:mi:ss')"
						+ " where custcode = ?";

				ps = conn.prepareStatement(sql);
				DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("CUSTNAME"));
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("FACTCUSTCODE"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.get("FACTCUSTNAME"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.get("TYPE"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, request
						.getAttribute("sys.userName"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.get("CUSTCODE"));
				result = ps.executeUpdate();
				// �ύ����
				conn.commit();

			}

			// ����
			else {
				sql = "insert into wb_erp.SuperiorMerchants (ID, CUSTCODE, CUSTNAME, FACTCUSTCODE, FACTCUSTNAME, MEMO, TYPE, OPERATOR, TS)"
						+ " values (?, ?, ?, ?, ?, '��������', ?, ?, to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'))";

				ps = conn.prepareStatement(sql);
				PK_ID = SysUtil.getId();
				ps.setString(1, PK_ID);
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CUSTCODE"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("CUSTNAME"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("FACTCUSTCODE"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("FACTCUSTNAME"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("TYPE"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, request
						.getAttribute("sys.userName"));

				result = ps.executeUpdate();
				// �ύ����
				conn.commit();
			}
		}

		// �ر���Դ
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);
	}

	/**
	 * 13ֱ���ͻ����ܿͻ��ƽ�
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_Customertransfer(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result;

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		PreparedStatement ps1 = null;
		ResultSet rSet = null;

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			int result2 = 0;
			// ����Ƿ���ڿͻ���Ϣ
			sql = "select count(1) as CT from wb_erp.bd_cubasdoc b where  b.custcode=?";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("CUSTCODE"));
			rSet = ps.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result2 == 0) {
				throw new Exception("��" + f + "�пͻ����벻���ڣ�������");
			}

			// ����Ƿ��ҵ��Ա��Ϣ
			sql = "select count(1) as CT from wb_erp.view_ddm_userinfo b where  b.USER_NAME IN (?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("PSNCODEBEFORE"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("PSNCODEAFTER"));
			rSet = ps.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result2 < 1) {
				throw new Exception("��" + f + 1 + "��ҵ��Ա���벻���ڣ�������");
			}

			// У���Ƿ���ڴ˿ͻ��������������£���������������

			sql = "select count(1) as CT from wb_erp.Customertransfer where custcode=?";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("CUSTCODE"));
			rSet = ps.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result2 > 0) {

				sql = "update wb_erp.Customertransfer a "
						+ " set DQ       = ?,"
						+ "   YXB      = ?,"
						+ "   CUSTNAME = ?,"
						+ "   TRANSFERTYPE   = ?,"
						+ "   MONTH          = ?,"
						+ "   SALESTRUBEFORE = ?,"
						+ "   PSNCODEBEFORE  = ?,"
						+ "   PSNNAMEBEFORE  = ?,"
						+ "   SALESTRUAFTER  = ?,"
						+ "   PSNCODEAFTER   = ?,"
						+ "   PSNNAMEAFTER   = ?,"
						+ "   COPERATOR      =?,"
						+ "   TS             = to_char(SYSDATE, 'YYYY-MM-DD HH24:mi:ss')"
						+ " where custcode=?";

				ps = conn.prepareStatement(sql);
				DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("DQ"));
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("YXB"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.get("CUSTNAME"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.get("TRANSFERTYPE"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.get("MONTH"));
				DbUtil
						.setObject(ps, 6, Types.VARCHAR, vo
								.get("SALESTRUBEFORE"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.get("PSNCODEBEFORE"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.get("PSNNAMEBEFORE"));
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo.get("SALESTRUAFTER"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.get("PSNCODEAFTER"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.get("PSNNAMEAFTER"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, request
						.getAttribute("sys.userName"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.get("CUSTCODE"));
				result = ps.executeUpdate();
				// �ύ����
				conn.commit();

			}

			// ����
			else {
				sql = "insert into wb_erp.Customertransfer "
						+ "(ID,DQ,YXB,CUSTCODE, CUSTNAME,  TRANSFERTYPE, MONTH, SALESTRUBEFORE, PSNCODEBEFORE, PSNNAMEBEFORE, SALESTRUAFTER, PSNCODEAFTER, PSNNAMEAFTER, COPERATOR, TS) "
						+ "values (?,?,?, ?,?, ?, ?, ?, ?,?, ?, ?, ?,?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'))";

				ps = conn.prepareStatement(sql);
				PK_ID = SysUtil.getId();
				ps.setString(1, PK_ID);
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("DQ"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("YXB"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("CUSTCODE"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CUSTNAME"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("TRANSFERTYPE"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("MONTH"));
				DbUtil
						.setObject(ps, 8, Types.VARCHAR, vo
								.opt("SALESTRUBEFORE"));
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("PSNCODEBEFORE"));
				DbUtil
						.setObject(ps, 10, Types.VARCHAR, vo
								.opt("PSNNAMEBEFORE"));
				DbUtil
						.setObject(ps, 11, Types.VARCHAR, vo
								.opt("SALESTRUAFTER"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("PSNCODEAFTER"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("PSNNAMEAFTER"));
				DbUtil.setObject(ps, 14, Types.VARCHAR, request
						.getAttribute("sys.userName"));

				result = ps.executeUpdate();
				// �ύ����
				conn.commit();
			}
		}

		// �ر���Դ
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * �����г���Χά�� yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_scfwwh(JSONObject vo, HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		String PK_ID = null;
		String sql = "";
		int result = 0;
		if (null != vo) {
			Connection conn = DbUtil.getConnection();
			DbUtil.startTrans(conn, "");
			PreparedStatement ps = null;
			// PreparedStatement ps1 = null;

			// ��ɾ��
			// sql= "DELETE wb_erp.APP_WB_MARKET_ZX WHERE USERCODE = ?";
			// ps1 = conn.prepareStatement(sql);
			// DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("USERCODE"));
			// result=ps1.executeUpdate();

			// ����
			sql = "INSERT INTO wb_erp.APP_WB_MARKET_ZX "
					+ "(PK_ID, USERCODE, USERNAME, PROVINCE, CITY, AREA, PRODUCTLINE, SXRQ, SFSX) "
					+ "VALUES " + " (?, ?, ?, ?, ?, ?, ?, ?, ?)";
			ps = conn.prepareStatement(sql);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("USERCODE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("USERNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("PROVINCE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CITY"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("AREA"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PRODUCTLINE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("SXRQ"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("SFSX"));
			result = ps.executeUpdate();

			// �ύ����
			conn.commit();

			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);
			DbUtil.closeConnection(conn);
		}

	}

	/**
	 * ����ҵ��Ա��ͻ���Ӧ��ϵά��
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_CUSTMOERRELATSALESMAN(JSONObject vo,
			HttpServletRequest request, HttpServletResponse response,
			StringBuffer sBuffer, int row) throws Exception {
		String PK_ID = null;
		String sql = "";
		int result = 0;
		if (null != vo) {
			Connection conn = DbUtil.getConnection();
			DbUtil.startTrans(conn, "");
			PreparedStatement ps = null;
			PreparedStatement ps1 = null;
			ResultSet rSet = null;
			// ��ѯ�Ƿ�ͬһ����
			
		//�޸ķſ�Ȩ�޽��е�������
		/*	sql = "select 1 as CT "
					+ "from dual where (select distinct zsj.zzname "
					+ "       from wb_erp.zsj_ddm_userinfo zsj "
					+ "      where zsj.user_name in (?,?)) in "
					+ "    (select zu.deptname "
					+ "       from wb_erp.APP_ZBXX_USER zu "
					+ "      where zu.MANGERCODE = ?) or '����' in (select zu.deptname from wb_erp.APP_ZBXX_USER zu where zu.MANGERCODE = ?)";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("PSNCODE"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("PSNCODE_CUSTMANAGER"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result != 1) {
				sBuffer.append("��" + row + "��" + vo.get("PSNNAME") + "("
						+ vo.get("PSNCODE") + "),");
			}
			*/

			// У���Ƿ���ڲ���Ƭ��
			result = 0;
			sql = "select count(1) as CT from ( "
					+ "select 1 from wb_erp.org_dept b where name like ? and rownum=1 "
					+ " union all  "
					+ "select 1 from wb_erp.org_dept b where name like ? and rownum=1 "
					+ "union all  "
					+ "select 1 from wb_erp.org_dept b where name like ? and rownum=1)";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("YXB"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("XQ"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("SALESTR"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result != 3) {
				System.out.println(sql);
				throw new Exception("��" + row + "��Ӫ����/С��/Ƭ����ϵͳƥ�䣡�����µ���");
			}

			// У���Ʒ���ֶ�
			result = 0;
			sql = "select 1 AS CT " + "  from nc.bd_prodline@sbtnc a "
					+ "  where isseal = 'N'  " + "and dr = 0 "
					+ "and prodlinename = ?  ";

			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("RELATIONTYPE"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result != 1) {
				throw new Exception("��" + row + "�в�Ʒ�ߴ���������");
			}
			/*
			// У��ͻ�����ʱ��
			if(!vo.get("MONTH1").equals("")){
			result = 0;
			sql = "SELECT 1 AS CT FROM DUAL  WHERE TO_CHAR(SYSDATE,'YYYY-MM')=?";

			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH1"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result != 1) {
				throw new Exception("��" + row + "���Ͽͻ������·ݱ���Ϊ���£�������");
			}
			}
			*/
			
			// У��ͻ����� �뿪���鳤+������Ա��������һ������
			if(vo.get("PSNCODE_CUSTMANAGER").toString().equals(""))
			{
				if(vo.get("MANAGERCODE").toString().equals("")&&vo.get("PSNCODE").toString().equals("") &&vo.get("MEMO1").toString().equals(""))
				{
					throw new Exception("��" + row + "�У��ͻ������򣨿����鳤+����ҵ��Ա������һ��Ϊ����");
				}
				else
				{
					// У���Ƿ���ڲ���Ƭ��
					Double d = 0.0;
					Double d1 = 0.0;
					Double d2 = 0.0;
					Double d3 = 0.0;
					if(vo.get("MEMO1")==null)
					{
						throw new Exception("��" + row + "���������������Ϊ��");
					}
					d1 = Double.parseDouble(vo.get("MEMO1").toString());
					if (vo.get("MEMO2") == null
							|| vo.get("MEMO2").toString().equals(""))
						d2 = 0.0;
					else
						d2 = Double.parseDouble(vo.get("MEMO2").toString());

					if (vo.get("MEMO3") == null
							|| vo.get("MEMO3").toString().equals(""))
						d3 = 0.0;
					else
						d3 = Double.parseDouble(vo.get("MEMO3").toString());

					d = d1 + d2 + d3;
					if (d != 1) {
						throw new Exception("��" + row + "�н������������ӱ���Ϊ1");
					}
				}
			}
				
			
			// ��ѯ�Ƿ����
			result = 0;
			sql = "select count(*) CT from wb_erp.zsj_so_custmoerrelatsalesman a where a.CUSTCODE= ? and a.RELATIONTYPE= (select a.pk_prodline from nc.bd_prodline@sbtnc a where  isseal='N' and dr=0 and prodlinename= ?) and a.TYPE= ? ";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("CUSTCODE"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("RELATIONTYPE"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("TYPE"));
			rSet = ps1.executeQuery();
			//System.out.println(vo.get("CUSTCODE"));
			if (rSet.next()) {
				result = rSet.getInt("CT");
			}
			if (result > 0) {
				// �޸�
				sql = "update wb_erp.zsj_so_custmoerrelatsalesman "
						+ " set PSNCODE         = ?, "
						+ "  PSNNAME         = ?, "
						+ "   JOBNAME     = (select du.postname "
						+ "   from wb_erp.zsj_ddm_userinfo du "
						+ "    where du.user_name = ?), "
						+ "  VSALESTRUNAME_3 = (select du.zzname "
						+ "      from wb_erp.zsj_ddm_userinfo du "
						+ "     where du.user_name = ?), "
						+ "       VSALESTRUNAME   = (select du.bmname "
						+ "      from wb_erp.zsj_ddm_userinfo du "
						+ "                           where du.user_name = ?), "
						+ "       COPERATORID     = ?, "
						+ "       FSTATUS         = 1, "
						+ "       TS              = to_char(SYSDATE, 'YYYY-MM-DD HH24:mi:ss'), "
						+ "       PSNCODEMANA_TECHNICAL   = ?, "
						+ "       MEMO1           = ?, "
						+ "       PSNCODE2        = ?, "
						+ "       PSNNAME2        = ?, "
						+ "       MEMO2           = ?, "
						+ "       PSNCODE3        = ?, "
						+ "       PSNNAME3        = ?, "
						+ "       MEMO3           = ?, "
						+ "       MANAGERCODE     = ?, "
						+ "       MANAGERNAME     = ?, "
						+ "       MEMO            = ?, "
						+ "       YXB             = ? ,"
						+ "       salestr         = ? ," 
						+ "       XQ = ? ,"
						+ "       PSNNAMEMANA_TECHNICAL = ? ," 
						+ "       MONTH2 = ? ,"
						+ "       MPIG = nvl(?,MPIG) ," 
						+ "      RPIG = nvl(?,RPIG),  "
						
						+ "       PSNCODE_CUSTMANAGER = ? ,"
						+ "       PSNNAME_CUSTMANAGER = ? ," 
						+ "       PSNCODE1_TECHNICAL = ? ,"
						+ "       PSNNAME1_TECHNICAL = ? ," 
						+ "       TECHNICAL1_RATE = ? ,"
						+ "       PSNCODE2_TECHNICAL = ? ,"
						+ "       PSNNAME2_TECHNICAL = ? ," 
						+ "       TECHNICAL2_RATE = ? "
						
						+ " where CUSTCODE = ? "
						+ " and RELATIONTYPE = (select a.pk_prodline "
						+ "  from nc.bd_prodline@sbtnc a "
						+ "  where isseal = 'N' " + " and dr = 0 "
						+ "  and prodlinename = ?) " + "   and TYPE = 'Y' ";
				ps = conn.prepareStatement(sql);
				DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("PSNCODE"));
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("PSNNAME"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.get("PSNCODE"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.get("PSNCODE"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.get("PSNCODE"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, request
						.getAttribute("sys.userName"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.get("PSNCODEMANA_TECHNICAL"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.get("MEMO1"));
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo.get("PSNCODE2"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.get("PSNNAME2"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.get("MEMO2"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.get("PSNCODE3"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.get("PSNNAME3"));
				DbUtil.setObject(ps, 14, Types.VARCHAR, vo.get("MEMO3"));
				DbUtil.setObject(ps, 15, Types.VARCHAR, vo.get("MANAGERCODE"));
				DbUtil.setObject(ps, 16, Types.VARCHAR, vo.get("MANAGERNAME"));
				DbUtil.setObject(ps, 17, Types.VARCHAR, vo.get("MEMO"));
				DbUtil.setObject(ps, 18, Types.VARCHAR, vo.get("YXB"));
				DbUtil.setObject(ps, 19, Types.VARCHAR, vo.get("SALESTR"));
				DbUtil.setObject(ps, 20, Types.VARCHAR, vo.get("XQ"));
				DbUtil.setObject(ps, 21, Types.VARCHAR, vo.get("PSNNAMEMANA_TECHNICAL"));
				DbUtil.setObject(ps, 22, Types.VARCHAR, vo.get("MONTH2"));
				DbUtil.setObject(ps, 23, Types.VARCHAR, vo.get("MPIG"));
				DbUtil.setObject(ps, 24, Types.VARCHAR, vo.get("RPIG"));
				
				DbUtil.setObject(ps, 25, Types.VARCHAR, vo.get("PSNCODE_CUSTMANAGER"));
				DbUtil.setObject(ps, 26, Types.VARCHAR, vo.get("PSNNAME_CUSTMANAGER"));
				DbUtil.setObject(ps, 27, Types.VARCHAR, vo.get("PSNCODE1_TECHNICAL"));
				DbUtil.setObject(ps, 28, Types.VARCHAR, vo.get("PSNNAME1_TECHNICAL"));
				DbUtil.setObject(ps, 29, Types.VARCHAR, vo.get("TECHNICAL1_RATE"));
				DbUtil.setObject(ps, 30, Types.VARCHAR, vo.get("PSNCODE2_TECHNICAL"));
				DbUtil.setObject(ps, 31, Types.VARCHAR, vo.get("PSNNAME2_TECHNICAL"));
				DbUtil.setObject(ps, 32, Types.VARCHAR, vo.get("TECHNICAL2_RATE"));
				
				DbUtil.setObject(ps, 33, Types.VARCHAR, vo.get("CUSTCODE"));
				DbUtil.setObject(ps, 34, Types.VARCHAR, vo.get("RELATIONTYPE"));
				result = ps.executeUpdate();

				// �޸Ķ����Ʒ�ߵĿͻ���ʵ�ʾ�����ֻ����һ��
				/*
				 * sql="update wb_erp.zsj_so_custmoerrelatsalesman a set a.managercustcode=?,a.managercustname=?,a.managertype=? where custcode=?"
				 * ; ps = conn.prepareStatement(sql); DbUtil.setObject(ps, 1,
				 * Types.VARCHAR, vo.get("MANAGERCUSTCODE"));
				 * DbUtil.setObject(ps, 2, Types.VARCHAR,
				 * vo.get("MANAGERCUSTNAME")); DbUtil.setObject(ps, 3,
				 * Types.VARCHAR, vo.get("MANAGERTYPE")); DbUtil.setObject(ps,
				 * 4, Types.VARCHAR, vo.get("CUSTCODE")); result =
				 * ps.executeUpdate();
				 */
			} else {
				// ����
				sql = "insert into wb_erp.zsj_so_custmoerrelatsalesman "
						+ "(PK_DZ,BILLCODE,DATEVERSION,CUSTCODE,CUSTNAME, PSNCODE,PSNNAME,JOBNAME,VSALESTRUNAME_3, "
						+ "VSALESTRUNAME,ISRECENTLY,COPERATORID,DBILLDATE,CAPPROVEID,DAPPROVEDATE,FSTATUS, "
						+ "DR,TS,TYPE,RELATIONTYPE,SALESTR,MEMO,YXB,MANAGERCODE,MANAGERNAME,XQ,PSNCODEMANA_TECHNICAL,MEMO2,MEMO3,MEMO1,PSNCODE2,PSNNAME2,PSNCODE3,PSNNAME3,PSNNAMEMANA_TECHNICAL,MONTH2,MPIG,RPIG,  PSNCODE_CUSTMANAGER ,PSNNAME_CUSTMANAGER,PSNCODE1_TECHNICAL, PSNNAME1_TECHNICAL, TECHNICAL1_RATE, PSNCODE2_TECHNICAL, PSNNAME2_TECHNICAL, TECHNICAL2_RATE) "
						+ "values "
						+ "(?,'',to_char(sysdate, 'YYYY-MM-DD'),?,?,?,?, "
						+ "(select du.postname from wb_erp.zsj_ddm_userinfo du where du.user_name = ?), "
						+ "(select du.zzname from wb_erp.zsj_ddm_userinfo du where du.user_name = ?), "
						+ "(select du.bmname from wb_erp.zsj_ddm_userinfo du where du.user_name = ?), "
						+ "'Y',?,to_char(SYSDATE, 'YYYY-MM-DD HH24:mi:ss'),'','',0,0,to_char(SYSDATE, 'YYYY-MM-DD HH24:mi:ss'), "
						+ "?,(select a.pk_prodline from nc.bd_prodline@sbtnc a where  isseal='N' and dr=0 and prodlinename= ?),"
						+ "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
				ps = conn.prepareStatement(sql);
				PK_ID = SysUtil.getId();
				ps.setString(1, PK_ID);
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("CUSTCODE"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("CUSTNAME"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("PSNCODE"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PSNNAME"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("PSNCODE"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PSNCODE"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PSNCODE"));
				DbUtil.setObject(ps, 9, Types.VARCHAR, request
						.getAttribute("sys.userName"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("TYPE"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("RELATIONTYPE"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("SALESTR"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("MEMO"));
				DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("YXB"));
				DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("MANAGERCODE"));
				DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("MANAGERNAME"));
				DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("XQ"));
				DbUtil
						.setObject(ps, 18, Types.VARCHAR, vo
								.opt("PSNCODEMANA_TECHNICAL"));
				DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("MEMO2"));
				DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("MEMO3"));
				DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("MEMO1"));
				DbUtil.setObject(ps, 22, Types.VARCHAR, vo.opt("PSNCODE2"));
				DbUtil.setObject(ps, 23, Types.VARCHAR, vo.opt("PSNNAME2"));
				DbUtil.setObject(ps, 24, Types.VARCHAR, vo.opt("PSNCODE3"));
				DbUtil.setObject(ps, 25, Types.VARCHAR, vo.opt("PSNNAME3"));
				DbUtil.setObject(ps, 26, Types.VARCHAR, vo.opt("PSNNAMEMANA_TECHNICAL"));
				DbUtil.setObject(ps, 27, Types.VARCHAR, vo.opt("MONTH2"));
				DbUtil.setObject(ps, 28, Types.VARCHAR, vo.opt("MPIG"));
				DbUtil.setObject(ps, 29, Types.VARCHAR, vo.opt("RPIG"));
				
				DbUtil.setObject(ps, 30, Types.VARCHAR, vo.get("PSNCODE_CUSTMANAGER"));
				DbUtil.setObject(ps, 31, Types.VARCHAR, vo.get("PSNNAME_CUSTMANAGER"));
				DbUtil.setObject(ps, 32, Types.VARCHAR, vo.get("PSNCODE1_TECHNICAL"));
				DbUtil.setObject(ps, 33, Types.VARCHAR, vo.get("PSNNAME1_TECHNICAL"));
				DbUtil.setObject(ps, 34, Types.VARCHAR, vo.get("TECHNICAL1_RATE"));
				DbUtil.setObject(ps, 35, Types.VARCHAR, vo.get("PSNCODE2_TECHNICAL"));
				DbUtil.setObject(ps, 36, Types.VARCHAR, vo.get("PSNNAME2_TECHNICAL"));
				DbUtil.setObject(ps, 37, Types.VARCHAR, vo.get("TECHNICAL2_RATE"));

				result = ps.executeUpdate();
			}
			// �ύ����
			conn.commit();

			// �ر���Դ
			DbUtil.closeResultSet(rSet);
			DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);
			DbUtil.closeConnection(conn);
		}

	}

	/**
	 * ����ɴ���
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_app_jr_kde(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// ��ɾ��
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);
		// result=ps1.executeUpdate();

		// ����
		sql = "insert into wb_erp.app_jr_kde "
				+ "(PK_ID, KHBM, YXB, KHXM, KHDH, ZCSJ, JBD2, JBD3, FXDKYE, ZXDKYE, XYKYE, FXSXED, DQKDED,TS) "
				+ "values " + "(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,sysdate)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("KHBM"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("KHXM"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("KHDH"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ZCSJ"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("JBD2"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("JBD3"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("FXDKYE"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("ZXDKYE"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("XYKYE"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("FXSXED"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("DQKDED"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}

		result = ps.executeBatch();

		// �ύ����
		conn.commit();

		// �ر���Դ
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * ���������ѯ yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_FX(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// ��ɾ��
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);
		// result=ps1.executeUpdate();

		// ����
		sql = "INSERT INTO wb_erp.APP_JR_FX "
				+ "(PK_ID, KHBS, MQYH, DQ, SQ, YXB, GCMC, KHLX, KHBM, KHXM, SZXS, JYZE, TJED, JYEQJ, DH, SFZH, HZNX, YYZZH, XCMQJG, YYSM, XCMQR, MQRQ, CLQQRQ, XDSX, WTLB, XDSHSM, TJRQ, TJHBH, HKKH, FKRQ, FKJE, FKLL, FKQS, NCSKDRQ, SFDHTZTH, WHR, WHRQ, BZ, YFKSJ, CCTS, FKHS, TQHKSJ, YHKSJ, DQYE, TS) "
				+ "VALUES "
				+ "(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, sysdate)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("KHBS"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("MQYH"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("DQ"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("SQ"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("GCMC"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("KHLX"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KHBM"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KHXM"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("SZXS"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("JYZE"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("TJED"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("JYEQJ"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("DH"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("SFZH"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("HZNX"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("YYZZH"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("XCMQJG"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("YYSM"));
			DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("XCMQR"));
			DbUtil.setObject(ps, 22, Types.VARCHAR, vo.opt("MQRQ"));
			DbUtil.setObject(ps, 23, Types.VARCHAR, vo.opt("CLQQRQ"));
			DbUtil.setObject(ps, 24, Types.VARCHAR, vo.opt("XDSX"));
			DbUtil.setObject(ps, 25, Types.VARCHAR, vo.opt("WTLB"));
			DbUtil.setObject(ps, 26, Types.VARCHAR, vo.opt("XDSHSM"));
			DbUtil.setObject(ps, 27, Types.VARCHAR, vo.opt("TJRQ"));
			DbUtil.setObject(ps, 28, Types.VARCHAR, vo.opt("TJHBH"));
			DbUtil.setObject(ps, 29, Types.VARCHAR, vo.opt("HKKH"));
			DbUtil.setObject(ps, 30, Types.VARCHAR, vo.opt("FKRQ"));
			DbUtil.setObject(ps, 31, Types.VARCHAR, vo.opt("FKJE"));
			DbUtil.setObject(ps, 32, Types.VARCHAR, vo.opt("FKLL"));
			DbUtil.setObject(ps, 33, Types.VARCHAR, vo.opt("FKQS"));
			DbUtil.setObject(ps, 34, Types.VARCHAR, vo.opt("NCSKDRQ"));
			DbUtil.setObject(ps, 35, Types.VARCHAR, vo.opt("SFDHTZTH"));
			DbUtil.setObject(ps, 36, Types.VARCHAR, vo.opt("WHR"));
			DbUtil.setObject(ps, 37, Types.VARCHAR, vo.opt("WHRQ"));
			DbUtil.setObject(ps, 38, Types.VARCHAR, vo.opt("BZ"));
			DbUtil.setObject(ps, 39, Types.VARCHAR, vo.opt("YFKSJ"));
			DbUtil.setObject(ps, 40, Types.VARCHAR, vo.opt("CCTS"));
			DbUtil.setObject(ps, 41, Types.VARCHAR, vo.opt("FKHS"));
			DbUtil.setObject(ps, 42, Types.VARCHAR, vo.opt("TQHKSJ"));
			DbUtil.setObject(ps, 43, Types.VARCHAR, vo.opt("YHKSJ"));
			DbUtil.setObject(ps, 44, Types.VARCHAR, vo.opt("DQYE"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}
		result = ps.executeBatch();

		// �ύ����
		conn.commit();

		// �ر���Դ
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * ����ֱ����ѯ yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_ZX(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// ��ɾ��
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);
		// result=ps1.executeUpdate();

		// ����
		sql = "INSERT INTO wb_erp.APP_JR_ZX "
				+ "(PK_ID, KHLX, DKYH, SSSC, DQ, JD, YXB, GCMC, KHBM, KHXM, KHDH, SZXS, SCTHRQ, SFZH, CLMZTS, YZCJ, YZNX, TJED, XTZCSJ, JYNJYE, JCMZTS, CLRZTS, DQYHFZ, DQMJFZ, XYGLDS, JSOARQ, DSHRQ, MQRY, HSMQRQ, XDSX, WTLB, XDSHSM, TJRQ, TJHBH, SFBQCL, CLQQRQ, SFTP, TPFL, TPSM, KJTS, XQMZTS, JDYJ, FKJE, FKRQ, NCXZRQ, SFJBZJ, FKQS, WLSBM, WLSMC, DH, WLSSFZH, YHFZE, JZ, BZ, YWY, YHBJRQ, SJJQBJRQ, HKKH, DKLL, XDTJYF, XTZCSJ2, JYNJYE2, TS) "
				+ "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, SYSDATE)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("KHLX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("DKYH"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("SSSC"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("DQ"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("JD"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("GCMC"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KHBM"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KHXM"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("KHDH"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("SZXS"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("SCTHRQ"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("SFZH"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("CLMZTS"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("YZCJ"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("YZNX"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("TJED"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("XTZCSJ"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("JYNJYE"));
			DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("JCMZTS"));
			DbUtil.setObject(ps, 22, Types.VARCHAR, vo.opt("CLRZTS"));
			DbUtil.setObject(ps, 23, Types.VARCHAR, vo.opt("DQYHFZ"));
			DbUtil.setObject(ps, 24, Types.VARCHAR, vo.opt("DQMJFZ"));
			DbUtil.setObject(ps, 25, Types.VARCHAR, vo.opt("XYGLDS"));
			DbUtil.setObject(ps, 26, Types.VARCHAR, vo.opt("JSOARQ"));
			DbUtil.setObject(ps, 27, Types.VARCHAR, vo.opt("DSHRQ"));
			DbUtil.setObject(ps, 28, Types.VARCHAR, vo.opt("MQRY"));
			DbUtil.setObject(ps, 29, Types.VARCHAR, vo.opt("HSMQRQ"));
			DbUtil.setObject(ps, 30, Types.VARCHAR, vo.opt("XDSX"));
			DbUtil.setObject(ps, 31, Types.VARCHAR, vo.opt("WTLB"));
			DbUtil.setObject(ps, 32, Types.VARCHAR, vo.opt("XDSHSM"));
			DbUtil.setObject(ps, 33, Types.VARCHAR, vo.opt("TJRQ"));
			DbUtil.setObject(ps, 34, Types.VARCHAR, vo.opt("TJHBH"));
			DbUtil.setObject(ps, 35, Types.VARCHAR, vo.opt("SFBQCL"));
			DbUtil.setObject(ps, 36, Types.VARCHAR, vo.opt("CLQQRQ"));
			DbUtil.setObject(ps, 37, Types.VARCHAR, vo.opt("SFTP"));
			DbUtil.setObject(ps, 38, Types.VARCHAR, vo.opt("TPFL"));
			DbUtil.setObject(ps, 39, Types.VARCHAR, vo.opt("TPSM"));
			DbUtil.setObject(ps, 40, Types.VARCHAR, vo.opt("KJTS"));
			DbUtil.setObject(ps, 41, Types.VARCHAR, vo.opt("XQMZTS"));
			DbUtil.setObject(ps, 42, Types.VARCHAR, vo.opt("JDYJ"));
			DbUtil.setObject(ps, 43, Types.VARCHAR, vo.opt("FKJE"));
			DbUtil.setObject(ps, 44, Types.VARCHAR, vo.opt("FKRQ"));
			DbUtil.setObject(ps, 45, Types.VARCHAR, vo.opt("NCXZRQ"));
			DbUtil.setObject(ps, 46, Types.VARCHAR, vo.opt("SFJBZJ"));
			DbUtil.setObject(ps, 47, Types.VARCHAR, vo.opt("FKQS"));
			DbUtil.setObject(ps, 48, Types.VARCHAR, vo.opt("WLSBM"));
			DbUtil.setObject(ps, 49, Types.VARCHAR, vo.opt("WLSMC"));
			DbUtil.setObject(ps, 50, Types.VARCHAR, vo.opt("DH"));
			DbUtil.setObject(ps, 51, Types.VARCHAR, vo.opt("WLSSFZH"));
			DbUtil.setObject(ps, 52, Types.VARCHAR, vo.opt("YHFZE"));
			DbUtil.setObject(ps, 53, Types.VARCHAR, vo.opt("JZ"));
			DbUtil.setObject(ps, 54, Types.VARCHAR, vo.opt("BZ"));
			DbUtil.setObject(ps, 55, Types.VARCHAR, vo.opt("YWY"));
			DbUtil.setObject(ps, 56, Types.VARCHAR, vo.opt("YHBJRQ"));
			DbUtil.setObject(ps, 57, Types.VARCHAR, vo.opt("SJJQBJRQ"));
			DbUtil.setObject(ps, 58, Types.VARCHAR, vo.opt("HKKH"));
			DbUtil.setObject(ps, 59, Types.VARCHAR, vo.opt("DKLL"));
			DbUtil.setObject(ps, 60, Types.VARCHAR, vo.opt("XDTJYF"));
			DbUtil.setObject(ps, 61, Types.VARCHAR, vo.opt("XTZCSJ2"));
			DbUtil.setObject(ps, 62, Types.VARCHAR, vo.opt("JYNJYE2"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}
		result = ps.executeBatch();

		// �ύ����
		conn.commit();

		// �ر���Դ
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * ������Ϣ����Ŀ���ѯ yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_TXXLMB(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// ��ɾ��
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);
		// result=ps1.executeUpdate();

		// ����
		sql = "INSERT INTO wb_erp.APP_JR_TXXLMB "
				+ "(PK_ID, BLYH, YXB, KHBM, KHXM, FKJE, XLJSQJ, XLJS, ZLMB, DQSJXL, QEBXXL, TS) "
				+ "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, sysdate)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("BLYH"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("KHBM"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("KHXM"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("FKJE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("XLJSQJ"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("XLJS"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("ZLMB"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("DQSJXL"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("QEBXXL"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}

		result = ps.executeBatch();

		// �ύ����
		conn.commit();

		// �ر���Դ
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * ������Ϣ�����ѯ yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_TXQK(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// ��ɾ��
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);
		// result=ps1.executeUpdate();

		// ����
		sql = "INSERT INTO wb_erp.APP_JR_TXQK "
				+ "(PK_ID, BXFA, BXYF, BLYH, YXB, GCMC, KHBM, KHXM, DH, SFZH, MQRQ, FKRQ, FKJE, DKLX, HKWBSJ, MQQYNXL, MQHYNXL, BXL, BXJE, BXSM, TS) "
				+ "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, sysdate)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("BXFA"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("BXYF"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("BLYH"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("GCMC"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("KHBM"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("KHXM"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("DH"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("SFZH"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MQRQ"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.get("FKRQ"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("FKJE"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("DKLX"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("HKWBSJ"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("MQQYNXL"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("MQHYNXL"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("BXL"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("BXJE"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("BXSM"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}

		result = ps.executeBatch();

		// �ύ����
		conn.commit();

		// �ر���Դ
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * ���뻹����ѯ yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_HB(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// ��ɾ��
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);
		// result=ps1.executeUpdate();

		// ����
		sql = "INSERT INTO wb_erp.APP_JR_HB "
				+ "(PK_ID, KHLX, YXB, DKYH, KHBM, KHXM, KHDH, YHSJ, YHJE, HKKH, TS) "
				+ "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, sysdate)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("KHLX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("DKYH"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("KHBM"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("KHXM"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("KHDH"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("YHSJ"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("YHJE"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("HKKH"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}

		result = ps.executeBatch();

		// �ύ����
		conn.commit();

		// �ر���Դ
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * ����Ԥ����ѯ yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_YJ(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// ��ɾ��
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);
		// result=ps1.executeUpdate();

		// ����
		sql = "INSERT INTO wb_erp.APP_JR_YJ "
				+ "(PK_ID, YXB, KHBM, KHXM, QYXL, SYXL, DYXL, DXB, YJYY, JCYJMB, YJCS, TS) "
				+ "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, sysdate)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("YXB"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("KHBM"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("KHXM"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("QYXL"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("SYXL"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("DYXL"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("DXB"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("YJYY"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("JCYJMB"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("YJCS"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}

		result = ps.executeBatch();

		// �ύ����
		conn.commit();

		// �ر���Դ
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * 12����ſ���Ϣ
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_FKXX(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		ResultSet rSet = null;

		// �޸�
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// У���Ƿ�����쳣�Ƽ���ſͻ�
			int result2 = 0;
			sql = "select count(*) as CT from wb_erp.APP_JR_ZX where  TJHBH=? and khbm=? and KHXM=?";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("TJHBH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("KHBM"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.get("KHXM"));
			rSet = ps.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result2 == 0) {
				// DbUtil.closeStatement(ps1);
				DbUtil.closeStatement(ps);
				DbUtil.closeConnection(conn);
				throw new Exception("�Ƽ���ţ�" + vo.get("TJHBH") + "��ͻ���ƥ�䣡�����µ���");
			}

			sql = "update wb_erp.APP_JR_ZX set  FKRQ=?,FKJE=?,HKKH=?,NCXZRQ=? where TJHBH=?";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("FKRQ"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("FKJE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.get("HKKH"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.get("NCXZRQ"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.get("TJHBH"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}
		result = ps.executeBatch();

		// �ύ����
		conn.commit();

		// �ر���Դ
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * �������������Ϣ boyang
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_YQXX(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// ��ɾ��
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);1426
		// result=ps1.executeUpdate();

		// ����
		sql = "INSERT INTO wb_erp.APP_JR_YQXX "
				+ "(PK_YQXX, KHLX, KHBM, ZQ, YXB, KHXM, KHDH, SZXS, BLYH, DKED, HKKH, YQLB, YQKSSJ, YQTS,YQJE,GJQK,QYHFQK,SFZH,JQRQ,ZHJE,YQYE) "
				+ "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?,?,?,?,?,?,?)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("KHLX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("KHBM"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ZQ"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("KHXM"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("KHDH"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("SZXS"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("BLYH"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("DKED"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("HKKH"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("YQLB"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("YQKSSJ"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("YQTS"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("YQJE"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("GJQK"));

			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("QYHFQK"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("SFZH"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("JQRQ"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("ZHJE"));
			DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("YQYE"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}

		result = ps.executeBatch();

		// �ύ����
		conn.commit();

		// �ر���Դ
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * �����Ӧ�ҵ��쵼������רԱ yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_GDLDorJRZY(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// ��ɾ��
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);
		// result=ps1.executeUpdate();

		// ����
		sql = "INSERT INTO wb_erp.APP_JR_AUTH "
				+ "(PK_ID, DQ, JD, YXB, ZXSC, JRQY, JRXFSC, USERCODE, USERNAME, GDLDBM, GDLDXM, JRZYBM, JRZYXM, ISALL,TS) "
				+ "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0,sysdate)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("DQ"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("JD"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ZXSC"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("JRQY"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("JRXFSC"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("USERCODE"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("USERNAME"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("GDLDBM"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("GDLDXM"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("JRZYBM"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("JRZYXM"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}

		result = ps.executeBatch();

		// �ύ����
		conn.commit();

		// �ر���Դ
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	public static void getBIFile(HttpServletRequest request,
			HttpServletResponse response) throws Exception {

		InputStream in = (InputStream) request.getAttribute("uploadFile");
		String fileName = request.getAttribute("uploadFile__name").toString();
		String fileType = fileName.substring(fileName.lastIndexOf(".") + 1,
				fileName.length());
		String imptype = request.getAttribute("imptype").toString();
		Map<String, String> map = new HashMap<String, String>();
		if ("1".equals(imptype)) { // ��������Ʒ�ߡ��ڵ㡢�ͻ�����������Ʒ�����������
			map.put("����", "DEF");
			map.put("����", "DEF1");
			map.put("��Ʒ��", "DEF2");
			map.put("�ڵ�", "DEF3");
			map.put("�ͻ�", "DEF4");
			map.put("����", "DEF5");
			map.put("��Ʒ", "DEF6");
			map.put("����", "DEF7");
			map.put("���", "DEF8");
		} else if ("2".equals(imptype)) { // ��Ʒ�ߡ��ڵ㡢�ͻ�����������Ʒ���������,��������
			map.put("����", "DEF");
			map.put("��Ʒ��", "DEF1");
			map.put("�ڵ�", "DEF2");
			map.put("�ͻ�", "DEF3");
			map.put("����", "DEF4");
			map.put("��Ʒ", "DEF5");
			map.put("�������", "DEF6");
			map.put("��������", "DEF7");
		} else if ("3".equals(imptype)) { // �ڵ㡢��׼���
			map.put("����", "DEF");
			map.put("�ڵ�", "DEF1");
			map.put("��׼���", "DEF2");
		} else if ("4".equals(imptype)) { // ��Ʒ�ߡ��ڵ㡢Ӫ��Ա��������������
			map.put("����", "DEF");
			map.put("��Ʒ��", "DEF1");
			map.put("�ڵ�", "DEF2");
			map.put("Ӫ��Ա", "DEF3");
			map.put("������", "DEF4");
			map.put("�������", "DEF5");
		} else if ("5".equals(imptype)) { // ��Ʒ�ߡ��ڵ㡢Ӫ��Ա���ͻ���������������
			map.put("����", "DEF");
			map.put("��Ʒ��", "DEF1");
			map.put("�ڵ�", "DEF2");
			map.put("Ӫ��Ա", "DEF3");
			map.put("�ͻ�", "DEF4");
			map.put("������", "DEF5");
			map.put("�������", "DEF6");
		} else if ("6".equals(imptype)) { // ��Ʒ�ߡ��ڵ�
			map.put("����", "DEF");
			map.put("��Ʒ��", "DEF1");
			map.put("�ڵ�", "DEF2");
		}

		readBI(in, fileType, map, request, response);

	}
	
	/**
	 * 29���õ���
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_costimport(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		sql = "DELETE wb_erp.SBT_COSTIMPORT WHERE MONTH in "+Times;
		PreparedStatement ps2 = conn.prepareStatement(sql);
		ps2.executeUpdate();
		DbUtil.closeStatement(ps2);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? AND MONTH=?";
			PreparedStatement ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));
			DbUtil.setObject(ps1, 5, Types.VARCHAR, vo.get("MONTH"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("��" + f + "��" + vo.get("ORG4").toString()
						+ "���Ų����ڣ���������µ��룡ע����0�����һ��");
			}

			if (f == 0) {
				// ��ɾ��,�����籣�����·ݼ���Ŀ�Ľ��и���ɾ��
				sql = "DELETE wb_erp.SBT_DetailsCharges WHERE month = ? and km4=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("KM4"));
				ps1.executeUpdate();
			}
			DbUtil.closeStatement(ps1);

			sql = "insert into wb_erp.SBT_COSTIMPORT (id,month,psnname,psncode,cpxx,ORG1,  ORG2,  ORG3,	ORG4,price,	KM1,	KM2,	KM3,	KM4,TS,	COPERATOR) "
					+ "values (?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("PSNNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("PSNCODE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CPXX"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, request
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
	
	private static void imp_sbt_guding(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		String Times = "(";
		DecimalFormat decimalFormat = new DecimalFormat("###################.###########");
		for(int f = 0; f < voList.size(); f++) {
			JSONObject obj = voList.get(f);
			if(Times.indexOf(decimalFormat.format(obj.get("MONTH")))==-1){
				if(f==0) 
				{
					Times += "'"+decimalFormat.format(obj.get("MONTH"))+"'";
				}
				else 
				{
					Times += ",'"+decimalFormat.format(obj.get("MONTH"))+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.sbt_detailsguding WHERE MONTH in "+Times;
		PreparedStatement ps2 = conn.prepareStatement(sql);
		ps2.executeUpdate();
		DbUtil.closeStatement(ps2);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? AND MONTH=?";
			PreparedStatement ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));
			DbUtil.setObject(ps1, 5, Types.VARCHAR, vo.get("MONTH"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("��" + f + "��" + vo.get("ORG4").toString()
						+ "���Ų����ڣ���������µ��룡ע����0�����һ��");
			}

			if (f == 0) {
				// ��ɾ��,�����籣�����·ݼ���Ŀ�Ľ��и���ɾ��
				sql = "DELETE wb_erp.SBT_DetailsCharges WHERE month = ? and km4=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("KM4"));
				ps1.executeUpdate();
			}
			DbUtil.closeStatement(ps1);

			sql = "insert into wb_erp.sbt_detailsguding (id,month,cpx,ORG1,  ORG2,  ORG3,	ORG4,money,	KM1,	KM2,	KM3,	KM4,PZH,TS,	COPERATOR) "
					+ "values (?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("PZH"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, request
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
	
	private static void imp_saletarget(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//ɾ����ʷ����
			sql = "DELETE wb_erp.SALETARGET WHERE ZZNAME=? and TYPE =? and MONTH=? and MEASURE=? and PRODLINENAME=?";
			PreparedStatement ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ZZNAME"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("TYPE"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("MONTH"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("MEASURE"));
			DbUtil.setObject(ps1, 5, Types.VARCHAR, vo.get("PRODLINENAME"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);

			sql = "insert into wb_erp.SALETARGET (ID,ZZNAME,MONTH,TYPE,TARGET,MEASURE,PRODLINENAME)"
					+ "values (? ,	?,	?,	?,	?,	?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("ZZNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("TARGET"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("MEASURE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PRODLINENAME"));
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
	 * ��ȡ�ļ�
	 * 
	 * @param in
	 * @param fileType
	 * @throws Exception
	 */
	public static void readBI(InputStream in, String fileType,
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
								if (headRow.getCell(j) != null
										&& headRow.getCell(j)
												.getStringCellValue().contains(
														"����")
										&& HSSFDateUtil
												.isCellDateFormatted(cell)) {
									cellVal = cell.getDateCellValue();
								} else {
									cellVal = ExcelObject.getCellValue(cell);
								}
								jsonObject.put(map.get(headRow.getCell(j)
										.getStringCellValue().toString()),
										cellVal);
							}
							list.add(jsonObject);
						}
					}
				}

			} else {
				throw new Exception("ֻ֧��2003�汾��Excel���룡");
			}

			for (int f = 0; f < list.size(); f++) {
				impBI(list.get(f), request, response);
			}
			in.close();
		} catch (Exception e) {
			dqrow = dqrow + 1;
			throw e;

		}
	}

	/**
	 * �����г���Χά�� yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void impBI(JSONObject vo, HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		String PK_ID = null;
		String sql = "";
		int result = 0;
		if (null != vo) {
			Connection conn = DbUtil.getConnection();
			DbUtil.startTrans(conn, "");
			PreparedStatement ps = null;
			PreparedStatement ps1 = null;
			String imptype = request.getAttribute("imptype").toString();

			// ��ɾ��
			sql = "DELETE wb_erp.BI_EXCELTABLE WHERE IMP_TYPE = ?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.CHAR, imptype);
			result = ps1.executeUpdate();

			// ����
			sql = "INSERT INTO wb_erp.BI_EXCELTABLE "
					+ "(PK_ID, IMP_TYPE,DEF, DEF1, DEF2, DEF3, DEF4, DEF5, DEF6, DEF7, DEF8,ts) "
					+ "VALUES " + "(?, ?,?, ?, ?, ?, ?, ?, ?, ?, ?,sysdate)";
			ps = conn.prepareStatement(sql);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.CHAR, imptype);
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("DEF"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("DEF1"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("DEF2"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("DEF3"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("DEF4"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("DEF5"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("DEF6"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("DEF7"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("DEF8"));
			result = ps.executeUpdate();

			// �ύ����
			conn.commit();

			// �ر���Դ
			DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);
			DbUtil.closeConnection(conn);
		}

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