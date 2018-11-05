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
public class ExcelImportMonthlyProfit {
	public static void getFile(HttpServletRequest request,
			HttpServletResponse response) throws Exception {

		InputStream in = (InputStream) request.getAttribute("uploadFile");
		String fileName = request.getAttribute("uploadFile__name").toString();
		String fileType = fileName.substring(fileName.lastIndexOf(".") + 1,
				fileName.length());
		String imptype = request.getAttribute("imptype").toString();
		Map<String, String> map = new HashMap<String, String>();
		if ("1".equals(imptype)) { //NC�¹�������
			map.put("�¹�������", "NEWUNITNAME");
			map.put("ȡ����������", "QSUNITNAME");
			map.put("���ȼ�", "LEVELS");
			map.put("����", "CREATETIME");
		}
		else if("2".equals(imptype)){//Ԥ���ϵ���
			map.put("��˾����","UNITNAME");
			map.put("��Ʒ����","INVCODE");
			map.put("��Ʒ����","INVNAME");
			map.put("�ֵ���","PRICE");
			map.put("����", "CREATETIME");
		}
		else if("3".equals(imptype)) {//��۵���
			map.put("��˾����", "UNITNAME");
			map.put("�䷽����", "INVCODE");
			map.put("�䷽����", "INVNAME");
			map.put("�ֵ���", "PRICE");
			map.put("����", "CREATETIME");
		}
		else if("4".equals(imptype)) {//����ɱ�
			map.put("������������", "UNITNAME");
			map.put("��������", "EVENT");
			map.put("�ɱ��������", "INVCODE_PF");
			map.put("�ɱ�����", "INVNAME_PF");
			map.put("�������", "INVCODE");
			map.put("�������", "INVNAME");
			map.put("���", "PRICE");
			map.put("����", "CREATETIME");
		}
		else if("5".equals(imptype)) {//�����ɱ�����
			map.put("��˾", "UNITNAME");
			map.put("���۷���","SaleType");
			map.put("�������","INVCODE");
			map.put("�������","INVNAME");
			map.put("����", "PROPERTY");
			map.put("�������򷽳�Ʒ�ɱ�����","PRICE");
			map.put("���������̯","FT");
			map.put("����", "CREATETIME");
		}
		else if("6".equals(imptype)){//�ɹ�KCP�������
			map.put("��˾", "UNITNAME");
			map.put("�������", "INVCODE");
			map.put("�������", "INVNAME");
			map.put("����KCP���䣨�ֵ�����", "YMKCP");
			map.put("����KCP���䣨�ֵ�����", "DPKCP");
			map.put("�մ��ʽ���Ϣ���䣨�ֵ�����", "CCZJ");
			map.put("������KCP���䣨�ֵ�����", "JCKKCP");
			map.put("����", "CREATETIME");
		}
		else if("7".equals(imptype)) {//�����Ϣ
			map.put("��˾", "UNITNAME");
			map.put("�ּ���", "JT");
			map.put("����", "CREATETIME");
		}
		else if("8".equals(imptype)) {//�������۾ɷ�̯
			map.put("��˾", "UNITNAME");
			map.put("��̯���", "PRICE");
			map.put("����", "CREATETIME");
		}
		else if("9".equals(imptype)) {//�̶��ʲ�����
			map.put("ս��", "VSALESTRUNAME");
			map.put("�ּ���", "JT");
			map.put("����", "CREATETIME");
		}
		else if("10".equals(imptype)) {//������̯
			map.put("ս��", "VSALESTRUNAME");
			map.put("��˾����", "UNITCODE");
			map.put("��˾", "UNITNAME");
			map.put("���", "PRICE");
			map.put("����", "CREATETIME");
		}
		//Ӧ�޳�NC����ɱ�
		else if("11".equals(imptype)) {
			map.put("��˾", "UNITNAME");
			map.put("�������","INVCODE");
			map.put("�������","INVNAME");
			map.put("����", "PROPERTY");
			map.put("Ӧ�޳�NC�䶯����ɱ�", "PRICE1");
			map.put("Ӧ�޳�NC�̶�����ɱ���ԭʼ��", "PRICE2");
			map.put("Ӧ�޳�NC����ɱ��ϼƣ�ԭʼ��", "PRICE3");
			map.put("�䶯�ɱ�������","BDCB");
			map.put("�̶��ɱ�������","GDCB");
			map.put("����", "CREATETIME");
		}
		else if("12".equals(imptype)) {
			map.put("��˾", "UNITNAME");
			map.put("�������","INVCODE");
			map.put("�������","INVNAME");
			map.put("�ֳɱ�", "PRICE");
			map.put("����", "CREATETIME");
		}
		else if("13".equals(imptype)) {//�¹������չ�����
			map.put("��˾", "UNITNAME");
			map.put("�������", "INVCODE");
			map.put("�������", "INVNAME");
			map.put("�ֳɱ�", "PRICE");
			map.put("ȡ����������", "FETCHUNITNAME");
			map.put("���ȼ�", "LEVELS");
			map.put("ԭ����", "PRICE1");
			map.put("�䶯����", "PRICE2");
			map.put("�̶�����", "PRICE3");
			map.put("����", "CREATETIME");
			map.put("����","PROPERTION");
		}
		else if("14".equals(imptype)) {//�ɱ��ۺϵ���
			map.put("��˾����", "UNITNAME");
			map.put("��Ʒ����", "INVCODE");
			map.put("��Ʒ����", "INVNAME");
			map.put("�䷽����", "INVCODE_PF");
			map.put("�䷽����", "INVNAME_PF");
			map.put("�ֵ���", "PRICE");
			map.put("�������", "CB");
			map.put("��ע", "BZ");
			map.put("����", "CREATETIME");
		}
		else if("15".equals(imptype)) {
			map.put("�ͻ�����", "CUSTCODE");
			map.put("�ͻ�����", "CUSTNAME");
			map.put("С��", "OLDXQ");
			map.put("����С��", "NEWXQ");
			map.put("����", "CREATETIME");
		}
		else if("16".equals(imptype)) {
			map.put("����", "UNITNAME");
			map.put("���", "YEARS");
			map.put("1�»���", "RATE1");
			map.put("2�»���", "RATE2");
			map.put("3�»���", "RATE3");
			map.put("4�»���", "RATE4");
			map.put("5�»���", "RATE5");
			map.put("6�»���", "RATE6");
			map.put("7�»���", "RATE7");
			map.put("8�»���", "RATE8");
			map.put("9�»���", "RATE9");
			map.put("10�»���", "RATE10");
			map.put("11�»���", "RATE11");
			map.put("12�»���", "RATE12");
		}
		else if("17".equals(imptype)) {
			map.put("�·�", "MONTHS");
			map.put("��˾", "UNITNAME");
			map.put("������������", "QTSR");
			map.put("˰������", "SWSR");
			map.put("�������", "CWFY");
		}
		else if("18".equals(imptype)) {
			map.put("���۷���", "SALETYPE");
			map.put("���۴���", "DOCNAME");
			map.put("�Ͻ�������˾����-�繫˾", "PRICE1");
			map.put("������������-�繫˾", "PRICE2");
			map.put("���������׼", "STANDPRICE1");
			map.put("�Ͻ���Ʒ�������׼", "STANDPRICE2");
			map.put("����", "CREATETIME");
		}
		else if("19".equals(imptype)) {
			map.put("�������", "UNITNAME");
			map.put("NC�ͻ�����", "BSCNAME");
			map.put("���", "PRICE");
			map.put("����", "MONTH");
		}
		else if("20".equals(imptype)) {
			map.put("�㼶", "CJ");
			map.put("ս��", "ZQ");
			map.put("ʡ��", "SQ");
			map.put("��", "B");
			map.put("������ò�������", "TYPE");
			map.put("����", "CREATETIME");
		}
		else if("21".equals(imptype)) {
			map.put("��Ŀ", "PRO");
			map.put("���۷���", "SALETYPE");
			map.put("��������������", "QTBLR");
			map.put("�����Ӽ۱�׼", "JJBZ");
			map.put("�ֱ䶯����ṹ", "BDZZJG");
			map.put("����", "CREATETIME");
		}
		else if("22".equals(imptype)) {
			map.put("����", "UNITNAME");
			map.put("���", "SHORTNAME");
			map.put("��˾����Ӫ����", "YXB");
		}
		else if("23".equals(imptype)) {
			map.put("�·�", "CREATETIME");
			map.put("Ӫ����", "YXB");
			map.put("С��", "XQ");
			map.put("��Ʒ��", "CPX");
			map.put("��Ŀ", "PRO");
			map.put("��̯��ʽ", "TYPE");
			map.put("��Ԫ��", "PRICE");
			map.put("˵��", "BZ");
		}
		else if("24".equals(imptype)) {
			map.put("�·�", "CREATETIME");
			map.put("Ӫ����", "YXB");
			map.put("���۷���", "SALETYPE");
			map.put("��Ʒ����", "INVCODE");
			map.put("��Ʒ����", "INVNAME");
			map.put("��Ʒ��", "CPX");
		}
		else if("25".equals(imptype)) {
			map.put("�·�", "CREATETIME");
			map.put("����", "YXB");
			map.put("��Ʒ��", "CPX");
		}
		else if("26".equals(imptype)) {
			map.put("�·�", "CREATETIME");
			map.put("�㼶", "CJ");
			map.put("����", "B");
			map.put("��Ʒ��", "CPX");
			map.put("��Ԫ��", "PRICE");
			map.put("˵��", "BZ");
		}
		else if("27".equals(imptype)) {
			map.put("�·�", "CREATETIME");
			map.put("��˾","UNITNAME");
			map.put("����Ʒ����", "INVCODE");
			map.put("����Ʒ����","INVNAME");
			map.put("����Ʒ����","PRICE" );
		}
		else if("28".equals(imptype)) {
			map.put("�·�", "CREATETIME");
			map.put("��˾","UNITNAME");
			map.put( "�ƽ���Ʒ����","INVCODE");
			map.put( "�ƽ���Ʒ","INVNAME");
			map.put("��Ʒ��","CPX" );
			map.put("��������","TYPE" );
			map.put( "�������","PRICE");
			map.put( "˵��","BZ");
		}
		else if("29".equals(imptype)) {
			map.put("�·�", "CREATETIME");
			map.put("�㼶","CJ");
			map.put("��Ʒ��","CPX" );
			map.put("ë��ϵ��","RATE" );
		}
		else if("30".equals(imptype)) {
			map.put("�·�", "CREATETIME");
			map.put("����˰��", "RATE");
		}
		else if("31".equals(imptype)) {
			map.put("�����·�", "CREATETIME");
			map.put("��������", "TYPE");
			map.put("����(�����׼��", "PRICE");
			map.put("����˵��", "BZ");
		}
		else if("32".equals(imptype)) {
			map.put("����", "UNITNAME");
			map.put("���۷���", "SALETYPE");
			map.put("�������", "INVCODE");
			map.put("�������", "INVNAME");
			map.put("ҵ������", "CPX");
			map.put("��׼ë��", "PRICE1");
			map.put("������ë��", "PRICE2");
			map.put("�·�", "CREATETIME");
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
				SimpleDateFormat format = new SimpleDateFormat("yyyy-MM");
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
									if (cellVal == null) {
										/*//����ɱ�����ģ���������
										if("4".equals(request.getAttribute(
												"imptype").toString())&&(j==2||j==3||j==4||j==5)){
											cellVal = "";
										}
										else {
											if("11".equals(request.getAttribute(
												"imptype").toString())&&(j==2||j==3||j==4||j==5)) {
												cellVal = "";
											}
											else {
												if(("14".equals(request.getAttribute(
														"imptype").toString())||"13".equals(request.getAttribute(
																"imptype").toString()))&&(j==1||j==2||j==3||j==4||j==5||j==6)) {
													cellVal = "";
												}
												else {
													if("15".equals(request.getAttribute("imptype").toString())) {
														cellVal = "";
													}
													else {
														if("21".equals(request.getAttribute("imptype").toString())||
																"22".equals(request.getAttribute("imptype").toString())||
																"23".equals(request.getAttribute("imptype").toString())||
																"25".equals(request.getAttribute("imptype").toString())||
																"26".equals(request.getAttribute("imptype").toString())||
																"27".equals(request.getAttribute("imptype").toString())||
																"32".equals(request.getAttribute("imptype").toString())) {
														}
														else {
														throw new Exception("��" + i + "��" + j
																+ "��Ϊ�գ�����д");
														}
													}
												}
											}
										}*/
										cellVal = "";
									}
								}
								jsonObject.put(map.get(headRow.getCell(j)
										.getStringCellValue().toString().trim()
										.replaceAll("\r|\n", "")), cellVal);
							}
							//System.out.print(dqrow);
							list.add(jsonObject);
						}
					}
				}
			} else {
				throw new Exception("ֻ֧��2003�汾��Excel���룡");
			}
			StringBuffer sBuffer = new StringBuffer();
			String imptype = request.getAttribute("imptype").toString();
			//����NC�¹���
			if ("1".equals(imptype)) {
				imp_ncnewcompany(list, request, response, imptype);
			} 
			//����Ԥ���ϵ���
			else if("2".equals(imptype)) {
				imp_yhladjust(list, request, response, imptype);
			}
			//������۵���
			else if("3".equals(imptype)) {
				imp_yfadjust(list, request, response, imptype);
			}
			//��������ɱ�
			else if("4".equals(imptype)) {
				imp_specialcost(list, request, response, imptype);
			}
			//��������ɱ�����
			else if("5".equals(imptype)) {
				imp_glcbadjust(list, request, response, imptype);
			}
			//����ɹ�KCP�������
			else if("6".equals(imptype)) {
				imp_cgkcp(list, request, response, imptype);
			}
			//����������
			else if("7".equals(imptype)) {
				imp_chjx(list, request, response, imptype);
			}
			//�����������۾ɷ�̯
			else if("8".equals(imptype)) {
				imp_zjshare(list, request, response, imptype);
			}
			//����̶��ʲ���Ϣ
			if("9".equals(imptype)) {
				imp_assetjx(list, request, response, imptype);
			}
			//����������̯
			else if("10".equals(imptype)) {
				imp_otherft(list, request, response, imptype);
			}
			//����Ӧ�޳�NC����ɱ�
			else if("11".equals(imptype)) {
				imp_rejectnccost(list, request, response, imptype);
			}
			//����NC�չ������ɱ�
			else if("12".equals(imptype)) {
				imp_sgcompany(list, request, response, imptype);
			}
			//�����¹������չ�����
			if("13".equals(imptype)) {
				imp_newcompany(list, request, response, imptype);
			}
			//����ɱ��ۺϵ���
			else if("14".equals(imptype)) {
				imp_zhadjust(list, request, response, imptype);
			}
			//��������С��
			else if("15".equals(imptype)) {
				imp_xzxq(list, request, response, imptype);
			}
			//���뺣�⹤������
			else if("16".equals(imptype)) {
				imp_hwhl(list, request, response, imptype);
			}
			//������������Ͳ������
			else if("17".equals(imptype)) {
				imp_qtsr(list, request, response, imptype);
			}
			else if("18".equals(imptype)) {
				imp_qtsz(list, request, response, imptype);
			}
			else if("19".equals(imptype)) {
				imp_tcbsc(list, request, response, imptype);
			}
			else if("20".equals(imptype)) {
				imp_fytype(list, request, response, imptype);
			}
			else if("21".equals(imptype)) {
				imp_zzfy(list, request, response, imptype);
			}
			else if("22".equals(imptype)) {
				imp_gcjc(list, request, response, imptype);
			}
			else if("23".equals(imptype)) {
				imp_salebt(list, request, response, imptype);
			}
			else if("24".equals(imptype)) {
				imp_sjmlpz(list, request, response, imptype);
			}
			else if("25".equals(imptype)) {
				imp_sjmlorgcpx(list, request, response, imptype);
			}
			else if("26".equals(imptype)) {
				imp_mlcd(list, request, response, imptype);
			}
			else if("27".equals(imptype)) {
				imp_cxpproduct(list, request, response, imptype);
			}
			else if("28".equals(imptype)) {
				imp_jtcxfy(list, request, response, imptype);
			}
			else if("29".equals(imptype)) {
				imp_mlxs(list, request, response, imptype);
			}
			else if("30".equals(imptype)) {
				imp_sds(list, request, response, imptype);
			}
			else if("31".equals(imptype)) {
				imp_fyjt(list, request, response, imptype);
			}
			else if("32".equals(imptype)) {
				imp_bzml(list, request, response, imptype);
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

	//NC�¹�������
	private static void imp_ncnewcompany(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_newnccompany WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// У���¹�˾�����Ƿ����
			if (!vo.get("NEWUNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("NEWUNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "���¹�˾���벻���ڣ���������µ��룡");
				}
			}
			
			// У��ȡ����˾�����Ƿ����
			if (!vo.get("QSUNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("QSUNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "��ȡ����˾���벻���ڣ���������µ��룡");
				}
			}
			
			sql = "insert into wb_erp.app_main_newnccompany (ID,Newpk_corp,Fetchpk_corp,COPERATOR,CreateTime,LEVELS)"
					+ "values (?,(select pk_corp FROM wb_erp.bd_corp p where p.memo=?),(select pk_corp FROM wb_erp.bd_corp p where p.memo=?),?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("NEWUNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("QSUNITNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CREATETIME"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("LEVELS"));
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

	//Ԥ���ϵ�������
	private static void imp_yhladjust(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_yhladjust WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// У���¹�˾�����Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�й�˾���벻���ڣ���������µ��룡");
				}
			}
			
			//У���Ʒ�����Ƿ����
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�в�Ʒ���벻���ڣ���������µ��룡");
				}
			}
			
			sql = "insert into wb_erp.app_main_yhladjust (id,pk_corp,pk_invbasdoc,InvName,PRICE,COPERATOR,CreateTime)"
					+ "values (?,(select pk_corp FROM wb_erp.bd_corp p where p.MEMO=?),(select pk_invbasdoc FROM WB_ERP.bd_invbasdoc B WHERE invcode=?),?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("INVCODE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("CREATETIME"));

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
	
	//��۵�������
	private static void imp_yfadjust(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_yfadjust WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//У�鹫˾�����Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�й�˾���벻���ڣ���������µ��룡");
				}
			}
			
			//У���䷽�����Ƿ����
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "���䷽���벻���ڣ���������µ��룡");
				}
			}
			
			sql = "insert into wb_erp.app_main_yfadjust (id,pk_corp,pk_invbasdoc,InvName,PRICE,COPERATOR,CreateTime)"
					+ "values (?,(select pk_corp FROM wb_erp.bd_corp p where p.MEMO=?),(select pk_invbasdoc FROM WB_ERP.bd_invbasdoc B WHERE invcode=?),?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("INVCODE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("CREATETIME"));

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

	//����ɱ�����
	private static void imp_specialcost(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_specialcost WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//У��������������Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�е����������벻���ڣ���������µ��룡");
				}
			}
			
			if(vo.get("INVCODE_PF").toString().equals("")&&vo.get("INVCODE").toString().equals("")) {
				throw new Exception("���������" + f + "�д��������ɱ�����������ڣ���������µ��룡");
			}
			
			//У��ɱ���������Ƿ����
			if (!vo.get("INVCODE_PF").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.BD_INVBASDOC B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE_PF"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�гɱ�������벻���ڣ���������µ��룡");
				}
			}
			
			//У���������Ƿ����
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.BD_INVBASDOC B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�д�����벻���ڣ���������µ��룡");
				}
			}
			
			sql = "insert into wb_erp.app_main_specialcost (id,pk_corp,Event,pk_invbasdoc_pf,InvName_PF,pk_invbasdoc,InvName,Price,COPERATOR,CreateTime)"
					+ "values (?,(select pk_corp FROM wb_erp.bd_corp p where p.memo=?),?,(select pk_invbasdoc FROM WB_ERP.bd_invbasdoc B WHERE invcode=?),?,(select pk_invbasdoc FROM WB_ERP.bd_invbasdoc B WHERE invcode=?),?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("EVENT"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("INVCODE_PF"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("INVNAME_PF"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("INVCODE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("CREATETIME"));

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
	
	//�����ɱ���������
	private static void imp_glcbadjust(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.APP_MAIN_GLCBADJUST WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//У��������������Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE MEMO=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�е����������벻���ڣ���������µ��룡");
				}
				DbUtil.closeStatement(ps1);
			}
			
			//У���������Ƿ����
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�д�����벻���ڣ���������µ��룡");
				}
				DbUtil.closeStatement(ps1);
			}
			
			sql = "insert into wb_erp.APP_MAIN_GLCBADJUST (id,pk_corp,DOCNAME,pk_invbasdoc,InvName,PROPERTY,PRICE,ft,COPERATOR,CreateTime)"
					+ "values (?,(select pk_corp FROM wb_erp.bd_corp p where p.memo=?),?,(select pk_invbasdoc FROM WB_ERP.bd_invbasdoc B WHERE invcode=?),?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("SaleType"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("INVCODE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("PROPERTY"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("FT"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("CREATETIME"));

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

	//�ɹ�KCP������䵼��
	private static void imp_cgkcp(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_purkcpprofit WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//У��������������Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE MEMO=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�е����������벻���ڣ���������µ��룡");
				}
				 DbUtil.closeStatement(ps1);
			}
			
			//У���������Ƿ����
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�д�����벻���ڣ���������µ��룡");
				}
				 DbUtil.closeStatement(ps1);
			}
			
			sql = "insert into wb_erp.app_main_purkcpprofit (id,pk_corp,pk_invbasdoc,invname,YMKCP,DPKCP,CCZJ,JCKKCP,COPERATOR,CreateTime)"
					+ "values (?,(select pk_corp FROM wb_erp.bd_corp p where p.memo=?),(select pk_invbasdoc FROM WB_ERP.bd_invbasdoc B WHERE invcode=?),?,?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("INVCODE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("YMKCP"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("DPKCP"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("CCZJ"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("JCKKCP"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("CREATETIME"));

			ps.execute();

			// �ύ����
			System.out.println(f);
			// �ر���Դ
			 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}		

	//�����Ϣ
	private static void imp_chjx(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_chjx WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//У��������������Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�е����������벻���ڣ���������µ��룡");
				}
			}
			
			/*//У���������Ƿ����
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE unitcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�д�����벻���ڣ���������µ��룡");
				}
			}*/
			
			sql = "insert into wb_erp.app_main_chjx (id,pk_corp,JT,COPERATOR,CreateTime)"
					+ "values (?,(select pk_corp FROM wb_erp.bd_corp p where p.MEMO=?),?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("JT"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CREATETIME"));

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

	//�������۾ɷ�̯����
	private static void imp_zjshare(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_zjshare WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//У��������������Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE MEMO=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�е����������벻���ڣ���������µ��룡");
				}
			}
			
			/*//У���������Ƿ����
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE unitcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�д�����벻���ڣ���������µ��룡");
				}
			}*/
			
			sql = "insert into wb_erp.app_main_zjshare (id,pk_corp,PRICE,COPERATOR,CreateTime)"
					+ "values (?,(select pk_corp FROM wb_erp.bd_corp p where p.memo=?),?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CREATETIME"));

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
	
	//�̶��ʲ����ᵼ��
	private static void imp_assetjx(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;
		//ResultSet rSet = null;
		//int result2 = 0;
		String Times = "(";
		for(int f = 0; f < voList.size(); f++) {
			JSONObject obj = voList.get(f);
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_assetjx WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			/*//У��������������Ƿ����
			if (!vo.get("UNITCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE unitcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�е����������벻���ڣ���������µ��룡");
				}
			}*/
			
			/*//У���������Ƿ����
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE unitcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�д�����벻���ڣ���������µ��룡");
				}
			}*/
			
			sql = "insert into wb_erp.app_main_assetjx (id,VSALESTRUNAME,JT,COPERATOR,CreateTime)"
					+ "values (?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("VSALESTRUNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("JT"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CREATETIME"));

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
	
	//������̯����
	private static void imp_otherft(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_otherft WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//У��������������Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE MEMO=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�е����������벻���ڣ���������µ��룡");
				}
			}
			
			/*//У���������Ƿ����
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE unitcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�д�����벻���ڣ���������µ��룡");
				}
			}*/
			
			sql = "insert into wb_erp.app_main_otherft (id,VSALESTRUNAME,pk_corp,price,COPERATOR,CreateTime)"
					+ "values (?,?,(select pk_corp FROM wb_erp.bd_corp p where p.MEMO=?),?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("VSALESTRUNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("CREATETIME"));

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
	
	//Ӧ�޳�NC����ɱ�����
	private static void imp_rejectnccost(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_rejectnccost WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//У��������������Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�е����������벻���ڣ���������µ��룡");
				}
				DbUtil.closeStatement(ps1);
			}
			
			//У���������Ƿ����
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�д�����벻���ڣ���������µ��룡");
				}
				DbUtil.closeStatement(ps1);
			}
			
			sql = "insert into wb_erp.app_main_rejectnccost (id,pk_corp,pk_invbasdoc,invname,property,price1,price2,price3,BDCB,GDCB,COPERATOR,CreateTime)"
					+ "values (?,(select pk_corp FROM wb_erp.bd_corp p where p.memo=?),(select pk_invbasdoc FROM WB_ERP.bd_invbasdoc B WHERE invcode=?),?,?,?,?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("INVCODE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PROPERTY"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("PRICE1"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PRICE2"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PRICE3"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("BDCB"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("GDCB"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("CREATETIME"));

			ps.execute();

			// �ύ����
			//System.out.println(f);
			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}
	
	//NC�չ������ɱ�����
	private static void imp_sgcompany(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.APP_MAIN_SGCOMPANY WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//У��������������Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�е����������벻���ڣ���������µ��룡");
				}
			}
			
			//У���������Ƿ����
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�д�����벻���ڣ���������µ��룡");
				}
			}
			
			sql = "insert into wb_erp.APP_MAIN_SGCOMPANY (id,pk_corp,pk_invbasdoc,invname,price,COPERATOR,CreateTime)"
					+ "values (?,(select pk_corp FROM wb_erp.bd_corp p where p.memo=?),(select pk_invbasdoc FROM WB_ERP.bd_invbasdoc B WHERE invcode=?),?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("INVCODE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("CREATETIME"));

			ps.execute();

			// �ύ����
			//System.out.println(f);
			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}
	
	//�¹������չ���������
	private static void imp_newcompany(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.APP_MAIN_NEWCOMPANY WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//У�鹤�������Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�й������벻���ڣ���������µ��룡");
				}
			}
			
			//У��������������Ƿ����
			if (!vo.get("FETCHUNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("FETCHUNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�е����������벻���ڣ���������µ��룡");
				}
			}
			
			//У���������Ƿ����
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�д�����벻���ڣ���������µ��룡");
				}
			}
			
			sql = "insert into wb_erp.APP_MAIN_NEWCOMPANY (id,pk_corp,pk_invbasdoc,invname,price,fetchpk_corp,levels,price1,price2,price3,COPERATOR,CreateTime,PROPERTION)"
					+ "values (?,(select pk_corp FROM wb_erp.bd_corp p where p.memo=?),(select pk_invbasdoc FROM WB_ERP.bd_invbasdoc B WHERE invcode=?),?,?,(select pk_corp FROM wb_erp.bd_corp p where p.memo=?),?,?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("INVCODE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("FETCHUNITNAME"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("LEVELS"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PRICE1"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("PRICE2"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("PRICE3"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("CREATETIME"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("PROPERTION"));
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
	
	//�ɱ��ۺϵ�������
	private static void imp_zhadjust(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			if(f==4705) {
				//System.out.println(f);
			}
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_zhadjust WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//У��������������Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�е����������벻���ڣ���������µ��룡");
				}
				DbUtil.closeStatement(ps1);
			}
			
			/*if(vo.get("INVCODE_PF").toString().equals("")&&vo.get("INVCODE").toString().equals("")&&!vo.get) {
				throw new Exception("���������" + f + "�д��������ɱ�����������ڣ���������µ��룡");
			}*/
			
			//У��ɱ���������Ƿ����
			if (!vo.get("INVCODE_PF").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.BD_INVBASDOC B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE_PF"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�гɱ�������벻���ڣ���������µ��룡");
				}
				DbUtil.closeStatement(ps1);
			}
			
			//У���������Ƿ����
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.BD_INVBASDOC B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�д�����벻���ڣ���������µ��룡");
				}
				DbUtil.closeStatement(ps1);
			}
			
			sql = "insert into wb_erp.app_main_zhadjust (id,pk_corp,pk_invbasdoc,InvName,pk_invbasdoc_PF,InvName_PF,Price,CB,BZ,COPERATOR,CreateTime)"
					+ "values (?,(select pk_corp FROM wb_erp.bd_corp p where p.memo=?),(select pk_invbasdoc FROM WB_ERP.bd_invbasdoc B WHERE invcode=?),?,(select pk_invbasdoc FROM WB_ERP.bd_invbasdoc B WHERE invcode=?),?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("INVCODE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("INVCODE_PF"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("INVNAME_PF"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("CB"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("BZ"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("CREATETIME"));

			ps.execute();

			// �ύ����
			//System.out.println(f);
			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}
	
	//����С������
	private static void imp_xzxq(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			//System.out.println(f);
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_xzxq WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_xzxq (id,custcode,custname,oldxq,newxq,createtime)"
					+ "values (?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("OLDXQ"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("NEWXQ"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("CREATETIME"));

			ps.execute();
				
	
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));

			// �ύ����
			//System.out.println(f);
			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}
	
	//������ʵ���
	private static void imp_hwhl(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
		PreparedStatement ps1;
		for(int f = 0; f < voList.size(); f++) {
			JSONObject obj = voList.get(f);
			//ɾ����������ʷ����
			sql = "DELETE wb_erp.app_main_hwhl WHERE YEARS = "+obj.get("YEARS").toString()+" and pk_corp=(select pk_corp from wb_erp.bd_corp where memo='"+obj.get("UNITNAME").toString()+"')";
			ps1 = conn.prepareStatement(sql);
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//�жϹ����Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�е����������벻���ڣ���������µ��룡");
				}
			}
			
			sql = "insert into wb_erp.app_main_hwhl (id,pk_corp,unitname,years,rate1,rate2,rate3,rate4,rate5,rate6,rate7,rate8,rate9,rate10,rate11,rate12)"
					+ "values (?,(select pk_corp from wb_erp.bd_corp where memo = ?),?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("YEARS"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("RATE1"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("RATE2"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("RATE3"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("RATE4"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("RATE5"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("RATE6"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("RATE7"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("RATE8"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("RATE9"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("RATE10"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("RATE11"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("RATE12"));
			ps.execute();

			// �ύ����
			//System.out.println(f);
			// �ر���Դ
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}
	
	//������������Ͳ������
	private static void imp_qtsr(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			//System.out.println(f);
			if(Times.indexOf(obj.get("MONTHS").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("MONTHS").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("MONTHS").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_qtsr WHERE MONTHS in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//�жϹ����Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�е����������벻���ڣ���������µ��룡");
				}
			}
			
			sql = "insert into wb_erp.app_main_qtsr (id,pk_corp,unitname,months,qtsr,swsr,cwfy)"
					+ "values (?,(select pk_corp from wb_erp.bd_corp where memo = ?),?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("MONTHS"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("QTSR"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("SWSR"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("CWFY"));

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
	
	//����������֧
	private static void imp_qtsz(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			System.out.println(f);
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_qtsz WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//�жϹ����Ƿ����
//			if (!vo.get("UNITNAME").toString().equals("")) {
//				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
//				ps1 = conn.prepareStatement(sql);
//				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
//				rSet = ps1.executeQuery();
//				if (rSet.next()) {
//					result2 = rSet.getInt("CT");
//				}
//				if (result2 == 0) {
//					throw new Exception("���������" + f + "�е����������벻���ڣ���������µ��룡");
//				}
//			}
			
			sql = "insert into wb_erp.app_main_qtsz (id,SALETYPE,DOCNAME,PRICE1,PRICE2,STANDPRICE1,STANDPRICE2,CREATETIME)"
					+ "values (?,?,?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SALETYPE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("DOCNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("PRICE1"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PRICE2"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("STANDPRICE1"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("STANDPRICE2"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("CREATETIME"));

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
	
	//Ӧ�޳����´�����
	private static void imp_tcbsc(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			System.out.println(f);
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
		sql = "DELETE wb_erp.app_main_tcbsc WHERE MONTH in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//�жϹ����Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "������������Ʋ����ڣ���������µ��룡");
				}
			}
			
			sql = "insert into wb_erp.app_main_tcbsc (id,pk_corp,unitname,bscname,price,month)"
					+ "values (?,(select pk_corp FROM WB_ERP.bd_corp B WHERE memo=?),?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("BSCNAME"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("MONTH"));

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
	
	private static void imp_fytype(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			System.out.println(f);
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_fytype WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_fytype (id,CJ,ZQ,SQ,B,TYPE,CREATETIME)"
					+ "values (?,?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CJ"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ZQ"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("SQ"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("B"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("CREATETIME"));

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
	
	private static void imp_zzfy(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			System.out.println(f);
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_zzfy WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_zzfy (id,PRO,SALETYPE,QTBLR,JJBZ,BDZZJG,CREATETIME)"
					+ "values (?,?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("PRO"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("SALETYPE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("QTBLR"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("JJBZ"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("BDZZJG"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("CREATETIME"));

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
	
	private static void imp_gcjc(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//�жϹ����Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				PreparedStatement ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�й������Ʋ����ڣ���������µ��룡");
				}
			}
			
			sql = "insert into wb_erp.app_main_gcjc (pk_corp,shortname,yxb)"
					+ "values ((select pk_corp FROM WB_ERP.bd_corp B WHERE memo=?),?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SHORTNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("YXB"));
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
	
	private static void imp_salebt(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			System.out.println(f);
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_salebt WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_salebt (id,yxb,xq,cpx,pro,type,price,bz,createtime)"
					+ "values (?,?,?,?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("XQ"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PRO"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("BZ"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("CREATETIME"));			
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
	
	private static void imp_sjmlpz(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			System.out.println(f);
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_sjmlpz WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_sjmlpz (id,yxb,saletype,pk_invbasdoc,invname,cpx,createtime)"
					+ "values (?,?,?,(select pk_invbasdoc from wb_erp.bd_invbasdoc where invcode = ?),?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("SALETYPE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("INVCODE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("CREATETIME"));			
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

	private static void imp_sjmlorgcpx(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			System.out.println(f);
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_sjmlorgcpx WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_sjmlorgcpx (id,yxb,cpx,createtime)"
					+ "values (?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("CREATETIME"));		
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

	private static void imp_mlcd(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			System.out.println(f);
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_mlcd WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_mlcd (id,cj,b,cpx,price,bz,createtime)"
					+ "values (?,?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CJ"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("B"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("BZ"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("CREATETIME"));		
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

	private static void imp_cxpproduct(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			System.out.println(f);
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_cxpproduct WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//�жϹ����Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�й������Ʋ����ڣ���������µ��룡");
				}
			}
			//�жϲ�Ʒ�����Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�й������Ʋ����ڣ���������µ��룡");
				}
			}
			sql = "insert into wb_erp.app_main_cxpproduct (id,pk_corp,pk_invbasdoc,invname,price,createtime)"
					+ "values (?,(select pk_corp from wb_erp.bd_corp where memo = ?),(select pk_invbasdoc from wb_erp.bd_invbasdoc where invcode = ?),?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("INVCODE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("CREATETIME"));		
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

	private static void imp_jtcxfy(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			System.out.println(f);
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_jtcxfy WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//�жϹ����Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�й������Ʋ����ڣ���������µ��룡");
				}
			}
			//�жϲ�Ʒ�Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�в�Ʒ���벻���ڣ���������µ��룡");
				}
			}
			sql = "insert into wb_erp.app_main_jtcxfy (id,pk_corp,pk_invbasdoc,invname,cpx,type,price,bz,createtime)"
					+ "values (?,(select pk_corp from wb_erp.bd_corp where memo = ?),(select pk_invbasdoc from wb_erp.bd_invbasdoc where invcode = ?),?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("INVCODE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("BZ"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("CREATETIME"));		
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
	
	private static void imp_mlxs(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			System.out.println(f);
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_mlxs WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_mlxs (id,cj,cpx,rate,createtime)"
					+ "values (?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CJ"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("RATE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CREATETIME"));		
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
	
	private static void imp_sds(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
			System.out.println(f);
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_sds WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_sds (id,rate,createtime)"
					+ "values (?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("RATE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("CREATETIME"));		
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
	
	private static void imp_fyjt(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
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
		/*String Times = "(";
		for(int f = 0; f < voList.size(); f++) {
			JSONObject obj = voList.get(f);
			System.out.println(f);
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_fyjt WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);*/
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_fyjt (id,type,price,bz,createtime)"
					+ "values (?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("BZ"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CREATETIME"));
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
	
	private static void imp_bzml(List<JSONObject> voList,
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
		String Times = "(";
		for(int f = 0; f < voList.size(); f++) {
			JSONObject obj = voList.get(f);
			System.out.println(f);
			if(Times.indexOf(obj.get("CREATETIME").toString())==-1){
				if(f==0) 
				{
					Times += "'"+obj.get("CREATETIME").toString()+"'";
				}
				else 
				{
					Times += ",'"+obj.get("CREATETIME").toString()+"'";
				}
			}
		}
		Times += ")";
		//ɾ����������ʷ����
		sql = "DELETE wb_erp.app_main_bzml WHERE CREATETIME in "+Times;
		ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//�жϹ����Ƿ����
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�й������Ʋ����ڣ���������µ��룡");
				}
				DbUtil.closeStatement(ps1);
			}
			//�жϲ�Ʒ�Ƿ����
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("���������" + f + "�в�Ʒ���벻���ڣ���������µ��룡");
				}
				DbUtil.closeStatement(ps1);
			}
			sql = "insert into wb_erp.app_main_bzml (id,pk_corp,pk_invbasdoc,invname,saletype,cpx,price1,price2,createtime)"
					+ "values (?,(select pk_corp from wb_erp.bd_corp where memo = ?),(select pk_invbasdoc from wb_erp.bd_invbasdoc where invcode = ?),?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("INVCODE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("SALETYPE"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PRICE1"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PRICE2"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("CREATETIME"));		
			ps.execute();

			// �ύ����
			System.out.println(f);
			// �ر���Դ
			 DbUtil.closeStatement(ps1);
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