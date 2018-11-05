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

/**
 * EXcel 导入
 * 
 * @author Administrator
 * 
 */
public class ExcelImportOther {
	public static void getFile(HttpServletRequest request,
			HttpServletResponse response) throws Exception {

		InputStream in = (InputStream) request.getAttribute("uploadFile");
		String fileName = request.getAttribute("uploadFile__name").toString();
		String fileType = fileName.substring(fileName.lastIndexOf(".") + 1,
				fileName.length());
		String imptype = request.getAttribute("imptype").toString();
		Map<String, String> map = new HashMap<String, String>();
		if ("1".equals(imptype)) { //猪肉价格导入
			map.put("省份", "PROVINCE");
			map.put("日期", "CREATETIME");
			map.put("价格", "PRICE");
		}
		else if("2".equals(imptype)){//工厂费用导入
			map.put("销售月份","MONTH");
			map.put("数据来源","SOURCE");
			map.put("OA号","REQUESTID");
			map.put("流程名称","WORKFLOWNAME");
			map.put("发起人","FQR");
			map.put("金额","MONEY");
			map.put("支出事由","MEMO");
			map.put("受益工厂","FACTORY");
			map.put("费用类型","TYPE");
			map.put("一级科目","KM1");
			map.put("二级科目","KM2");
			map.put("三级科目","KM3");
			map.put("四级科目","KM4");
			map.put("匹配科目","CHANGEKM");
		}
		else if("3".equals(imptype)) {//抓猪目标
			map.put("客户编码","CUSTCODE");
			map.put("客户名称","CUSTNAME");
			map.put("目标月份","MONTHS");
			map.put("抓猪头数","NNUMBER");
		}
		else if("4".equals(imptype)) {
			map.put("公猪存栏", "GZCL");
			map.put("未配后备猪", "WPHBZ");
			map.put("母猪死亡数", "MZSW");
			map.put("母猪淘汰数", "MZTT");
			map.put("脱产母猪数", "TCMZ");
			map.put("新增配种母猪", "XZPZMZ");
			map.put("新增配种后备", "XZPZHB");
			map.put("返情母猪头数", "FQMZ");
			map.put("流产母猪头数", "LCMZ");
			map.put("分娩母猪头数", "FMMZ");
			map.put("产活仔总数", "TOTALCHZ");
			map.put("产健仔总数", "TOTALCJZ");
			map.put("分娩舍仔猪死淘数", "FMSZZST");
			map.put("断奶母猪数", "DNMZ");
			map.put("断奶仔猪数", "DNZZ");
			map.put("保育进栏数", "BYJL");
			map.put("转至育肥数", "YFZHB");
			map.put("转至育肥数", "ZZYF");
			map.put("保育销售数", "BYXS");
			map.put("保育死淘数", "BYST");
			map.put("育肥进栏数", "YFJL");
			map.put("育肥销售数", "YFXS");
			map.put("育肥死淘数", "YFST");
			map.put("母猪料（吨）", "MZL");
			map.put("三   宝（吨）", "SB");
			map.put("育肥料（吨）", "YFL");
			map.put("合   计（吨）", "TOTALNUMBER");
			map.put("周会/月中/月末", "ZHYZYM");
			map.put("对比/示范试验", "DBSFSY");
			map.put("5S（含全场消毒）", "V_5S");
			map.put("三改内容", "SGNR");
			map.put("后备公猪", "HBGZ");
			map.put("业务员ID", "USERID");
			map.put("客户编号", "CUSTCODE");
			map.put("创建日期", "CREATETIME");
		}
		read(in, fileType, map, request, response);

	}

	/**
	 * 读取文件
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
				/*//猪肉价格表数据读取
				if("1".equals(imptype)) {					
				}
				//工厂费用表数据读取
				if("2".equals(imptype)) {					
				}*/
				// 第一个报表
				if (sheet != null && sheet.getLastRowNum() > 0) {
					for (int i = 1; i <= sheet.getLastRowNum(); i++) {
						dqrow = i;
						JSONObject jsonObject = new JSONObject();
						// 得到当前工作表的行数
						row = sheet.getRow(i);
						if (null != row) {
							for (int j = 0; j < row.getLastCellNum(); j++) {
								if (j > headRow.getLastCellNum() - 1) {
									break;
								}
								cell = row.getCell(j);
								if (cell == null) {
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
									if (cellVal == null)
										jsonObject.put(map.get(headRow.getCell(j)
												.getStringCellValue().toString().trim()
												.replaceAll("\r|\n", "")), "");
								}
								jsonObject.put(map.get(headRow.getCell(j)
										.getStringCellValue().toString().trim()
										.replaceAll("\r|\n", "")), cellVal);
							}
							list.add(jsonObject);
						}
					}
				}
			} else {
				throw new Exception("只支持2003版本的Excel导入！");
			}
			StringBuffer sBuffer = new StringBuffer();
			String imptype = request.getAttribute("imptype").toString();
			//导入猪肉价格
			if ("1".equals(imptype)) {
				imp_areaprice(list, request, response, imptype);
			} 
			//导入工厂费用
			else if("2".equals(imptype)) {
				imp_factory(list, request, response, imptype);
			}
			//抓猪目标导入
			else if("3".equals(imptype)) {
				imp_zpigtarget(list, request, response, imptype);
			}
			else if("4".equals(imptype)) {
				imp_daytechnician(list, request, response, imptype);
			}
			in.close();
			if (sBuffer.length() > 0) {
				request.setAttribute("msg", "如下人员未导入(其他导入成功)："
						+ sBuffer.toString());
			}

		} catch (Exception e) {
			dqrow = dqrow + 1;
			throw e;

		}
	}

	private static void imp_areaprice(List<JSONObject> voList,
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_areaprice WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			/*sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? AND MONTH=?";
			
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
				throw new Exception("第" + f + "行" + vo.get("ORG4").toString()
						+ "部门不存在，请检查后重新导入！");
			}
			if (f == 0) {
				// 先删除
			}*/

		


				sql = "insert into wb_erp.app_main_areaprice (PROVINCE,	CREATETIME,	PRICE)"
						+ "values (?,	?,	?)";
				ps = conn.prepareStatement(sql);
				DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("PROVINCE"));
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CREATETIME"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("PRICE"));

				ps.execute();
	
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));

			// 提交事务
			System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	private static void imp_factory(List<JSONObject> voList,
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.sbt_factory_detail WHERE MONTH in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		/*if(voList.size()>0) {
			// 先删除
			sql = "DELETE wb_erp.sbt_factory_detail WHERE MONTH = to_char(sysdate,'yyyymm')";
			PreparedStatement ps1 = conn.prepareStatement(sql);
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}*/
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			/*sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? AND MONTH=?";
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
				throw new Exception("第" + f + "行" + vo.get("ORG4").toString()
						+ "部门不存在，请检查后重新导入！");
			}
			if (f == 0) {
				
			}*/
			
				sql = "insert into wb_erp.sbt_factory_detail (MONTH,SOURCE,REQUESTID,WORKFLOWNAME,FQR,MONEY,MEMO,FACTORY,TYPE,KM1,KM2,KM3,KM4,CHANGEKM)"
						+ "values (?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
				ps = conn.prepareStatement(sql);
				DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SOURCE"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("REQUESTID"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("WORKFLOWNAME"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("FQR"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("MONEY"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("MEMO"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("FACTORY"));
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("TYPE"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KM1"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("KM2"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("KM3"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("KM4"));
				DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("CHANGEKM"));

				ps.execute();
	
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));

			// 提交事务
			System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}
	
	private static void imp_zpigtarget(List<JSONObject> voList,
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
		//删除包含的历史数据
		for(int f = 0; f < voList.size(); f++) {
			JSONObject obj = voList.get(f);
			//删除历史导入行
			sql = "DELETE wb_erp.APP_MAIN_ZPIGTARGET WHERE MONTHS = '" + obj.get("MONTHS").toString() 
				+ "' and PK_CUSBASDOC = (select pk_cubasdoc from WB_ERP.BD_CUBASDOC where custcode = '" + obj.get("CUSTCODE").toString() +"')";
		}
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			//校验客户编码是否存在
			if (!vo.get("CUSTCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.BD_CUBASDOC B WHERE CUSTCODE=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("CUSTCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行客户编码不存在，请检查后重新导入！");
				}
			}
			
			sql = "insert into wb_erp.APP_MAIN_ZPIGTARGET (ID,PK_CUSBASDOC,MONTHS,NNUMBER)"
					+ "values (?,(select PK_CUBASDOC as CT FROM WB_ERP.BD_CUBASDOC B WHERE CUSTCODE=?),?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("MONTHS"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("NNUMBER"));

			ps.execute();

			// 提交事务
			System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	private static void imp_daytechnician(List<JSONObject> voList,
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
		//删除包含的历史数据
		for(int f = 0; f < voList.size(); f++) {
			JSONObject obj = voList.get(f);
			//删除历史导入行
			sql = "DELETE wb_erp.app_main_daytechnician WHERE CREATETIME = '" + obj.get("CREATETIME").toString() 
				+ "' and custcode = '" + obj.get("CUSTCODE").toString() +"'";
		}
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_daytechnician (id,gzcl,wphbz,mzsw,mztt,tcmz,xzpzmz,xzpzhb,fqmz,lcmz,fmmz,totalchz,totalcjz,fmszzst,dnmz,dnzz,byjl,zzyf,byxs,byst,yfjl,yfzhb,yfxs,yfst,mzl,sb,yfl,totalnumber,zhyzym,dbsfsy,v_5s,sgnr,hbgz,userid,custcode,createtime)"
					+ "values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("GZCL"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("WPHBZ"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("MZSW"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("MZTT"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("TCMZ"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("XZPZMZ"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("XZPZHB"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("FQMZ"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("LCMZ"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("FMMZ"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("TOTALCHZ"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("TOTALCJZ"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("FMSZZST"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("DNMZ"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("DNZZ"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("BYJL"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("ZZYF"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("BYXS"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("BYST"));
			DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("YFJL"));
			DbUtil.setObject(ps, 22, Types.VARCHAR, vo.opt("YFZHB"));
			DbUtil.setObject(ps, 23, Types.VARCHAR, vo.opt("YFXS"));
			DbUtil.setObject(ps, 24, Types.VARCHAR, vo.opt("YFST"));
			DbUtil.setObject(ps, 25, Types.VARCHAR, vo.opt("MZL"));
			DbUtil.setObject(ps, 26, Types.VARCHAR, vo.opt("SB"));
			DbUtil.setObject(ps, 27, Types.VARCHAR, vo.opt("YFL"));
			DbUtil.setObject(ps, 28, Types.VARCHAR, vo.opt("TOTALNUMBER"));
			DbUtil.setObject(ps, 29, Types.VARCHAR, vo.opt("ZHYZYM"));
			DbUtil.setObject(ps, 30, Types.VARCHAR, vo.opt("DBSFSY"));
			DbUtil.setObject(ps, 31, Types.VARCHAR, vo.opt("V_5S"));
			DbUtil.setObject(ps, 32, Types.VARCHAR, vo.opt("SGNR"));
			DbUtil.setObject(ps, 33, Types.VARCHAR, vo.opt("HBGZ"));
			DbUtil.setObject(ps, 34, Types.VARCHAR, vo.opt("USERID"));
			DbUtil.setObject(ps, 35, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps, 36, Types.VARCHAR, vo.opt("CREATETIME"));

			ps.execute();

			// 提交事务
			System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 下载模板
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