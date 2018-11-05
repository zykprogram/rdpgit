package com.sbt.tool.hr63;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.Types;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import com.webbuilder.common.Main;
import com.webbuilder.utils.DbUtil;
import com.webbuilder.utils.FileUtil;
import com.webbuilder.utils.StringUtil;
import com.webbuilder.utils.WebUtil;
import com.webbuilder.utils.ZipUtil;
/**
 * EXcel 导入
 * @author Administrator
 *
 */
public class ExcelImportTool {
	public static void getFile(HttpServletRequest request,
			HttpServletResponse response)throws Exception  {

		InputStream in = (InputStream) request.getAttribute("uploadFile");
		String fileName = request.getAttribute("uploadFile__name").toString();
		String fileType = fileName.substring(fileName.lastIndexOf(".") + 1,
				fileName.length());
		String imptype = request.getAttribute("imptype").toString();
		Map<String, String> map = new HashMap<String, String>();
		if ("1".equals(imptype)) { //

			map.put("工号", "USER_CODE");
			map.put("打卡时间", "SJ");
			
		}else if ("2".equals(imptype)) 
		{
			map.put("年度", "CYEAR");
			map.put("月份", "CMONTH");
			map.put("期间", "CPERIOD");
			map.put("日期", "CDAY");
			map.put("系统", "XT");
			map.put("大区", "DQ");
			map.put("组织", "ORGNAME");
			map.put("上周在职", "SZZZ");
			map.put("本周在职", "BZZZ");
			map.put("净增人数", "JZRS");
			map.put("本周入职", "BARZ");
			map.put("本周离职", "BZLZ");
			map.put("本周调入", "BZDR");
			map.put("本周调出", "BZDC");
		}

		read(in, fileType,map,request,response);
		
	} 
	
	
     

	/**
	 * 读取文件
	 * @param in
	 * @param fileType
	 * @throws Exception 
	 */
	public static void read(InputStream in, String fileType,Map<String, String> map,HttpServletRequest request,
			HttpServletResponse response) throws Exception   {
		int dqrow=0;
		try {
			List<JSONObject> list = new ArrayList<JSONObject>();
			//
			if (fileType.equals("xls")) {
				HSSFWorkbook wb = new HSSFWorkbook(in);
				HSSFSheet sheet = wb.getSheetAt(0); 
				HSSFRow headRow = sheet.getRow(0);
				HSSFRow row = null;
				HSSFCell cell = null;
				
				// 第一个报表
				if (sheet != null && sheet.getLastRowNum() > 0) {
					for (int i = 1; i <= sheet.getLastRowNum(); i++) {
						dqrow=i;
						JSONObject jsonObject = new JSONObject();
						// 得到当前工作表的行数
						row = sheet.getRow(i);
						
						if (null != row) {
							for (int j = 0; j < row.getLastCellNum(); j++) {
								if (j>headRow.getLastCellNum()-1) {
									break;
								}
								cell = row.getCell(j);
								if (headRow.getCell(j)!=null&&headRow.getCell(j).getStringCellValue().contains("打卡时间"))
								{
										
								   Date cellVal =cell.getDateCellValue();
								   jsonObject.put(map.get(headRow.getCell(j).getStringCellValue().toString()), cellVal);
							    }
								else
								{
									Object   cellVal = cell.getStringCellValue();
									jsonObject.put(map.get(headRow.getCell(j).getStringCellValue().toString()), cellVal);
								}
								
								
							}
							list.add(jsonObject);
						}
					}
				}
			} else if (fileType.equals("xlsx")) {//考虑到2007 性能问题 暂时屏蔽
			
				XSSFWorkbook wb1 = new XSSFWorkbook(in);
				XSSFSheet sheet1 = wb1.getSheetAt(0); 
				XSSFRow headRow =  sheet1.getRow(0);
				XSSFRow row = null;
				XSSFCell cell = null;
				
				sheet1 = wb1.getSheetAt(0);
				if (sheet1 != null && sheet1.getLastRowNum() > 0) {
					for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
						dqrow=i;
						JSONObject jsonObject = new JSONObject();
						// 得到当前工作表的行数
						row = sheet1.getRow(i);
						
						if (null != row) {
							for (int j = 0; j < row.getLastCellNum(); j++) {
								if (j>headRow.getLastCellNum()-1) {
									break;
								}
								cell = row.getCell(j);
								if (headRow.getCell(j)!=null&&headRow.getCell(j).getStringCellValue().contains("打卡时间"))
								{
										
								   Date cellVal =cell.getDateCellValue();
								   jsonObject.put(map.get(headRow.getCell(j).getStringCellValue().toString()), cellVal);
							    }
								else
								{
									Object   cellVal = cell.getStringCellValue();
									jsonObject.put(map.get(headRow.getCell(j).getStringCellValue().toString()), cellVal);
								}
								
								
							}
							list.add(jsonObject);
						}
					}
				}

			}
			 else {
				throw new Exception("只支持2003版本的Excel导入！");
			}
			String imptype = request.getAttribute("imptype").toString();
//			for (int f = 0; f < list.size(); f++) {
//				if ("1".equals(imptype)) {
//					imp_scfwwh(list.get(f),request,response);
//				}
//				
//			}
			if ("1".equals(imptype)) {
				imp_scfwwh(list,request,response);
			}
			if ("2".equals(imptype)) {
				imp_hrrysl(list,request,response);
			}
			in.close();
		} catch (Exception e){
			dqrow=dqrow+1;
			throw e;
		
		}
	}
	
	/**
	 *  将HR人员数量情况导入到主平台中
	 *  boyang
	 * @param vo
	 * @throws Exception 
	 * @throws Exception 
	 * @throws Exception
	 */
	public static void imp_hrrysl(List<JSONObject> vo,HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		String sql = "";
		if (null != vo) {
			Connection conn = DbUtil.getConnection();
			conn.setAutoCommit(false);
			//DbUtil.startTrans(conn, "");
			PreparedStatement ps= null;

			//插入
			sql = "INSERT INTO wb_erp.HRRYSL (CYEAR,CMONTH,CPERIOD,CDAY,XT,DQ,ORGNAME,SZZZ,BZZZ,JZRS,BZRZ,BZLZ,BZDR,BZDC)"+
					"VALUES "+
					 " (?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			for (int f = 0; f < vo.size(); f++) {
				DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get(f).get("CYEAR"));
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get(f).opt("CMONTH"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.get(f).opt("CPERIOD"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.get(f).opt("CDAY"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.get(f).opt("XT"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.get(f).opt("DQ"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.get(f).opt("ORGNAME"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.get(f).opt("SZZZ"));
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo.get(f).opt("BZZZ"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.get(f).opt("JZRS"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.get(f).opt("BARZ"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.get(f).opt("BZLZ"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.get(f).opt("BZDR"));
				DbUtil.setObject(ps, 14, Types.VARCHAR, vo.get(f).opt("BZDC"));
				ps.addBatch();
			}
			ps.executeBatch();
			//result=ps.executeUpdate();
			//提交事务
			conn.commit();
		
			//关闭资源
			//DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);
			DbUtil.closeConnection(conn);
			
		}
	}
	
	
	
	/**
	 *  导入钉钉数据到HR63 中间表
	 *  yezq
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_scfwwh(List<JSONObject>  vo,HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		String PK_ID = null;
		String sql = "";
		if (null != vo) {
			Connection conn = DbUtil.getConnection("jdbc/wb_mssql2");
			conn.setAutoCommit(false);
			//DbUtil.startTrans(conn, "");
			PreparedStatement ps = null;
		
			//插入
			sql = "INSERT INTO DD_KQ (USER_CODE,SJ)"+
					"VALUES "+
					 " (?, ?)";
			ps = conn.prepareStatement(sql);
			for (int f = 0; f < vo.size(); f++) {
				DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get(f).get("USER_CODE"));
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get(f).opt("SJ"));
				ps.addBatch();
			}
			ps.executeBatch();
			//提交事务
			conn.commit();
			
			//关闭资源
			//DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);
			DbUtil.closeConnection(conn);
		}

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
			files[i] = new File(Main.path,ja.optString(i));
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


