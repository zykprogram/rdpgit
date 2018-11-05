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

public class ExcelImportDingTalk {
	public static void getFile(HttpServletRequest request,
			HttpServletResponse response) throws Exception {

		InputStream in = (InputStream) request.getAttribute("uploadFile");
		String fileName = request.getAttribute("uploadFile__name").toString();
		String fileType = fileName.substring(fileName.lastIndexOf(".") + 1,
				fileName.length());
		String imptype = request.getAttribute("imptype").toString();
		Map<String, String> map = new HashMap<String, String>();
		if ("1".equals(imptype)) { // 客户沟通平台消息推送内容

			map.put("客户编码", "USERID");
			map.put("消息内容", "MESSAGE");
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
								 * Exception("第" + i + "行" + j + "列为空，请填写"); }
								 */
								
//							  System.out.print(cellVal);
//								  System.out.println(i + "行");
//								  System.out.println(map.get(headRow.getCell(j)
//								  .getStringCellValue().toString().trim()
//								 .replaceAll("\r|\n", "")));
								
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

			if ("1".equals(imptype)) {
				IMP_sbt_dingtalk_postmessage(list, request, response);
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
	
	
	
	/**
	 * 片组奖励分配比例
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void IMP_sbt_dingtalk_postmessage(List<JSONObject> voList,
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

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			sql = "insert into wb_erp.sbt_dingtalk_postmessage (ID,USERID,	MESSAGE,CREATETIME) "
					+ " values (?,?,?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'))";

			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("USERID"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("MESSAGE"));
			ps.execute();
			// 提交事务
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
