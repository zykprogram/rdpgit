package com.sbt.tool.zxcrm;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.Types;
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

			map.put("��Ա����", "USERCODE");
			map.put("����", "USERNAME");
			map.put("ʡ", "PROVINCE");
			map.put("��", "CITY");
			map.put("��/��", "AREA");
			map.put("��Ʒ��", "PRODUCTLINE");
			map.put("����������������", "SALESTRID");
			map.put("��ע", "MEMO");
		}
		read(in, fileType,map,request,response);
		
	} 
	
	
     

	/**
	 * ��ȡ�ļ�
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
				Object cellVal = null;
				// ��һ������
				if (sheet != null && sheet.getLastRowNum() > 0) {
					for (int i = 1; i <= sheet.getLastRowNum(); i++) {
						dqrow=i;
						JSONObject jsonObject = new JSONObject();
						// �õ���ǰ�����������
						row = sheet.getRow(i);
						
						if (null != row) {
							for (int j = 0; j < row.getLastCellNum(); j++) {
								if (j>headRow.getLastCellNum()-1) {
									break;
								}
								cell = row.getCell(j);
								if (headRow.getCell(j)!=null&&headRow.getCell(j).getStringCellValue().contains("����")&&HSSFDateUtil.isCellDateFormatted(cell)) {
									cellVal = cell.getDateCellValue();
								}else {
									cellVal = ExcelObject.getCellValue(cell);
								}
								jsonObject.put(map.get(headRow.getCell(j).getStringCellValue().toString()), cellVal);
							}
							list.add(jsonObject);
						}
					}
				}

			} else {
				throw new Exception("ֻ֧��2003�汾��Excel���룡");
			}
			String imptype = request.getAttribute("imptype").toString();
			for (int f = 0; f < list.size(); f++) {
				if ("1".equals(imptype)) {
					imp_scfwwh(list.get(f),request,response);
				}
			}
			in.close();
		} catch (Exception e){
			dqrow=dqrow+1;
			throw e;
		
		}
	}
	
	
	
	
	
	/**
	 *  �����г���Χά��
	 *  yezq
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_scfwwh(JSONObject vo,HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		String PK_ID = null;
		String sql = "";
		int result=0;
		if (null != vo) {
			Connection conn = DbUtil.getConnection();
			DbUtil.startTrans(conn, "");
			PreparedStatement ps = null;
			//PreparedStatement ps1 = null;
			
			//��ɾ��
//			sql= "DELETE wb_erp.APP_WB_MARKET_ZX WHERE USERCODE = ?";
//			ps1 = conn.prepareStatement(sql);
//			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("USERCODE"));
//			result=ps1.executeUpdate();
			
			//����
			sql = "INSERT INTO wb_erp.APP_WB_MARKET_ZX "+
					  "(PK_ID, USERCODE, USERNAME, PROVINCE, CITY, AREA, PRODUCTLINE) "+
					"VALUES "+
					 " (?, ?, ?, ?, ?, ?, ?)";
			ps = conn.prepareStatement(sql);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("USERCODE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("USERNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("PROVINCE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CITY"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("AREA"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PRODUCTLINE"));
			result=ps.executeUpdate();
			
			
			
			//�ύ����
			conn.commit();
			
			//�ر���Դ
			//DbUtil.closeStatement(ps1);
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


