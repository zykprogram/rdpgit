package com.sbt.tool;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Date;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.json.JSONArray;

import com.webbuilder.common.Main;
import com.webbuilder.common.Str;
import com.webbuilder.utils.DateUtil;
import com.webbuilder.utils.DbUtil;
import com.webbuilder.utils.FileUtil;
import com.webbuilder.utils.StringUtil;
import com.webbuilder.utils.SysUtil;
import com.webbuilder.utils.WebUtil;
import com.webbuilder.utils.ZipUtil;

public class FileUpload {

	/**
	 * 杨波 文件上传功能 存入的地址为绝对地址
	 * 
	 * @param request
	 * @param response
	 * @throws Exception
	 */
	public static void uploadFile(HttpServletRequest request,
			HttpServletResponse response) throws Exception {

		String table_name = request.getAttribute("tabname").toString();
		String table_fjdz = request.getAttribute("fjdz").toString();
		String table_pk = request.getAttribute("tabpk").toString();

		String dir = request.getAttribute("dir").toString();

		File file = new File(dir); // 判断文件夹是否存在,如果不存在则创建文件夹
		if (!file.exists() && !file.isDirectory()) {
			file.mkdirs();
		}
		String dbtype = request.getAttribute("dbtype").toString();
		String xmbhVal = request.getAttribute("xmbh").toString();
		InputStream stream = (InputStream) request.getAttribute("uploadFile");
		String fileName = request.getAttribute("uploadFile__name").toString();
		// String filePath=dir+"\\"+fileName;
		String filePath = dir;

		if (StringUtil.isEqual(request.getAttribute("type").toString(), "1")) {
			if (StringUtil.isSame(FileUtil.extractFileExt(fileName), "zip"))
				ZipUtil.unzip(stream, new File(dir));
			else
				throw new Exception(Str.format(request, "selectZip"));
		} else
			FileUtil.saveStream(stream, new File(dir, fileName));
		if (dbtype.equals("1"))
			update(xmbhVal, filePath, request, dbtype, table_name, table_fjdz,
					table_pk);
	}

	/**
	 * 杨波 文件上传功能 根据客户网要求 上传后在数据库中存入的地址为相对地址
	 * 
	 * @param request
	 * @param response
	 * @throws Exception
	 */
	public static void uploadFileCN(HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		String fileNameNew = "";// 根据主键定义的文件名

		String table_name = request.getAttribute("tabname").toString();
		String table_fjdz = request.getAttribute("fjdz").toString();
		String table_pk = request.getAttribute("tabpk").toString();

		String dir = request.getAttribute("dir").toString();
		String DBdir = request.getAttribute("DBdir").toString();

		
		
		File file = new File(dir); // 判断文件夹是否存在,如果不存在则创建文件夹
		if (!file.exists() && !file.isDirectory()) {
			file.mkdirs();
		}
		String dir2=Main.path+"/img/"+DBdir;
		System.out.print("复制到项目根目录的文件为：" +dir2);
		File file2= new File(Main.path+"/img/"+DBdir);//客户网需要用到
		if (!file2.exists() && !file2.isDirectory()) {
			file2.mkdirs();
		}
		
		String dbtype = request.getAttribute("dbtype").toString();
		String xmbhVal = request.getAttribute("xmbh").toString();
		InputStream stream = (InputStream) request.getAttribute("uploadFile");
		String fileName = request.getAttribute("uploadFile__name").toString();
		String fileName2 = FileUtil.extractFileExt(fileName);
		
		//校验上传的文件格式
		if (fileName2.equals("jpg") ||fileName2.equals("JPG")||fileName2.equals("gif")||fileName2.equals("GIF")
				 || fileName2.equals("png")||fileName2.equals("PNG")
			|| fileName2.toUpperCase().equals("FLV")|| fileName2.toUpperCase().equals("SWF")) {
			if (StringUtil
					.isEqual(request.getAttribute("type").toString(), "1")) {
				if (StringUtil.isSame(FileUtil.extractFileExt(fileName), "zip"))
					ZipUtil.unzip(stream, new File(dir));
				else
					throw new Exception(Str.format(request, "selectZip"));
			} else
//				System.out.print(fileName.substring(fileName.indexOf(".") + 1,
//						fileName.length()));

			fileNameNew = xmbhVal + "."+fileName2;// 修改上传的文件名为主键
			FileUtil.saveStream(stream, new File(dir, fileNameNew));
			File cpfile = new File(dir,fileNameNew);
			File cpnewFile = new File(file2, fileNameNew);
			FileUtil.copyFile(cpfile, cpnewFile, true, false);//由于客户网中新网模块 需要在前台展示图片
			if (dbtype.equals("1"))
				update(xmbhVal, DBdir+"/"+fileNameNew, request, dbtype, table_name, table_fjdz,
						table_pk);
			
		}
		else
		{
			throw new Exception("上传的图片格式不正确");
		}
	}

	/**
	 * 修改项目附件地址到表中 杨波 根据AJAX传递过来的表名和附件地址 动态修改
	 * 
	 * @param request
	 * @param response
	 * @throws Exception
	 */

	public static void update(String xmbhVal, String filePath,
			HttpServletRequest request, String type, String table_name,
			String BZFJDZ, String table_pk) throws Exception {
		int rs3 = 0;
		String xqfjTemp = filePath;

		rs3 = DbUtil.update(request, "update " + table_name + " set  " + BZFJDZ
				+ " = replace('" + xqfjTemp + "','\\\','/') where " + table_pk
				+ " = '" + xmbhVal + "' ");

		if (rs3 > 0) {
			System.out.print("修改项目附件地址到表中成功....");
		} else {
			System.out.print("修改项目附件地址到表中失败....");
		}

	}

	/**
	 * 下载
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
			files[i] = new File(ja.optString(i));
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
	
	/**
	 *  叶志强
	 * @param request
	 * @param response
	 * @throws Exception
	 */
		public static void getFiles(HttpServletRequest request,
				HttpServletResponse response) throws Exception {
			File dir = new File(request.getParameter("dir"));
			File[] fs = dir.listFiles();

			if (fs == null || fs.length == 0) {
				WebUtil.response(response, "[]");
				return;
			}
			boolean isFirst = true;
			StringBuilder buf = new StringBuilder();
			buf.append("[");
			for (File file : fs) {
				if (isFirst)
					isFirst = false;
				else
					buf.append(",");
				loadFileInfo(file, buf);
			}
			buf.append("]");
			WebUtil.response(response, buf);
		}
		
		
		private static void loadFileInfo(File file, StringBuilder buf) {
			boolean isDir = file.isDirectory();
			buf.append("{text:");
			buf.append(StringUtil.quote(file.getName()));
			buf.append(",size:");
			if (isDir)
				buf.append("null");
			else
				buf.append(file.length());
			buf.append(",isDir:");
			buf.append(isDir);
			buf.append(",dir:");
			buf.append(StringUtil.quote(FileUtil.getPath(file)));
			buf.append(",type:");
			if (isDir)
				buf.append("Str.folder");
			else
				buf.append(StringUtil.quote(FileUtil.extractFileExt(file.getName()).toLowerCase()));
			buf.append(",date:\"");
			buf.append(DateUtil.toString(new Date(file.lastModified())));
			buf.append("\"}");
		}
		
		public static void getImage(HttpServletRequest request,
				HttpServletResponse response) throws Exception {
			String fileName = WebUtil.decode(request.getParameter("file"));
			String fileExt = FileUtil.extractFileExt(fileName).toLowerCase();
			String imgTypes[] = { "gif", "jpg", "jpeg", "png", "bmp" };
			File file = null;
			if (StringUtil.indexOf(imgTypes, fileExt) == -1) {
				file = new File(Main.path, "webbuilder/images/delete.gif");
				fileExt="gif";
			} else {
				file = new File(fileName);
			}

			response.reset();
			
			response.setContentType("image/" + fileExt);
			FileInputStream is = new FileInputStream(file);
			try {
				SysUtil.isToOs(is, response.getOutputStream());
			} finally {
				is.close();
			}
			response.flushBuffer();
		}

}
