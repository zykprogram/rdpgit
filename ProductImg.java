/**
 * 
 */
package com.sbt.tool;

import java.io.File;
import java.io.InputStream;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.json.JSONArray;
import org.json.JSONObject;

import com.webbuilder.utils.DbUtil;
import com.webbuilder.utils.FileUtil;

/**
 * @author Administrator
 *
 */
public class ProductImg {
	
     
	/**
	 * 杨波 
	 * 删除图片
	 * @param request
	 * @param response
	 * @throws Exception
	 */
	public static void delProductImg(HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		JSONArray ja = new JSONArray(request.getParameter("grid1"));
		JSONObject jo;
		int i, j = ja.length();
		String pk = "";
		String path = "";

		for (i = 0; i < j; i++) {
			jo = ja.getJSONObject(i);
			path=jo.getString("PATH");
			pk = jo.getString("IMGURL");
			deleteFile(path+pk);
		}
	
	}
	
	/**
	 * 杨波 
	 * 产品图片增加
	 * @param request
	 * @param response
	 * @throws Exception
	 */
	public static void addProductImg(HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		String fileNameNew = "";// 根据主键定义的文件名

		String sql = request.getAttribute("sql").toString();
		String dir = request.getAttribute("dir").toString();
		fileNameNew = request.getAttribute("fileNameNew").toString();

		File file = new File(dir); // 判断文件夹是否存在,如果不存在则创建文件夹
		if (!file.exists() && !file.isDirectory()) {
			file.mkdirs();
		}
		InputStream stream = (InputStream) request.getAttribute("uploadFile");
		String fileName = request.getAttribute("uploadFile__name").toString();
		String fileName2 = FileUtil.extractFileExt(fileName);
		
		//校验上传的文件格式
		if (fileName2.equals("jpg") || fileName2.equals("JPG") ||fileName2.equals("gif")
				|| fileName2.equals("GIF") || fileName2.equals("png")
				|| fileName2.equals("PNG")) {
			FileUtil.saveStream(stream, new File(dir, fileNameNew));
				update(sql, request);
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

	public static void update(String sq,HttpServletRequest request) throws Exception {
        //StringBuffer sb=new StringBuffer();
		DbUtil.execute(request, sq);

	}
	 /**
     * 删除单个文件
     * @param   sPath    被删除文件的文件名
     * @return 单个文件删除成功返回true，否则返回false
     */
    public static boolean deleteFile(String sPath) {
        boolean flag = false;
        File file = new File(sPath);
        // 路径为文件且不为空则进行删除
        if (file.isFile() && file.exists()) {
            file.delete();
            flag = true;
        }
        return flag;
    }
}
