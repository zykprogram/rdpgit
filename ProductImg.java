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
	 * � 
	 * ɾ��ͼƬ
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
	 * � 
	 * ��ƷͼƬ����
	 * @param request
	 * @param response
	 * @throws Exception
	 */
	public static void addProductImg(HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		String fileNameNew = "";// ��������������ļ���

		String sql = request.getAttribute("sql").toString();
		String dir = request.getAttribute("dir").toString();
		fileNameNew = request.getAttribute("fileNameNew").toString();

		File file = new File(dir); // �ж��ļ����Ƿ����,����������򴴽��ļ���
		if (!file.exists() && !file.isDirectory()) {
			file.mkdirs();
		}
		InputStream stream = (InputStream) request.getAttribute("uploadFile");
		String fileName = request.getAttribute("uploadFile__name").toString();
		String fileName2 = FileUtil.extractFileExt(fileName);
		
		//У���ϴ����ļ���ʽ
		if (fileName2.equals("jpg") || fileName2.equals("JPG") ||fileName2.equals("gif")
				|| fileName2.equals("GIF") || fileName2.equals("png")
				|| fileName2.equals("PNG")) {
			FileUtil.saveStream(stream, new File(dir, fileNameNew));
				update(sql, request);
		}
		else
		{
			throw new Exception("�ϴ���ͼƬ��ʽ����ȷ");
		}
	}
	
	/**
	 * �޸���Ŀ������ַ������ � ����AJAX���ݹ����ı����͸�����ַ ��̬�޸�
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
     * ɾ�������ļ�
     * @param   sPath    ��ɾ���ļ����ļ���
     * @return �����ļ�ɾ���ɹ�����true�����򷵻�false
     */
    public static boolean deleteFile(String sPath) {
        boolean flag = false;
        File file = new File(sPath);
        // ·��Ϊ�ļ��Ҳ�Ϊ�������ɾ��
        if (file.isFile() && file.exists()) {
            file.delete();
            flag = true;
        }
        return flag;
    }
}
