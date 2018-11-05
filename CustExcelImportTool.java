package com.sbt.tool;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

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

import com.sbt.tool.vo.SO_CUSTMANVo;
import com.webbuilder.utils.DbUtil;

import java.io.InputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
/**
 * EXcel 导入
 * 主要针对营销报表平台业务员与客户关系导入
 * @author Administrator
 *
 */
public class CustExcelImportTool {
	public static void getFile(HttpServletRequest request,
			HttpServletResponse response)throws Exception  {

		InputStream in = (InputStream) request.getAttribute("uploadFile");
		String fileName = request.getAttribute("uploadFile__name").toString();
		String fileType = fileName.substring(fileName.lastIndexOf(".") + 1,
				fileName.length());
		read(in, fileType,request,response);
		// updateWbuser(null);
	} 

	/**
	 * 读取文件
	 * @param in
	 * @param fileType
	 * @throws Exception 
	 */
	public static void read(InputStream in, String fileType,HttpServletRequest request,
			HttpServletResponse response) throws Exception   {
		System.out.println("进入");
		int dqrow=0;
		String version="";
		try {
			
		
		List<SO_CUSTMANVo> zzlist = new ArrayList<SO_CUSTMANVo>();
			if (fileType.equals("xls")) {
				HSSFWorkbook wb = new HSSFWorkbook(in);
				HSSFSheet sheet1 = null;
				HSSFRow row = null;
				HSSFCell cell = null;
				HSSFCell typecell = null;
				String custcode;

				// /////////////////////////////////////////////////
				// 第一个报表
				sheet1 = wb.getSheetAt(0);
				if (sheet1 != null && sheet1.getLastRowNum() > 0) {
					for (int i = 2; i <= sheet1.getLastRowNum(); i++) {
						dqrow=i;
						System.out.println(i+"========"+i);
						SO_CUSTMANVo vo = new SO_CUSTMANVo();
						// 得到当前工作表的行数
						row = sheet1.getRow(i);
						if (null != row) {
							cell=row.getCell(0);
							if (null == cell.getStringCellValue()||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的数据版本不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								version=cell.getStringCellValue();
								vo.setDATEVERSION(cell.getStringCellValue());
							}
							typecell = row.getCell(1);
							if (null == typecell.getStringCellValue()||""==typecell.getStringCellValue()) {
								throw new Exception("第"+i+"行的客户编号不能为空!");
							}else
							{
								typecell.setCellType(HSSFCell.CELL_TYPE_STRING);
								custcode = typecell.getStringCellValue();
								vo.setCUSTCODE(custcode);
							}
							cell = row.getCell(2);
							if (null == cell.getStringCellValue()||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的客户姓名不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTNAME(cell.getStringCellValue());

							}
							
							cell = row.getCell(3);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的行政大区不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_5(cell.getStringCellValue());

							}

							cell = row.getCell(4);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的行政省不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_4(cell.getStringCellValue());

							}
								
							cell = row.getCell(5);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的行政市不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_3(cell.getStringCellValue());

							}
							
							cell = row.getCell(6);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的行政县不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_2(cell.getStringCellValue());

							}

							cell = row.getCell(7);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的行政乡不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_1(cell.getStringCellValue());
							}
							cell = row.getCell(8);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的业务员编号不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setPSNCODE(cell.getStringCellValue());
							}
									
							cell = row.getCell(9);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的业务员姓名不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setPSNNAME(cell.getStringCellValue());
							}
							
							cell = row.getCell(10);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的业务员岗位不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setJOBNAME(cell.getStringCellValue());
							}
							cell = row.getCell(11);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的营销大区不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME_5(cell.getStringCellValue());
							}
							cell = row.getCell(12);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的营销省区不能为空!");
								
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME_4(cell.getStringCellValue());
							}
							
							cell = row.getCell(13);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的营销部不能为空!");
								
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME_3(cell.getStringCellValue());
							}
							
							cell = row.getCell(14);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的营销小区不能为空!");
								
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME_2(cell.getStringCellValue());
							}
							
							cell = row.getCell(15);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的营销片组不能为空!");
								
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME(cell.getStringCellValue());
							}
							
							zzlist.add(vo);

						}
					}
				}

				// ///////////////////////////////////////////////////////////////////////

			} else if (fileType.equals("xlsx")) {
				System.out.println("in");
				XSSFWorkbook wb1 = new  XSSFWorkbook(in);
				XSSFSheet sheet1 = null;
				XSSFRow row = null;
				XSSFCell cell = null;
				XSSFCell typecell = null;
				String custcode = "";
			//System.out.println("ffff");
				sheet1 = wb1.getSheetAt(0);
				if (sheet1 != null && sheet1.getLastRowNum() > 0) {
					for (int i = 2; i <= sheet1.getLastRowNum(); i++) {
						SO_CUSTMANVo vo = new SO_CUSTMANVo();
						System.out.println(i+"========"+i);
						// 得到当前工作表的行数
						dqrow=i;
						row = sheet1.getRow(i);
						if (null != row) {
							cell=row.getCell(0);
							if (null == cell.getStringCellValue()||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的数据版本不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								version=cell.getStringCellValue();
								vo.setDATEVERSION(cell.getStringCellValue());

							}
							typecell = row.getCell(1);
							if (null == typecell.getStringCellValue()||""==typecell.getStringCellValue()) {
								throw new Exception("第"+i+"行的客户编号不能为空!");
							}else
							{
								typecell.setCellType(HSSFCell.CELL_TYPE_STRING);
								custcode = typecell.getStringCellValue();
								vo.setCUSTCODE(custcode);
							}
							cell = row.getCell(2);
							if (null == cell.getStringCellValue()||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的客户姓名不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTNAME(cell.getStringCellValue());

							}
							
							cell = row.getCell(3);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的行政大区不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_5(cell.getStringCellValue());

							}

							cell = row.getCell(4);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的行政省不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_4(cell.getStringCellValue());

							}
								
							cell = row.getCell(5);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的行政市不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_3(cell.getStringCellValue());

							}
							
							cell = row.getCell(6);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的行政县不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_2(cell.getStringCellValue());

							}
							
							cell = row.getCell(7);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的行政乡不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_1(cell.getStringCellValue());
							}
							cell = row.getCell(8);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的业务员编号不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setPSNCODE(cell.getStringCellValue());
							}
									
							cell = row.getCell(9);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的业务员姓名不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setPSNNAME(cell.getStringCellValue());
							}
							
							cell = row.getCell(10);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的业务员岗位不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setJOBNAME(cell.getStringCellValue());
							}
							cell = row.getCell(11);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的营销大区不能为空!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME_5(cell.getStringCellValue());
							}
							cell = row.getCell(12);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的营销省区不能为空!");
								
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME_4(cell.getStringCellValue());
							}
							
							cell = row.getCell(13);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的营销部不能为空!");
								
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME_3(cell.getStringCellValue());
							}
							
							cell = row.getCell(14);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的营销小区不能为空!");
								
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME_2(cell.getStringCellValue());
							}
							
							cell = row.getCell(15);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("第"+i+"行的营销片组不能为空!");
								
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME(cell.getStringCellValue());
							}
							
							zzlist.add(vo);

						}
					}
				}
			 
			} else {
				throw new Exception("您输入的excel格式不正确");
				
				//System.out.println("您输入的excel格式不正确");
			}
			deleteSo_CUST(version);
			updateSo_CUST();
			for (int f = 0; f < zzlist.size(); f++) {
				impSo_CustSalesMan(zzlist.get(f));
			}
			in.close();
		} catch (Exception e){
			dqrow=dqrow+1;
			throw new Exception("第"+dqrow+"行数据格式不正确，请修改");
			//System.out.println("报错了"+dqrow);
		//	e.printStackTrace();
		}
		
		//LogUtil.message("导入数据成功");
	}

	/**
	 *  导入新增的客户关系
	 *  yangbo
	 * @param vo
	 * @throws Exception
	 */
	public static void impSo_CustSalesMan(SO_CUSTMANVo vo) throws Exception {
		SO_CUSTMANVo zvo = vo;
		String sql = "";
		if (null != zvo) {
			Connection conn = DbUtil.getConnection("java:comp/env/jdbc/wb_dc41");
			PreparedStatement ps = null;
			  
				sql = "insert into zsj_so_custmoerrelatsalesman (PK_DZ, DATEVERSION, RELATIONTYPE, CUSTCODE, CUSTNAME," +
						"CUSTAREA_5,CUSTAREA_4,CUSTAREA_3,CUSTAREA_2,CUSTAREA_1,PSNCODE,PSNNAME,JOBNAME,VSALESTRUNAME_5," +
						"VSALESTRUNAME_4,VSALESTRUNAME_3,VSALESTRUNAME_2,VSALESTRUNAME,ISRECENTLY,COPERATORID,DBILLDATE,FSTATUS," +
						"DR,TS) "
						+ " values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
				ps = conn.prepareStatement(sql);
				ps.setString(1, "SO"+getId());
				ps.setString(2,zvo.getDATEVERSION());
				ps.setInt(3, 1);//1、猪料对应关系2、禽鱼料对应关系
			    ps.setString(4, zvo.getCUSTCODE());
				ps.setString(5, zvo.getCUSTNAME());
				ps.setString(6, zvo.getCUSTAREA_5());
				ps.setString(7, zvo.getCUSTAREA_4());
				ps.setString(8, zvo.getCUSTAREA_3());
				ps.setString(9, zvo.getCUSTAREA_2());
				ps.setString(10,zvo.getCUSTAREA_1());
				ps.setString(11, zvo.getPSNCODE());
				ps.setString(12, zvo.getPSNNAME());
				ps.setString(13, zvo.getJOBNAME());
				ps.setString(14, zvo.getVSALESTRUNAME_5());
				ps.setString(15, zvo.getVSALESTRUNAME_4());
				ps.setString(16, zvo.getVSALESTRUNAME_3());
				ps.setString(17, zvo.getVSALESTRUNAME_2());
				ps.setString(18, zvo.getVSALESTRUNAME());
				ps.setString(19, "Y");//最新版本
				ps.setString(20, "imp");
				ps.setString(21, "");
				ps.setInt(22, 1);
				ps.setInt(23, 0);
				ps.setString(24,GetNowDate());
				ps.execute();
				conn.commit();

		DbUtil.closeConnection(conn);
		}

	}
	
	/**  
	2. * 随机指定范围内N个不重复的数  
	3. * 在初始化的无重复待选数组中随机产生一个数放入结果中，  
	4. * 将待选数组被随机到的数，用待选数组(len-1)下标对应的数替换  
	5. * 然后从len-2里随机产生下一个随机数，如此类推  
	6. * @param max  指定范围最大值  
	7. * @param min  指定范围最小值  
	8. * @param n  随机数个数  
	9. * @return int[] 随机数结果集  
	10. */  
	 public static String  getId() {
		 // 生成 [0-n) 个不重复的随机数  // list 用来保存这些随机数 
		 String st="0000000000";
		 StringBuffer list =new StringBuffer();
		 int n = 10;  
		 Random rand = new Random(); 
		 boolean[] bool = new boolean[n]; 
		 int num = 0;  for (int i = 0; i < n; i++) {  
			 do {    
				 // 如果产生的数相同继续循环 
				 num = rand.nextInt(n);  
				 } while (bool[num]); 
			 bool[num] = true; 
			 list.append(num); 
			 }  
		 st=list.toString();
		return st;
	 }
	
	/**
	 *  更新历史版本
	 *  yangbo
	 * @param vo
	 * @throws Exception
	 */
	public static void updateSo_CUST() throws Exception {
		String sql2 = "";
			Connection conn = DbUtil.getConnection("java:comp/env/jdbc/wb_dc41");
			PreparedStatement ps2 = null;
			  sql2 ="update zsj_so_custmoerrelatsalesman set ISRECENTLY='N' where ISRECENTLY='Y'";
			  ps2 = conn.prepareStatement(sql2);
			  ps2.execute();
			  conn.commit();
			DbUtil.closeConnection(conn);

	}
	/**
	 *  删除当天导入的版本数据
	 *  yangbo
	 * @param vo
	 * @throws Exception
	 */
	public static void deleteSo_CUST(String version) throws Exception {
		String sql2 = "";
			Connection conn = DbUtil.getConnection("java:comp/env/jdbc/wb_dc41");
			PreparedStatement ps2 = null;
			  sql2 ="delete zsj_so_custmoerrelatsalesman where substr(DATEVERSION,0,10) ='"+version+"' ";
			  ps2 = conn.prepareStatement(sql2);
			  ps2.execute();
			  conn.commit();
			  DbUtil.closeConnection(conn);

	}
	/**
	 * yangbo 获取当前日期
	 * @return
	 */
	public static String  GetNowDate() {
		    String temp_str="";   
		    Date dt = new Date();   
		    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");   
		    temp_str=sdf.format(dt);   
		    return temp_str;  
	}
	
	/**
	 * yangbo 获取当前日期
	 * @return
	 */
	public static String  GetNowDate2() {
		    String temp_str="";   
		    Date dt = new Date();   
		    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");   
		    temp_str=sdf.format(dt);   
		    return temp_str;  
	}
}
