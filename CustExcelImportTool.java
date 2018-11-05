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
 * EXcel ����
 * ��Ҫ���Ӫ������ƽ̨ҵ��Ա��ͻ���ϵ����
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
	 * ��ȡ�ļ�
	 * @param in
	 * @param fileType
	 * @throws Exception 
	 */
	public static void read(InputStream in, String fileType,HttpServletRequest request,
			HttpServletResponse response) throws Exception   {
		System.out.println("����");
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
				// ��һ������
				sheet1 = wb.getSheetAt(0);
				if (sheet1 != null && sheet1.getLastRowNum() > 0) {
					for (int i = 2; i <= sheet1.getLastRowNum(); i++) {
						dqrow=i;
						System.out.println(i+"========"+i);
						SO_CUSTMANVo vo = new SO_CUSTMANVo();
						// �õ���ǰ�����������
						row = sheet1.getRow(i);
						if (null != row) {
							cell=row.getCell(0);
							if (null == cell.getStringCellValue()||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е����ݰ汾����Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								version=cell.getStringCellValue();
								vo.setDATEVERSION(cell.getStringCellValue());
							}
							typecell = row.getCell(1);
							if (null == typecell.getStringCellValue()||""==typecell.getStringCellValue()) {
								throw new Exception("��"+i+"�еĿͻ���Ų���Ϊ��!");
							}else
							{
								typecell.setCellType(HSSFCell.CELL_TYPE_STRING);
								custcode = typecell.getStringCellValue();
								vo.setCUSTCODE(custcode);
							}
							cell = row.getCell(2);
							if (null == cell.getStringCellValue()||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�еĿͻ���������Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTNAME(cell.getStringCellValue());

							}
							
							cell = row.getCell(3);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�������������Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_5(cell.getStringCellValue());

							}

							cell = row.getCell(4);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�����ʡ����Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_4(cell.getStringCellValue());

							}
								
							cell = row.getCell(5);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е������в���Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_3(cell.getStringCellValue());

							}
							
							cell = row.getCell(6);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е������ز���Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_2(cell.getStringCellValue());

							}

							cell = row.getCell(7);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е������粻��Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_1(cell.getStringCellValue());
							}
							cell = row.getCell(8);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�ҵ��Ա��Ų���Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setPSNCODE(cell.getStringCellValue());
							}
									
							cell = row.getCell(9);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�ҵ��Ա��������Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setPSNNAME(cell.getStringCellValue());
							}
							
							cell = row.getCell(10);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�ҵ��Ա��λ����Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setJOBNAME(cell.getStringCellValue());
							}
							cell = row.getCell(11);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�Ӫ����������Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME_5(cell.getStringCellValue());
							}
							cell = row.getCell(12);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�Ӫ��ʡ������Ϊ��!");
								
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME_4(cell.getStringCellValue());
							}
							
							cell = row.getCell(13);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�Ӫ��������Ϊ��!");
								
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME_3(cell.getStringCellValue());
							}
							
							cell = row.getCell(14);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�Ӫ��С������Ϊ��!");
								
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME_2(cell.getStringCellValue());
							}
							
							cell = row.getCell(15);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�Ӫ��Ƭ�鲻��Ϊ��!");
								
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
						// �õ���ǰ�����������
						dqrow=i;
						row = sheet1.getRow(i);
						if (null != row) {
							cell=row.getCell(0);
							if (null == cell.getStringCellValue()||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е����ݰ汾����Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								version=cell.getStringCellValue();
								vo.setDATEVERSION(cell.getStringCellValue());

							}
							typecell = row.getCell(1);
							if (null == typecell.getStringCellValue()||""==typecell.getStringCellValue()) {
								throw new Exception("��"+i+"�еĿͻ���Ų���Ϊ��!");
							}else
							{
								typecell.setCellType(HSSFCell.CELL_TYPE_STRING);
								custcode = typecell.getStringCellValue();
								vo.setCUSTCODE(custcode);
							}
							cell = row.getCell(2);
							if (null == cell.getStringCellValue()||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�еĿͻ���������Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTNAME(cell.getStringCellValue());

							}
							
							cell = row.getCell(3);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�������������Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_5(cell.getStringCellValue());

							}

							cell = row.getCell(4);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�����ʡ����Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_4(cell.getStringCellValue());

							}
								
							cell = row.getCell(5);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е������в���Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_3(cell.getStringCellValue());

							}
							
							cell = row.getCell(6);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е������ز���Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_2(cell.getStringCellValue());

							}
							
							cell = row.getCell(7);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е������粻��Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setCUSTAREA_1(cell.getStringCellValue());
							}
							cell = row.getCell(8);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�ҵ��Ա��Ų���Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setPSNCODE(cell.getStringCellValue());
							}
									
							cell = row.getCell(9);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�ҵ��Ա��������Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setPSNNAME(cell.getStringCellValue());
							}
							
							cell = row.getCell(10);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�ҵ��Ա��λ����Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setJOBNAME(cell.getStringCellValue());
							}
							cell = row.getCell(11);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�Ӫ����������Ϊ��!");
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME_5(cell.getStringCellValue());
							}
							cell = row.getCell(12);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�Ӫ��ʡ������Ϊ��!");
								
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME_4(cell.getStringCellValue());
							}
							
							cell = row.getCell(13);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�Ӫ��������Ϊ��!");
								
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME_3(cell.getStringCellValue());
							}
							
							cell = row.getCell(14);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�Ӫ��С������Ϊ��!");
								
							}else
							{
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								vo.setVSALESTRUNAME_2(cell.getStringCellValue());
							}
							
							cell = row.getCell(15);
							if (null == cell||""==cell.getStringCellValue()) {
								throw new Exception("��"+i+"�е�Ӫ��Ƭ�鲻��Ϊ��!");
								
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
				throw new Exception("�������excel��ʽ����ȷ");
				
				//System.out.println("�������excel��ʽ����ȷ");
			}
			deleteSo_CUST(version);
			updateSo_CUST();
			for (int f = 0; f < zzlist.size(); f++) {
				impSo_CustSalesMan(zzlist.get(f));
			}
			in.close();
		} catch (Exception e){
			dqrow=dqrow+1;
			throw new Exception("��"+dqrow+"�����ݸ�ʽ����ȷ�����޸�");
			//System.out.println("������"+dqrow);
		//	e.printStackTrace();
		}
		
		//LogUtil.message("�������ݳɹ�");
	}

	/**
	 *  ���������Ŀͻ���ϵ
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
				ps.setInt(3, 1);//1�����϶�Ӧ��ϵ2�������϶�Ӧ��ϵ
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
				ps.setString(19, "Y");//���°汾
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
	2. * ���ָ����Χ��N�����ظ�����  
	3. * �ڳ�ʼ�������ظ���ѡ�������������һ�����������У�  
	4. * ����ѡ���鱻������������ô�ѡ����(len-1)�±��Ӧ�����滻  
	5. * Ȼ���len-2�����������һ����������������  
	6. * @param max  ָ����Χ���ֵ  
	7. * @param min  ָ����Χ��Сֵ  
	8. * @param n  ���������  
	9. * @return int[] ����������  
	10. */  
	 public static String  getId() {
		 // ���� [0-n) �����ظ��������  // list ����������Щ����� 
		 String st="0000000000";
		 StringBuffer list =new StringBuffer();
		 int n = 10;  
		 Random rand = new Random(); 
		 boolean[] bool = new boolean[n]; 
		 int num = 0;  for (int i = 0; i < n; i++) {  
			 do {    
				 // �������������ͬ����ѭ�� 
				 num = rand.nextInt(n);  
				 } while (bool[num]); 
			 bool[num] = true; 
			 list.append(num); 
			 }  
		 st=list.toString();
		return st;
	 }
	
	/**
	 *  ������ʷ�汾
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
	 *  ɾ�����쵼��İ汾����
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
	 * yangbo ��ȡ��ǰ����
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
	 * yangbo ��ȡ��ǰ����
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
