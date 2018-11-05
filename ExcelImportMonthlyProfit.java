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
 * EXcel 导入
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
		if ("1".equals(imptype)) { //NC新工厂导入
			map.put("新工厂名称", "NEWUNITNAME");
			map.put("取数工厂名称", "QSUNITNAME");
			map.put("优先级", "LEVELS");
			map.put("日期", "CREATETIME");
		}
		else if("2".equals(imptype)){//预混料调整
			map.put("公司名称","UNITNAME");
			map.put("产品编码","INVCODE");
			map.put("产品名称","INVNAME");
			map.put("吨调整","PRICE");
			map.put("日期", "CREATETIME");
		}
		else if("3".equals(imptype)) {//鱼粉调整
			map.put("公司名称", "UNITNAME");
			map.put("配方编码", "INVCODE");
			map.put("配方名称", "INVNAME");
			map.put("吨调整", "PRICE");
			map.put("日期", "CREATETIME");
		}
		else if("4".equals(imptype)) {//特殊成本
			map.put("调整工厂名称", "UNITNAME");
			map.put("调整事项", "EVENT");
			map.put("成本对象编码", "INVCODE_PF");
			map.put("成本对象", "INVNAME_PF");
			map.put("存货编码", "INVCODE");
			map.put("存货名称", "INVNAME");
			map.put("金额", "PRICE");
			map.put("日期", "CREATETIME");
		}
		else if("5".equals(imptype)) {//关联成本调整
			map.put("公司", "UNITNAME");
			map.put("销售分类","SaleType");
			map.put("存货编码","INVCODE");
			map.put("存货名称","INVNAME");
			map.put("属性", "PROPERTY");
			map.put("关联购买方成品成本调整","PRICE");
			map.put("关联利润分摊","FT");
			map.put("日期", "CREATETIME");
		}
		else if("6".equals(imptype)){//采购KCP利润分配
			map.put("公司", "UNITNAME");
			map.put("存货编码", "INVCODE");
			map.put("存货名称", "INVNAME");
			map.put("玉米KCP分配（吨调整）", "YMKCP");
			map.put("豆粕KCP分配（吨调整）", "DPKCP");
			map.put("收储资金利息分配（吨调整）", "CCZJ");
			map.put("进出口KCP分配（吨调整）", "JCKKCP");
			map.put("日期", "CREATETIME");
		}
		else if("7".equals(imptype)) {//存货计息
			map.put("公司", "UNITNAME");
			map.put("吨计提", "JT");
			map.put("日期", "CREATETIME");
		}
		else if("8".equals(imptype)) {//制造性折旧分摊
			map.put("公司", "UNITNAME");
			map.put("分摊金额", "PRICE");
			map.put("日期", "CREATETIME");
		}
		else if("9".equals(imptype)) {//固定资产计提
			map.put("战区", "VSALESTRUNAME");
			map.put("吨计提", "JT");
			map.put("日期", "CREATETIME");
		}
		else if("10".equals(imptype)) {//其他分摊
			map.put("战区", "VSALESTRUNAME");
			map.put("公司编码", "UNITCODE");
			map.put("公司", "UNITNAME");
			map.put("金额", "PRICE");
			map.put("日期", "CREATETIME");
		}
		//应剔除NC制造成本
		else if("11".equals(imptype)) {
			map.put("公司", "UNITNAME");
			map.put("存货编码","INVCODE");
			map.put("存货名称","INVNAME");
			map.put("属性", "PROPERTY");
			map.put("应剔除NC变动制造成本", "PRICE1");
			map.put("应剔除NC固定制造成本（原始）", "PRICE2");
			map.put("应剔除NC制造成本合计（原始）", "PRICE3");
			map.put("变动成本差额调整","BDCB");
			map.put("固定成本差额调整","GDCB");
			map.put("日期", "CREATETIME");
		}
		else if("12".equals(imptype)) {
			map.put("公司", "UNITNAME");
			map.put("存货编码","INVCODE");
			map.put("存货名称","INVNAME");
			map.put("吨成本", "PRICE");
			map.put("日期", "CREATETIME");
		}
		else if("13".equals(imptype)) {//新工厂及收购工厂
			map.put("公司", "UNITNAME");
			map.put("存货编码", "INVCODE");
			map.put("存货名称", "INVNAME");
			map.put("吨成本", "PRICE");
			map.put("取数工厂名称", "FETCHUNITNAME");
			map.put("优先级", "LEVELS");
			map.put("原材料", "PRICE1");
			map.put("变动制造", "PRICE2");
			map.put("固定制造", "PRICE3");
			map.put("日期", "CREATETIME");
			map.put("属性","PROPERTION");
		}
		else if("14".equals(imptype)) {//成本综合调整
			map.put("公司名称", "UNITNAME");
			map.put("产品编码", "INVCODE");
			map.put("产品名称", "INVNAME");
			map.put("配方编码", "INVCODE_PF");
			map.put("配方名称", "INVNAME_PF");
			map.put("吨调整", "PRICE");
			map.put("调整金额", "CB");
			map.put("备注", "BZ");
			map.put("日期", "CREATETIME");
		}
		else if("15".equals(imptype)) {
			map.put("客户编码", "CUSTCODE");
			map.put("客户名称", "CUSTNAME");
			map.put("小区", "OLDXQ");
			map.put("修正小区", "NEWXQ");
			map.put("日期", "CREATETIME");
		}
		else if("16".equals(imptype)) {
			map.put("工厂", "UNITNAME");
			map.put("年份", "YEARS");
			map.put("1月汇率", "RATE1");
			map.put("2月汇率", "RATE2");
			map.put("3月汇率", "RATE3");
			map.put("4月汇率", "RATE4");
			map.put("5月汇率", "RATE5");
			map.put("6月汇率", "RATE6");
			map.put("7月汇率", "RATE7");
			map.put("8月汇率", "RATE8");
			map.put("9月汇率", "RATE9");
			map.put("10月汇率", "RATE10");
			map.put("11月汇率", "RATE11");
			map.put("12月汇率", "RATE12");
		}
		else if("17".equals(imptype)) {
			map.put("月份", "MONTHS");
			map.put("公司", "UNITNAME");
			map.put("账面其他收入", "QTSR");
			map.put("税务收入", "SWSR");
			map.put("财务费用", "CWFY");
		}
		else if("18".equals(imptype)) {
			map.put("销售分类", "SALETYPE");
			map.put("销售大类", "DOCNAME");
			map.put("上交其他公司利润-跨公司", "PRICE1");
			map.put("收其他部利润-跨公司", "PRICE2");
			map.put("关联利润标准", "STANDPRICE1");
			map.put("上交产品线利润标准", "STANDPRICE2");
			map.put("日期", "CREATETIME");
		}
		else if("19".equals(imptype)) {
			map.put("提货工厂", "UNITNAME");
			map.put("NC客户名称", "BSCNAME");
			map.put("金额", "PRICE");
			map.put("日期", "MONTH");
		}
		else if("20".equals(imptype)) {
			map.put("层级", "CJ");
			map.put("战区", "ZQ");
			map.put("省区", "SQ");
			map.put("部", "B");
			map.put("制造费用补贴分类", "TYPE");
			map.put("日期", "CREATETIME");
		}
		else if("21".equals(imptype)) {
			map.put("项目", "PRO");
			map.put("销售分类", "SALETYPE");
			map.put("代销其他部利润", "QTBLR");
			map.put("放养加价标准", "JJBZ");
			map.put("吨变动制造结构", "BDZZJG");
			map.put("日期", "CREATETIME");
		}
		else if("22".equals(imptype)) {
			map.put("工厂", "UNITNAME");
			map.put("简称", "SHORTNAME");
			map.put("公司归属营销部", "YXB");
		}
		else if("23".equals(imptype)) {
			map.put("月份", "CREATETIME");
			map.put("营销部", "YXB");
			map.put("小区", "XQ");
			map.put("产品线", "CPX");
			map.put("项目", "PRO");
			map.put("分摊方式", "TYPE");
			map.put("金额（元）", "PRICE");
			map.put("说明", "BZ");
		}
		else if("24".equals(imptype)) {
			map.put("月份", "CREATETIME");
			map.put("营销部", "YXB");
			map.put("销售分类", "SALETYPE");
			map.put("产品编码", "INVCODE");
			map.put("产品名称", "INVNAME");
			map.put("产品线", "CPX");
		}
		else if("25".equals(imptype)) {
			map.put("月份", "CREATETIME");
			map.put("部门", "YXB");
			map.put("产品线", "CPX");
		}
		else if("26".equals(imptype)) {
			map.put("月份", "CREATETIME");
			map.put("层级", "CJ");
			map.put("部门", "B");
			map.put("产品线", "CPX");
			map.put("金额（元）", "PRICE");
			map.put("说明", "BZ");
		}
		else if("27".equals(imptype)) {
			map.put("月份", "CREATETIME");
			map.put("公司","UNITNAME");
			map.put("促销品编码", "INVCODE");
			map.put("促销品名称","INVNAME");
			map.put("促销品单价","PRICE" );
		}
		else if("28".equals(imptype)) {
			map.put("月份", "CREATETIME");
			map.put("公司","UNITNAME");
			map.put( "计奖产品编码","INVCODE");
			map.put( "计奖产品","INVNAME");
			map.put("产品线","CPX" );
			map.put("促销类型","TYPE" );
			map.put( "促销金额","PRICE");
			map.put( "说明","BZ");
		}
		else if("29".equals(imptype)) {
			map.put("月份", "CREATETIME");
			map.put("层级","CJ");
			map.put("产品线","CPX" );
			map.put("毛利系数","RATE" );
		}
		else if("30".equals(imptype)) {
			map.put("月份", "CREATETIME");
			map.put("所得税率", "RATE");
		}
		else if("31".equals(imptype)) {
			map.put("计提月份", "CREATETIME");
			map.put("计提类型", "TYPE");
			map.put("当月(计提标准）", "PRICE");
			map.put("计提说明", "BZ");
		}
		else if("32".equals(imptype)) {
			map.put("工厂", "UNITNAME");
			map.put("销售分类", "SALETYPE");
			map.put("存货编码", "INVCODE");
			map.put("存货名称", "INVNAME");
			map.put("业务类型", "CPX");
			map.put("标准毛利", "PRICE1");
			map.put("分析用毛利", "PRICE2");
			map.put("月份", "CREATETIME");
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
				SimpleDateFormat format = new SimpleDateFormat("yyyy-MM");
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
										/*//特殊成本导入模板特殊情况
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
														throw new Exception("第" + i + "行" + j
																+ "列为空，请填写");
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
				throw new Exception("只支持2003版本的Excel导入！");
			}
			StringBuffer sBuffer = new StringBuffer();
			String imptype = request.getAttribute("imptype").toString();
			//导入NC新工厂
			if ("1".equals(imptype)) {
				imp_ncnewcompany(list, request, response, imptype);
			} 
			//导入预混料调整
			else if("2".equals(imptype)) {
				imp_yhladjust(list, request, response, imptype);
			}
			//导入鱼粉调整
			else if("3".equals(imptype)) {
				imp_yfadjust(list, request, response, imptype);
			}
			//导入特殊成本
			else if("4".equals(imptype)) {
				imp_specialcost(list, request, response, imptype);
			}
			//导入关联成本调整
			else if("5".equals(imptype)) {
				imp_glcbadjust(list, request, response, imptype);
			}
			//导入采购KCP利润分配
			else if("6".equals(imptype)) {
				imp_cgkcp(list, request, response, imptype);
			}
			//导入存货计提
			else if("7".equals(imptype)) {
				imp_chjx(list, request, response, imptype);
			}
			//导入制造性折旧分摊
			else if("8".equals(imptype)) {
				imp_zjshare(list, request, response, imptype);
			}
			//导入固定资产计息
			if("9".equals(imptype)) {
				imp_assetjx(list, request, response, imptype);
			}
			//导入其他分摊
			else if("10".equals(imptype)) {
				imp_otherft(list, request, response, imptype);
			}
			//导入应剔除NC制造成本
			else if("11".equals(imptype)) {
				imp_rejectnccost(list, request, response, imptype);
			}
			//导入NC收购工厂成本
			else if("12".equals(imptype)) {
				imp_sgcompany(list, request, response, imptype);
			}
			//导入新工厂及收购工厂
			if("13".equals(imptype)) {
				imp_newcompany(list, request, response, imptype);
			}
			//导入成本综合调整
			else if("14".equals(imptype)) {
				imp_zhadjust(list, request, response, imptype);
			}
			//导入修正小区
			else if("15".equals(imptype)) {
				imp_xzxq(list, request, response, imptype);
			}
			//导入海外工厂汇率
			else if("16".equals(imptype)) {
				imp_hwhl(list, request, response, imptype);
			}
			//导入其他收入和财务费用
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
				request.setAttribute("msg", "如下人员未导入(其他导入成功)："
						+ sBuffer.toString());
			}

		} catch (Exception e) {
			dqrow = dqrow + 1;
			throw e;

		}
	}

	//NC新工厂导入
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_newnccompany WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// 校验新公司编码是否存在
			if (!vo.get("NEWUNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("NEWUNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行新公司编码不存在，请检查后重新导入！");
				}
			}
			
			// 校验取数公司编码是否存在
			if (!vo.get("QSUNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("QSUNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行取数公司编码不存在，请检查后重新导入！");
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	//预混料调整导入
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_yhladjust WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// 校验新公司编码是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行公司编码不存在，请检查后重新导入！");
				}
			}
			
			//校验产品编码是否存在
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行产品编码不存在，请检查后重新导入！");
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}
	
	//鱼粉调整导入
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_yfadjust WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//校验公司编码是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行公司编码不存在，请检查后重新导入！");
				}
			}
			
			//校验配方编码是否存在
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行配方编码不存在，请检查后重新导入！");
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
			// 提交事务
			System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	//特殊成本导入
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_specialcost WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//校验调整工厂编码是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行调整工厂编码不存在，请检查后重新导入！");
				}
			}
			
			if(vo.get("INVCODE_PF").toString().equals("")&&vo.get("INVCODE").toString().equals("")) {
				throw new Exception("除标题外第" + f + "行存货编码与成本编码均不存在，请检查后重新导入！");
			}
			
			//校验成本对象编码是否存在
			if (!vo.get("INVCODE_PF").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.BD_INVBASDOC B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE_PF"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行成本对象编码不存在，请检查后重新导入！");
				}
			}
			
			//校验存货编码是否存在
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.BD_INVBASDOC B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行存货编码不存在，请检查后重新导入！");
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}
	
	//关联成本调整导入
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.APP_MAIN_GLCBADJUST WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//校验调整工厂编码是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE MEMO=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行调整工厂编码不存在，请检查后重新导入！");
				}
				DbUtil.closeStatement(ps1);
			}
			
			//校验存货编码是否存在
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行存货编码不存在，请检查后重新导入！");
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}	

	//采购KCP利润分配导入
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_purkcpprofit WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//校验调整工厂编码是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE MEMO=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行调整工厂编码不存在，请检查后重新导入！");
				}
				 DbUtil.closeStatement(ps1);
			}
			
			//校验存货编码是否存在
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行存货编码不存在，请检查后重新导入！");
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
			 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}		

	//存货计息
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_chjx WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//校验调整工厂编码是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行调整工厂编码不存在，请检查后重新导入！");
				}
			}
			
			/*//校验存货编码是否存在
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE unitcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行存货编码不存在，请检查后重新导入！");
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}		

	//制造性折旧分摊导入
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_zjshare WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//校验调整工厂编码是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE MEMO=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行调整工厂编码不存在，请检查后重新导入！");
				}
			}
			
			/*//校验存货编码是否存在
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE unitcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行存货编码不存在，请检查后重新导入！");
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}
	
	//固定资产计提导入
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_assetjx WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			/*//校验调整工厂编码是否存在
			if (!vo.get("UNITCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE unitcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行调整工厂编码不存在，请检查后重新导入！");
				}
			}*/
			
			/*//校验存货编码是否存在
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE unitcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行存货编码不存在，请检查后重新导入！");
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}
	
	//其他分摊导入
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_otherft WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//校验调整工厂编码是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE MEMO=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行调整工厂编码不存在，请检查后重新导入！");
				}
			}
			
			/*//校验存货编码是否存在
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE unitcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行存货编码不存在，请检查后重新导入！");
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}
	
	//应剔除NC制造成本导入
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_rejectnccost WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//校验调整工厂编码是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行调整工厂编码不存在，请检查后重新导入！");
				}
				DbUtil.closeStatement(ps1);
			}
			
			//校验存货编码是否存在
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行存货编码不存在，请检查后重新导入！");
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

			// 提交事务
			//System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}
	
	//NC收购工厂成本导入
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.APP_MAIN_SGCOMPANY WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//校验调整工厂编码是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行调整工厂编码不存在，请检查后重新导入！");
				}
			}
			
			//校验存货编码是否存在
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行存货编码不存在，请检查后重新导入！");
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

			// 提交事务
			//System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}
	
	//新工厂及收购工厂导入
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.APP_MAIN_NEWCOMPANY WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//校验工厂编码是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行工厂编码不存在，请检查后重新导入！");
				}
			}
			
			//校验调整工厂编码是否存在
			if (!vo.get("FETCHUNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("FETCHUNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行调整工厂编码不存在，请检查后重新导入！");
				}
			}
			
			//校验存货编码是否存在
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行存货编码不存在，请检查后重新导入！");
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}
	
	//成本综合调整导入
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_zhadjust WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//校验调整工厂编码是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行调整工厂编码不存在，请检查后重新导入！");
				}
				DbUtil.closeStatement(ps1);
			}
			
			/*if(vo.get("INVCODE_PF").toString().equals("")&&vo.get("INVCODE").toString().equals("")&&!vo.get) {
				throw new Exception("除标题外第" + f + "行存货编码与成本编码均不存在，请检查后重新导入！");
			}*/
			
			//校验成本对象编码是否存在
			if (!vo.get("INVCODE_PF").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.BD_INVBASDOC B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE_PF"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行成本对象编码不存在，请检查后重新导入！");
				}
				DbUtil.closeStatement(ps1);
			}
			
			//校验存货编码是否存在
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.BD_INVBASDOC B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行存货编码不存在，请检查后重新导入！");
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

			// 提交事务
			//System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}
	
	//修正小区导入
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
		//删除包含的历史数据
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

			// 提交事务
			//System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}
	
	//海外汇率导入
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
			//删除包含的历史数据
			sql = "DELETE wb_erp.app_main_hwhl WHERE YEARS = "+obj.get("YEARS").toString()+" and pk_corp=(select pk_corp from wb_erp.bd_corp where memo='"+obj.get("UNITNAME").toString()+"')";
			ps1 = conn.prepareStatement(sql);
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//判断工厂是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行调整工厂编码不存在，请检查后重新导入！");
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

			// 提交事务
			//System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}
	
	//导入其他收入和财务费用
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_qtsr WHERE MONTHS in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//判断工厂是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行调整工厂编码不存在，请检查后重新导入！");
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}
	
	//导入其他收支
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_qtsz WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//判断工厂是否存在
//			if (!vo.get("UNITNAME").toString().equals("")) {
//				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
//				ps1 = conn.prepareStatement(sql);
//				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
//				rSet = ps1.executeQuery();
//				if (rSet.next()) {
//					result2 = rSet.getInt("CT");
//				}
//				if (result2 == 0) {
//					throw new Exception("除标题外第" + f + "行调整工厂编码不存在，请检查后重新导入！");
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}
	
	//应剔除办事处费用
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_tcbsc WHERE MONTH in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//判断工厂是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行提货工厂名称不存在，请检查后重新导入！");
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
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
		//删除包含的历史数据
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
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
		//删除包含的历史数据
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
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
			//判断工厂是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				PreparedStatement ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行工厂名称不存在，请检查后重新导入！");
				}
			}
			
			sql = "insert into wb_erp.app_main_gcjc (pk_corp,shortname,yxb)"
					+ "values ((select pk_corp FROM WB_ERP.bd_corp B WHERE memo=?),?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SHORTNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("YXB"));
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
		//删除包含的历史数据
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
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
		//删除包含的历史数据
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
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
		//删除包含的历史数据
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
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
		//删除包含的历史数据
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_cxpproduct WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//判断工厂是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行工厂名称不存在，请检查后重新导入！");
				}
			}
			//判断产品编码是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行工厂名称不存在，请检查后重新导入！");
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_jtcxfy WHERE CREATETIME in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//判断工厂是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行工厂名称不存在，请检查后重新导入！");
				}
			}
			//判断产品是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行产品编码不存在，请检查后重新导入！");
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
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
		//删除包含的历史数据
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
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
		//删除包含的历史数据
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
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
		//删除包含的历史数据
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
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
		//删除包含的历史数据
		sql = "DELETE wb_erp.app_main_bzml WHERE CREATETIME in "+Times;
		ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//判断工厂是否存在
			if (!vo.get("UNITNAME").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("UNITNAME"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行工厂名称不存在，请检查后重新导入！");
				}
				DbUtil.closeStatement(ps1);
			}
			//判断产品是否存在
			if (!vo.get("INVCODE").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_invbasdoc B WHERE invcode=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVCODE"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行产品编码不存在，请检查后重新导入！");
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

			// 提交事务
			System.out.println(f);
			// 关闭资源
			 DbUtil.closeStatement(ps1);
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