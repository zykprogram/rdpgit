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

public class CustmoerAwardImportToll {
	public static void getFile(HttpServletRequest request,
			HttpServletResponse response) throws Exception {

		InputStream in = (InputStream) request.getAttribute("uploadFile");
		String fileName = request.getAttribute("uploadFile__name").toString();
		String fileType = fileName.substring(fileName.lastIndexOf(".") + 1,
				fileName.length());
		String imptype = request.getAttribute("imptype").toString();
		Map<String, String> map = new HashMap<String, String>();
		if ("1".equals(imptype)) { // 开发-奖励方案

			map.put("方案编号", "SECHEMNO");
			map.put("方案名称", "SECHEMNAME");
			map.put("战区", "ZHANGQ");
			map.put("营销部", "YXB");
			map.put("存货产品线", "PK_PRODLINE");
			map.put("存货组合", "INVBASDOCID");
			map.put("开始日期", "STARTMONTH");
			map.put("结束日期", "ENDMONTH");
			map.put("受益期数", "COUNTMONTH");
			map.put("销量、标吨", "SALETYPE");
			map.put("销量开始", "SALESTART");
			map.put("销量结束", "SALEEND");
			map.put("存货组合", "INVBASDOCID");
			map.put("开发代表", "MONEY_YWY");
			map.put("开发组长", "MONEY_ZZ");
		}
		if ("2".equals(imptype)) { // 开发-受益期内稽核条件

			map.put("方案编号", "SECHEMNO");
			map.put("方案名称", "SECHEMNAME");
			map.put("战区", "ZHANGQ");
			map.put("营销部", "YXB");
			map.put("存货产品线", "PK_PRODLINE");
			map.put("开始月份", "STARTMONTH");
			map.put("结束月份", "ENDMONTH");
			map.put("受益期", "COUNTMONTH");
			map.put("计稽核不计奖产品", "PK_INVBASDOCGROUP");
			map.put("销量、标吨", "SALETYPE");
			map.put("受益期内连续未提货月份数", "COUNTMONTH");
			map.put("奖励扣回比例", "RATE");
		}
		if ("3".equals(imptype)) { // 开发-受益期次月稽核条件
			map.put("方案编号", "SECHEMNO");
			map.put("方案名称", "SECHEMNAME");
			map.put("战区", "ZHANGQ");
			map.put("营销部", "YXB");
			map.put("开始月份", "STARTMONTH");
			map.put("结束月份", "ENDMONTH");
			map.put("受益期", "COUNTNUMBER");
			map.put("存货产品线", "PK_PRODLINE");
			map.put("计稽核不计奖（产品组合）", "PK_INVBASDOCGROUP");
			map.put("月数（N）", "COUNTMONTH");
			map.put("销量、标吨", "SALETYPE");
			map.put("受益期内连续未提货（月份数）", "COUNTMONTH");
			map.put("奖励扣回（P）", "RATE");
			map.put("低于比例（M）", "LOWERRATE");
		}
		if ("4".equals(imptype)) { // 开发-受益期次两月稽核条件
			map.put("方案编号", "SECHEMNO");
			map.put("方案名称", "SECHEMNAME");
			map.put("战区", "ZHANGQ");
			map.put("营销部", "YXB");
			map.put("开始月份", "STARTMONTH");
			map.put("结束月份", "ENDMONTH");
			map.put("受益期", "COUNTNUMBER");
			map.put("存货产品线", "PK_PRODLINE");
			map.put("计稽核不计奖产品组合", "PK_INVBASDOCGROUP");
			map.put("类型", "TYPE");
			map.put("销量、标吨", "SALETYPE");
			map.put("考核标吨", "KHNUMBER");
			map.put("奖励扣回比例", "RATE");
			map.put("备注", "MEMO");
		}
		if ("5".equals(imptype)) { // 客户-奖励方案1-客户月奖
			map.put("方案编号", "SCHEMENO");
			map.put("方案名称", "SCHEMENAME");
			map.put("工厂", "PK_CORP");
			map.put("奖励类型", "TYPE");
			map.put("开始日期", "STARTDATE");
			map.put("结束日期", "ENDDATE");
			map.put("奖励项目", "DETAILNAME");
			map.put("客户组合名称", "CUSTOMERGROUP");
			map.put("奖励标准库名称", "SCHEMEGROUP");
			map.put("定级标准类型", "DJTYPE");
			map.put("产品库名称", "INVBADOCGROUP");
			map.put("产品库1名称", "INVBADOCGROUP1");
			map.put("产品库2名称", "INVBADOCGROUP2");
			map.put("产品库3名称", "INVBADOCGROUP3");
			map.put("产品库4名称", "INVBADOCGROUP4");
			map.put("计量方式", "CONDITIONJLFS");
			map.put("客户计量方式", "CONDITION0");
			map.put("分销挂户分销产品", "CONDITION1");
			map.put("分销挂户直配产品", "CONDITION2");
			map.put("直销配送分销产品", "CONDITION4");
			map.put("直销配送直配产品", "CONDITION5");
			map.put("基数开始", "STARTDATE_JS");
			map.put("基数结束", "ENDDATE_JS");
			map.put("奖励基数", "RATE");
			map.put("产品组合条件", "RATE_CONDITION");
		}

		if ("6".equals(imptype)) { // 客户-奖励方案4-专销合同表
			map.put("合同编号", "HTNO");
			map.put("结算周期", "JSZQ");
			map.put("公司PK", "PK_CORP");
			map.put("产品", "INVBADOCGROUP");
			map.put("标准", "SCHEMEGROUP");
			map.put("结算方式", "JSFS");
			map.put("最低提货量", "MINNNUMBER");
			map.put("销量或标吨", "SALETYPE");
		}

		if ("7".equals(imptype)) { // 客户-奖励方案5-专销合同表签订情况表
			map.put("合同编号", "HTNO");
			map.put("合同类型", "HTTYPE");
			map.put("客户编码", "CUSTCODE");
			map.put("客户姓名", "CUSTNAME");
			map.put("计划生效日期", "STARTDATE");
			map.put("计划终止日期", "ENDDATE");
			map.put("分销挂户双金", "TYPE1");
			map.put("分销挂户猪三乐", "TYPE2");
			map.put("分销挂户预混料", "TYPE3");
			map.put("直销配送双金", "TYPE4");
			map.put("直销配送猪三乐", "TYPE5");
			map.put("直销配送预混料", "TYPE6");
			map.put("直销配送预混料", "TYPE6");
			map.put("客户本身提猪三乐计量方式", "TYPE7");
			map.put("客户计量方式", "TYPE8");
		}

		if ("8".equals(imptype)) { // 客户-奖励方案6-客户年终奖
			map.put("合同编号", "HTNO");
			map.put("合同类型", "HTTYPE");
			map.put("客户编码", "CUSTCODE");
			map.put("客户姓名", "CUSTNAME");
			map.put("计划生效日期", "STARTDATE");
			map.put("计划终止日期", "ENDDATE");
			map.put("产品", "INVBADOCGROUP");
			map.put("奖励标准", "SCHEMEGROUP");
			map.put("0客户计量方式", "TYPE0");
			map.put("7客户本身提猪三乐", "TYPE7");
			map.put("6客户本身提三胞胎", "TYPE6");
			map.put("5直配客户本身提双金产品", "TYPE5");
			map.put("1分销挂户双金", "TYPE1");
			map.put("2分销挂户猪三乐", "TYPE2");
			map.put("3直销配送双金", "TYPE3");
			map.put("4直销配送猪三乐", "TYPE4");
			map.put("分销挂户预混料", "TYPE8");
			map.put("直销配送三胞胎", "TYPE9");
			map.put("直销配送预混料", "TYPE10");
			map.put("销量或标吨", "SALETYPE");
		}
		if ("9".equals(imptype)) { // 特殊客户-不结算客户列表
			map.put("客户编码", "CUSTCODE");
			map.put("客户名称", "CUSTNAME");
			map.put("起始月份", "STARTMONTH");
			map.put("结束月份", "ENDMONTH");
			map.put("结算比例", "RATE");
			map.put("奖励类型", "TYPE");
		}
		if ("10".equals(imptype) || "11".equals(imptype)) {// 价值客户目标和客户利润目标
			map.put("客户编码", "CUSTCODE");
			map.put("客户姓名", "CUSTNAME");
			map.put("销量期间", "MONTHYEAR");
			map.put("预算目标（吨）", "TARGET");
			map.put("战区", "ZHANGQ");
			map.put("营销部", "YXB");
			map.put("客户分类", "CUSTTYPE");
			map.put("促销费（元）", "MONEYOFPROMTION");
			map.put("贷款需求（元）", "MONEYOFBANK");
			map.put("产品行销", "CPXX");
			map.put("片组", "SALESTR");
		}
		if ("12".equals(imptype)) { // 客户主管方案表
			map.put("方案编号", "SECHNO");
			map.put("方案名称", "SECHNAME");
			map.put("结算日期", "JSDATE");
			map.put("战区", "ZHANGQ");
			map.put("营销部", "YXB");
			map.put("片组", "SALESTR");
			map.put("组织", "UNIT");
			map.put("奖励标准", "STANDARD");
			map.put("基础任务", "BASENUMBER");
			map.put("目标任务", "HIGHERNUMBER");
			map.put("低于基础任务单价", "LOWERNUMBER");
			map.put("基础任务固定单价", "BASEGDPRICE");
			map.put("基础任务单价", "BASEPRICE");
			map.put("超目标单价系数", "XS");
			map.put("超目标任务单价", "HIGHERPRICE");
			map.put("超目标任务封顶单价", "HIGHERMAXPRICE");
		}

		if ("13".equals(imptype)) {// 片组奖励分配比例
			map.put("方案编号", "SECHNO");
			map.put("方案名称", "SECHNAME");
			map.put("结算日期", "JSMONTH");
			map.put("战区", "ZHANGQ");
			map.put("营销部", "YXB");
			map.put("片组", "SALESTR");
			map.put("客户主管编码", "PSNCODE");
			map.put("客户主管姓名", "PSNNAME");
			map.put("在岗系数", "STATUS");
			map.put("分配比例", "XS");
			map.put("预发比例", "BL");
		}
		if ("14".equals(imptype)) {// 任务分解表
			map.put("方案编号", "SCHEMNO");
			map.put("方案名称", "SCHEMNAME");
			map.put("战区", "ZHANGQ");
			map.put("营销部", "YXB");
			map.put("客户编码", "CUSTCODE");
			map.put("客户名称", "CUSTNAME");
			map.put("级别", "LEVEL_CUST");
			map.put("销量考核类型", "CUSTTYPE");
			map.put("销量考核取数口径", "INVBASDOCID");
			map.put("销量考核特殊产品取数口径", "INVBASDOCID_OTHER");
			map.put("销量/标吨", "SALETYPE");
			map.put("网点考核取数口径", "LEVEL_WD");
			map.put("销量月份1", "TARGET_MONTH1");
			map.put("片组1", "TARGET_SALESTR1");
			map.put("客户主管编码1", "TARGET_PSNCODE1");
			map.put("客户主管姓名1", "TARGET_PSNNAME1");
			map.put("销量月份2", "TARGET_MONTH2");
			map.put("片组2", "TARGET_SALESTR2");
			map.put("客户主管编码2", "TARGET_PSNCODE2");
			map.put("客户主管姓名2", "TARGET_PSNNAME2");
			map.put("销量月份3", "TARGET_MONTH3");
			map.put("片组3", "TARGET_SALESTR3");
			map.put("客户主管编码3", "TARGET_PSNCODE3");
			map.put("客户主管姓名3", "TARGET_PSNNAME3");
		}
		if ("15".equals(imptype)) {// 人才选择
			map.put("月份", "MONTH");
			map.put("姓名", "NAME");
			map.put("性别", "SEX");
			map.put("年龄", "AGE");
			map.put("学历", "SCHOOLLEVEL");
			map.put("籍贯", "ADDRESS");
			map.put("毕业院校", "SCHOOL");
			map.put("专业", "MAJOR");
			map.put("联系电话", "PHONE");
			map.put("面试阶段_面试官", "INTERVIEWER");
			map.put("面试阶段_面试分数", "INTERVIEWERSCORE");
			map.put("面试阶段_面试结果", "INTERVIEWERRESULT");
			map.put("面试阶段_面试评价", "EVALUATION");
			map.put("拟录用阶段_到岗", "COMEPOSTION");
			map.put("拟录用阶段_到岗月份", "COMEPOSTIONMONTH");
			map.put("拟录用阶段_到岗时间", "COMEPOSTIONTIME");
			map.put("拟录用阶段_未到岗原因", "REASON");
			map.put("观摩阶段_起始日期", "WATCHSTARTDATE");
			map.put("观摩阶段_结束日期", "WATCHENDDATE");
			map.put("观摩阶段_观摩部门", "WATCHDEPT");
			map.put("观摩阶段_观摩师傅", "WATCHMASTERWORKER");
			map.put("观摩阶段_观摩分数", "WATCHSCORE");
			map.put("观摩阶段_观摩结果", "WATCHRESULT");
			map.put("观摩阶段_师傅评价", "WATCHMASTEREVALUATION");
			map.put("入司培训阶段_起始日期", "TRAINSTARTDATE");
			map.put("入司培训阶段_结束日期", "TRAINENDDATE");
			map.put("入司培训阶段_培训分数", "TRAINSCORE");
			map.put("入司培训阶段_培训结果", "TRAINRESULT");
			map.put("入司培训阶段_培训评价", "TRAINEVALUATION");
			map.put("入司培训阶段_入职部门", "TRAINDEPT");
			map.put("入司培训阶段_入职岗位", "TRAINPOST");
		}
		if ("16".equals(imptype)) {// 人才培养
			map.put("中心/战区", "ZHANGQ");
			map.put("部门/公司", "YXB");
			map.put("带教业务属性", "TYPE");
			map.put("徒弟信息_人员编码", "PSNCODE_TD");
			map.put("徒弟信息_姓名", "PSNNAME_TD");
			map.put("徒弟信息_岗位", "POSTNAME_TD");
			map.put("师傅信息_人员编码", "PSNCODE_SF");
			map.put("师傅信息_姓名", "PSNNAME_SF");
			map.put("师傅信息_岗位", "POSTNAME_SF");
			map.put("师傅信息_人员编码1", "PSNCODE_SF1");
			map.put("师傅信息_姓名1", "PSNNAME_SF1");
			map.put("师傅信息_岗位1", "POSTNAME_SF1");
			map.put("带教开始时间", "BEGIN_TRAIN");
			map.put("带教结束时间", "END_TRAIN");
			map.put("带教天数", "TRAINGDAYS");
			map.put("验收时间", "ACCEPTANCETIME");
			map.put("成绩_学习地图", "SCORE1");
			map.put("成绩_验收成绩", "SCORE2");
			map.put("带教验收_综合成绩", "SCORE3");
			map.put("带教验收_验收结果", "TRAIN_CONCLUSION");
			map.put("干部理论成绩", "TRAIN_ACHIEVEMENT");
			map.put("上岗证编号", "TRAIN_POSTCARD");
			map.put("评价(师傅/部门长)_性格", "CHARACTER");
			map.put("评价(师傅/部门长)_责任心", "CONSCIENTIOUSNESS");
			map.put("评价(师傅/部门长)_价值观", "SENSEOFWORTH");
			map.put("评价(师傅/部门长)_综合能力", "COMPREHENSIVE");
			map.put("评价(师傅/部门长)_结论", "CONCLUSION");
			map.put("观摩阶段_顶岗部门", "TOPJOBDEPARTMENT");
			map.put("观摩阶段_顶岗岗位", "POSTPOSITION");
			map.put("观摩阶段_结论", "CONCLUSIONPOST");
			map.put("入司培训阶段_得分", "REGULAR_SCORE");
			map.put("入司培训阶段_结果", "REGULAR_CONCLUSION");
		}
		if ("17".equals(imptype)) {// 采购头寸
			map.put("日期", "BILLDATE");
			map.put("大区", "ZHANGQ");
			map.put("省区", "PROVICE");
			map.put("公司", "UNITNAME");
			map.put("原料名称", "INVNAME");
			map.put("库存量", "INHAND");
			map.put("订单", "ORDERNO");
			map.put("合计", "SUMNO");
			map.put("库存价", "INHANDPRICE");
			map.put("订单价", "ORDERPRICE");
			map.put("加权平均价", "AVGPRICE");
			map.put("行情价", "HQPRICE");
			map.put("吨盈(亏)", "YK");
			map.put("盈亏金额/万元", "YKMONEY");
			map.put("30日平均用量(吨)", "NNUMBER");
			map.put("头寸天数(天)", "DAYS");

		}

		if ("18".equals(imptype)) {// 利润模块-客户奖励手工单
			map.put("奖励类型", "TYPE");
			map.put("工厂", "UNITNAME");
			map.put("客户编码", "CUSTCODE");
			map.put("客户名称", "CUSTNAME");
			map.put("产品组合", "INVBASDOCGROUP");
			map.put("产品组合1", "INVBASDOCGROUP1");
			map.put("产品组合2", "INVBASDOCGROUP2");
			map.put("计提或调整月份", "JTMONTH");
			map.put("计提或调整金额", "JTMONEY");
			map.put("调整周期", "ZQ");
			map.put("备注", "MEMO");

		}
		if ("19".equals(imptype)) {// 基础_奖励标准
			// map.put("标准时间","STANDID");
			map.put("奖励标准名称", "STANDRADNAME");
			// map.put("类型", "TYPE");
			map.put("开始阶梯", "STARTNUM");
			map.put("结束阶梯", "ENDNUM");
			map.put("单价", "PRICE");
			// map.put("是否封存", "FLAG");
			// map.put("TS", "TS");
		}
		if("20".equals(imptype)) {//预警目标（战区）
			map.put("战区", "ZHANQ");
			map.put("月份", "MONTH");
			map.put("本月目标销量（吨）", "MBNNUMBER");
		}
		if("21".equals(imptype)) {//预警目标（行销）
			map.put("行销", "CPX");
			map.put("月份", "MONTH");
			map.put("本月目标销量（吨）", "MBNNUMBER");
		}
		if("22".equals(imptype)) {//预警目标（营销部）
			map.put("战区", "ZHANQ");
			map.put("营销部", "YXB");
			map.put("月份", "MONTH");
			map.put("本月目标销量（吨）", "MBNNUMBER");
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

								/*
								 * System.out.print(cellVal);
								 * System.out.println(i + "行");
								 * System.out.println(map.get(headRow.getCell(j)
								 * .getStringCellValue().toString().trim()
								 * .replaceAll("\r|\n", "")));
								 */

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
				imp_sbt_employee_scheme_kf(list, request, response);
			}
			if ("2".equals(imptype) || "3".equals(imptype)
					|| "4".equals(imptype)) {
				imp_sbt_employee_scheme_kf_check(list, request, response,
						imptype);
			}
			if ("5".equals(imptype)) {
				imp_SBT_CUST_SCHEME(list, request, response);
			}

			if ("6".equals(imptype)) {
				imp_SBT_CUST_HT(list, request, response);
			}
			if ("7".equals(imptype)) {
				imp_sbt_cust_ht_sign(list, request, response);
			}
			if ("8".equals(imptype)) {
				imp_sbt_cust_ht_year(list, request, response);
			}
			if ("9".equals(imptype)) {
				imp_sbt_cust_specialcust(list, request, response);
			}
			if ("10".equals(imptype) || "11".equals(imptype)) {
				imp_saletarget_customer(list, request, response, imptype);
			}
			if ("12".equals(imptype)) {
				imp_SBT_CUST_EMPLOYEE_SCHEME(list, request, response);
			}
			if ("13".equals(imptype)) {
				imp_sbt_cust_employee_standrad(list, request, response);
			}
			if ("14".equals(imptype)) {
				imp_SBT_CUST_EMPLOYEE_TARGET(list, request, response);
			}
			if ("15".equals(imptype)) {
				imp_sbt_hr_personchose(list, request, response);
			}
			if ("16".equals(imptype)) {
				imp_sbt_hr_persontraining(list, request, response);
			}
			if ("17".equals(imptype)) {
				imp_sbt_cg_markettc(list, request, response);
			}
			if ("18".equals(imptype)) {
				imp_sjlr_cust_sum(list, request, response);
			}
			if ("19".equals(imptype)) {
				imp_sbt_cust_standrad(list, request, response);
			}
			if("20".equals(imptype)) {//预警目标（战区）
				imp_app_main_yjmb_zq(list, request, response);
			}
			if("21".equals(imptype)) {//预警目标（行销）
				imp_app_main_yjmb_xx(list, request, response);
			}
			if("22".equals(imptype)) {//预警目标（营销部）
				imp_app_main_yjmb_yxb(list, request, response);
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
	 *月度利润-客户奖励手工单
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sjlr_cust_sum(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps2 = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// 先删除月度的数据
			sql = "DELETE wb_erp.sjlr_cust_sum where  jtmonth=? and type=? ";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("JTMONTH"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("TYPE"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		int i = 0;
		for (int f = 0; f < voList.size(); f++) {
			i++;
			System.out.println(i);
			JSONObject vo = voList.get(f);

			// 校验存货组合
			sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVBASDOCGROUP"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("INVBADOCGROUP").toString()
						+ "存货组合不存在，请检查后重新导入！");
			}
			DbUtil.closeStatement(ps1);
			if (!vo.get("INVBASDOCGROUP1").toString().equals("")) {
				// 校验存货组合1
				sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo
						.get("INVBASDOCGROUP1"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + 1 + "行"
							+ vo.get("INVBADOCGROUP").toString()
							+ "存货组合1不存在，请检查后重新导入！");
				}
				DbUtil.closeStatement(ps1);
			}
			// 校验存货组合2
			if (!vo.get("INVBASDOCGROUP2").toString().equals("")) {
				sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo
						.get("INVBASDOCGROUP2"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + 1 + "行"
							+ vo.get("INVBADOCGROUP").toString()
							+ "存货组合2不存在，请检查后重新导入！");
				}
				DbUtil.closeStatement(ps1);
			}

			sql = "insert into wb_erp.sjlr_cust_sum (TYPE, UNITNAME, CUSTCODE, CUSTNAME, INVBASDOCGROUP, JTMONTH, JTMONEY, ZQ, MEMO, ID,INVBASDOCGROUP1,INVBASDOCGROUP2)"
					+ " values (?,?,?,?, (select id  from wb_erp.sbt_cust_invbasdocgroup a  where nvbasdocgroup=? and rownum=1) , ?,?, ?, ?, ?,(select id  from wb_erp.sbt_cust_invbasdocgroup a  where nvbasdocgroup=? and rownum=1),(select id  from wb_erp.sbt_cust_invbasdocgroup a  where nvbasdocgroup=? and rownum=1))";

			ps2 = conn.prepareStatement(sql);
			DbUtil.setObject(ps2, 1, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps2, 2, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps2, 3, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps2, 4, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps2, 5, Types.VARCHAR, vo.opt("INVBASDOCGROUP"));
			DbUtil.setObject(ps2, 6, Types.VARCHAR, vo.opt("JTMONTH"));
			DbUtil.setObject(ps2, 7, Types.VARCHAR, vo.opt("JTMONEY"));
			DbUtil.setObject(ps2, 8, Types.VARCHAR, vo.opt("ZQ"));
			DbUtil.setObject(ps2, 9, Types.VARCHAR, vo.opt("MEMO"));
			String PK_ID = SysUtil.getId();
			ps2.setString(10, PK_ID);
			DbUtil.setObject(ps2, 11, Types.VARCHAR, vo.opt("INVBASDOCGROUP1"));
			DbUtil.setObject(ps2, 12, Types.VARCHAR, vo.opt("INVBASDOCGROUP2"));
			ps2.execute();
			ps2.close();
			// 关闭资源
			DbUtil.closeStatement(ps2);
		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}

	/**
	 *采购头寸
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_cg_markettc(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps2 = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// 先删除月度的数据
			sql = "DELETE wb_erp.sbt_cg_markettc where billdate=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("BILLDATE"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			sql = "insert into wb_erp.sbt_cg_markettc(billdate  ,zhangq  ,provice  ,unitname  ,invname  ,inhand  ,orderno  ,sumno  ,inhandprice  ,orderprice  ,avgprice  ,hqprice  ,yk  ,ykmoney	,nnumber	,days	) "
					+ " values(?  ,?  ,?  ,?  ,?  ,?  ,?  ,?  ,?  ,?  ,?  ,?  ,?  ,?	,?	,?	)";

			ps2 = conn.prepareStatement(sql);
			DbUtil.setObject(ps2, 1, Types.VARCHAR, vo.opt("BILLDATE"));
			DbUtil.setObject(ps2, 2, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps2, 3, Types.VARCHAR, vo.opt("PROVICE"));
			DbUtil.setObject(ps2, 4, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps2, 5, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps2, 6, Types.VARCHAR, vo.opt("INHAND"));
			DbUtil.setObject(ps2, 7, Types.VARCHAR, vo.opt("ORDERNO"));
			DbUtil.setObject(ps2, 8, Types.VARCHAR, vo.opt("SUMNO"));
			DbUtil.setObject(ps2, 9, Types.VARCHAR, vo.opt("INHANDPRICE"));
			DbUtil.setObject(ps2, 10, Types.VARCHAR, vo.opt("ORDERPRICE"));
			DbUtil.setObject(ps2, 11, Types.VARCHAR, vo.opt("AVGPRICE"));
			DbUtil.setObject(ps2, 12, Types.VARCHAR, vo.opt("HQPRICE"));
			DbUtil.setObject(ps2, 13, Types.VARCHAR, vo.opt("YK"));
			DbUtil.setObject(ps2, 14, Types.VARCHAR, vo.opt("YKMONEY"));
			DbUtil.setObject(ps2, 15, Types.VARCHAR, vo.opt("NNUMBER"));
			DbUtil.setObject(ps2, 16, Types.VARCHAR, vo.opt("DAYS"));
			ps2.execute();
			ps2.close();
			// 关闭资源

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}

	/**
	 * 人才培养
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_hr_persontraining(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps2 = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// 先删除月度的数据
			sql = "DELETE wb_erp.sbt_hr_persontraining where PSNCODE_TD=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("PSNCODE_TD"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			sql = "insert into wb_erp.sbt_hr_persontraining(ID, ZHANGQ,YXB,TYPE,	 PSNCODE_TD,	 PSNNAME_TD,	 POSTNAME_TD,	 PSNCODE_SF,	 PSNNAME_SF,	 POSTNAME_SF,	 BEGIN_TRAIN,	 END_TRAIN,	 TRAINGDAYS,	 ACCEPTANCETIME,	 SCORE1,	 SCORE2,	 SCORE3,	 TRAIN_CONCLUSION,	 CHARACTER,	 CONSCIENTIOUSNESS,	 SENSEOFWORTH,	 COMPREHENSIVE,	 CONCLUSION,	 TOPJOBDEPARTMENT,	 POSTPOSITION,	 CONCLUSIONPOST,	 REGULAR_SCORE,	 REGULAR_CONCLUSION, TS, COPERATOR,TRAIN_ACHIEVEMENT,TRAIN_POSTCARD,PSNCODE_SF1,PSNNAME_SF1,POSTNAME_SF1) "
					+ " values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?,?,?,?)";

			ps2 = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps2.setString(1, PK_ID);
			DbUtil.setObject(ps2, 2, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps2, 3, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps2, 4, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps2, 5, Types.VARCHAR, vo.opt("PSNCODE_TD"));
			DbUtil.setObject(ps2, 6, Types.VARCHAR, vo.opt("PSNNAME_TD"));
			DbUtil.setObject(ps2, 7, Types.VARCHAR, vo.opt("POSTNAME_TD"));
			DbUtil.setObject(ps2, 8, Types.VARCHAR, vo.opt("PSNCODE_SF"));
			DbUtil.setObject(ps2, 9, Types.VARCHAR, vo.opt("PSNNAME_SF"));
			DbUtil.setObject(ps2, 10, Types.VARCHAR, vo.opt("POSTNAME_SF"));
			DbUtil.setObject(ps2, 11, Types.VARCHAR, vo.opt("BEGIN_TRAIN"));
			DbUtil.setObject(ps2, 12, Types.VARCHAR, vo.opt("END_TRAIN"));
			DbUtil.setObject(ps2, 13, Types.VARCHAR, vo.opt("TRAINGDAYS"));
			DbUtil.setObject(ps2, 14, Types.VARCHAR, vo.opt("ACCEPTANCETIME"));
			DbUtil.setObject(ps2, 15, Types.VARCHAR, vo.opt("SCORE1"));
			DbUtil.setObject(ps2, 16, Types.VARCHAR, vo.opt("SCORE2"));
			DbUtil.setObject(ps2, 17, Types.VARCHAR, vo.opt("SCORE3"));
			DbUtil
					.setObject(ps2, 18, Types.VARCHAR, vo
							.opt("TRAIN_CONCLUSION"));
			DbUtil.setObject(ps2, 19, Types.VARCHAR, vo.opt("CHARACTER"));
			DbUtil.setObject(ps2, 20, Types.VARCHAR, vo
					.opt("CONSCIENTIOUSNESS"));
			DbUtil.setObject(ps2, 21, Types.VARCHAR, vo.opt("SENSEOFWORTH"));
			DbUtil.setObject(ps2, 22, Types.VARCHAR, vo.opt("COMPREHENSIVE"));
			DbUtil.setObject(ps2, 23, Types.VARCHAR, vo.opt("CONCLUSION"));
			DbUtil
					.setObject(ps2, 24, Types.VARCHAR, vo
							.opt("TOPJOBDEPARTMENT"));
			DbUtil.setObject(ps2, 25, Types.VARCHAR, vo.opt("POSTPOSITION"));
			DbUtil.setObject(ps2, 26, Types.VARCHAR, vo.opt("CONCLUSIONPOST"));
			DbUtil.setObject(ps2, 27, Types.VARCHAR, vo.opt("REGULAR_SCORE"));
			DbUtil.setObject(ps2, 28, Types.VARCHAR, vo
					.opt("REGULAR_CONCLUSION"));
			DbUtil.setObject(ps2, 29, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps2, 30, Types.VARCHAR, vo
					.opt("TRAIN_ACHIEVEMENT"));
			DbUtil.setObject(ps2, 31, Types.VARCHAR, vo.opt("TRAIN_POSTCARD"));
			DbUtil.setObject(ps2, 32, Types.VARCHAR, vo.opt("PSNCODE_SF1"));
			DbUtil.setObject(ps2, 33, Types.VARCHAR, vo.opt("PSNNAME_SF1"));
			DbUtil.setObject(ps2, 34, Types.VARCHAR, vo.opt("POSTNAME_SF1"));
			ps2.execute();
			ps2.close();
			// 关闭资源

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}

	/**
	 * 人才选择
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_hr_personchose(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps2 = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			sql = "insert into wb_erp.sbt_hr_personchose (ID, MONTH, NAME, SEX, AGE, SCHOOLLEVEL, ADDRESS, SCHOOL, MAJOR, PHONE, INTERVIEWER, INTERVIEWERSCORE, INTERVIEWERRESULT, EVALUATION, COMEPOSTION, COMEPOSTIONMONTH, COMEPOSTIONTIME, REASON, WATCHSTARTDATE, WATCHENDDATE, WATCHDEPT, WATCHMASTERWORKER, WATCHSCORE, WATCHRESULT, WATCHMASTEREVALUATION, TRAINSTARTDATE, TRAINENDDATE, TRAINSCORE, TRAINRESULT, TRAINEVALUATION, TRAINDEPT, TRAINPOST, TS, COPERATOR) "
					+ " values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?)";

			ps2 = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			// 赋值至子表的主外键
			String FK_ID = PK_ID;
			ps2.setString(1, PK_ID);
			DbUtil.setObject(ps2, 2, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps2, 3, Types.VARCHAR, vo.opt("NAME"));
			DbUtil.setObject(ps2, 4, Types.VARCHAR, vo.opt("SEX"));
			DbUtil.setObject(ps2, 5, Types.VARCHAR, vo.opt("AGE"));
			DbUtil.setObject(ps2, 6, Types.VARCHAR, vo.opt("SCHOOLLEVEL"));
			DbUtil.setObject(ps2, 7, Types.VARCHAR, vo.opt("ADDRESS"));
			DbUtil.setObject(ps2, 8, Types.VARCHAR, vo.opt("SCHOOL"));
			DbUtil.setObject(ps2, 9, Types.VARCHAR, vo.opt("MAJOR"));
			DbUtil.setObject(ps2, 10, Types.VARCHAR, vo.opt("PHONE"));
			DbUtil.setObject(ps2, 11, Types.VARCHAR, vo.opt("INTERVIEWER"));
			DbUtil
					.setObject(ps2, 12, Types.VARCHAR, vo
							.opt("INTERVIEWERSCORE"));
			DbUtil.setObject(ps2, 13, Types.VARCHAR, vo
					.opt("INTERVIEWERRESULT"));
			DbUtil.setObject(ps2, 14, Types.VARCHAR, vo.opt("EVALUATION"));
			DbUtil.setObject(ps2, 15, Types.VARCHAR, vo.opt("COMEPOSTION"));
			DbUtil
					.setObject(ps2, 16, Types.VARCHAR, vo
							.opt("COMEPOSTIONMONTH"));
			DbUtil.setObject(ps2, 17, Types.VARCHAR, vo.opt("COMEPOSTIONTIME"));
			DbUtil.setObject(ps2, 18, Types.VARCHAR, vo.opt("REASON"));
			DbUtil.setObject(ps2, 19, Types.VARCHAR, vo.opt("WATCHSTARTDATE"));
			DbUtil.setObject(ps2, 20, Types.VARCHAR, vo.opt("WATCHENDDATE"));
			DbUtil.setObject(ps2, 21, Types.VARCHAR, vo.opt("WATCHDEPT"));
			DbUtil.setObject(ps2, 22, Types.VARCHAR, vo
					.opt("WATCHMASTERWORKER"));
			DbUtil.setObject(ps2, 23, Types.VARCHAR, vo.opt("WATCHSCORE"));
			DbUtil.setObject(ps2, 24, Types.VARCHAR, vo.opt("WATCHRESULT"));
			DbUtil.setObject(ps2, 25, Types.VARCHAR, vo
					.opt("WATCHMASTEREVALUATION"));
			DbUtil.setObject(ps2, 26, Types.VARCHAR, vo.opt("TRAINSTARTDATE"));
			DbUtil.setObject(ps2, 27, Types.VARCHAR, vo.opt("TRAINENDDATE"));
			DbUtil.setObject(ps2, 28, Types.VARCHAR, vo.opt("TRAINSCORE"));
			DbUtil.setObject(ps2, 29, Types.VARCHAR, vo.opt("TRAINRESULT"));
			DbUtil.setObject(ps2, 30, Types.VARCHAR, vo.opt("TRAINEVALUATION"));
			DbUtil.setObject(ps2, 31, Types.VARCHAR, vo.opt("TRAINDEPT"));
			DbUtil.setObject(ps2, 32, Types.VARCHAR, vo.opt("TRAINPOST"));
			DbUtil.setObject(ps2, 33, Types.VARCHAR, request
					.getAttribute("sys.userName"));

			ps2.execute();
			ps2.close();
			// 关闭资源

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}

	/**
	 * 任务分解表
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_SBT_CUST_EMPLOYEE_TARGET(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;

		// 先循环删除所有相同的方案号

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// 先删除月度的数据
			sql = "DELETE wb_erp.SBT_CUST_EMPLOYEE_TARGET_month where schemno=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SCHEMNO"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);

			// 再删除主表数据
			sql = "DELETE wb_erp.SBT_CUST_EMPLOYEE_TARGET where schemno=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SCHEMNO"));
			ps1.executeUpdate();
			conn.commit();
			DbUtil.closeStatement(ps1);
			ps1.close();

		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PreparedStatement ps2 = null;
			// 校验销量取数口径
			sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
			ps2 = conn.prepareStatement(sql);
			DbUtil.setObject(ps2, 1, Types.VARCHAR, vo.get("INVBASDOCID"));
			rSet = ps2.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("INVBASDOCID").toString()
						+ "销量考核取数口径不存在，请检查后重新导入！");
			}
			ps2.close();
			rSet.close();

			// 校验销量取数口径
			if (!vo.get("INVBASDOCID_OTHER").toString().equals("")) {
				System.out.print(vo.get("INVBASDOCID_OTHER"));
				sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
				ps2 = conn.prepareStatement(sql);
				DbUtil.setObject(ps2, 1, Types.VARCHAR, vo
						.get("INVBASDOCID_OTHER"));
				rSet = ps2.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + 1 + "行"
							+ vo.get("INVBASDOCID_OTHER").toString()
							+ "销量考核取数口径不存在，请检查后重新导入！");
				}
			}
			ps2.close();
			rSet.close();

			sql = "insert into wb_erp.SBT_CUST_EMPLOYEE_TARGET (ID,SCHEMNO,	SCHEMNAME,	ZHANGQ,	YXB,	CUSTCODE,	CUSTNAME,	LEVEL_CUST,	CUSTTYPE,	INVBASDOCID,	INVBASDOCID_OTHER,	SALETYPE,	LEVEL_WD) "
					+ " values (?,?,?,	?,	?,	?,	?,	?,	?,	(select id from wb_erp.sbt_cust_invbasdocgroup a where  a.nvbasdocgroup=?),	(select id from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?),	?,?)";

			ps2 = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			// 赋值至子表的主外键
			String FK_ID = PK_ID;
			ps2.setString(1, PK_ID);
			DbUtil.setObject(ps2, 2, Types.VARCHAR, vo.opt("SCHEMNO"));
			DbUtil.setObject(ps2, 3, Types.VARCHAR, vo.opt("SCHEMNAME"));
			DbUtil.setObject(ps2, 4, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps2, 5, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps2, 6, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps2, 7, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps2, 8, Types.VARCHAR, vo.opt("LEVEL_CUST"));
			DbUtil.setObject(ps2, 9, Types.VARCHAR, vo.opt("CUSTTYPE"));
			DbUtil.setObject(ps2, 10, Types.VARCHAR, vo.opt("INVBASDOCID"));
			DbUtil.setObject(ps2, 11, Types.VARCHAR, vo
					.opt("INVBASDOCID_OTHER"));
			DbUtil.setObject(ps2, 12, Types.VARCHAR, vo.opt("SALETYPE"));
			DbUtil.setObject(ps2, 13, Types.VARCHAR, vo.opt("LEVEL_WD"));
			ps2.execute();
			ps2.close();
			// 关闭资源

			// 循环三次生成季度第一月、二月、三月的任务分解表
			sql = "insert into wb_erp.SBT_CUST_EMPLOYEE_TARGET_month (ID ,SCHEMNO, SCHEMNAME, ZHANGQ        ,YXB          ,CUSTCODE      , CUSTNAME    ,  FK_ID, TARGET_MONTH  , TARGET_SALESTR , TARGET_PSNCODE , TARGET_PSNNAME) "
					+ " values (? ,?, ?, ? ,?  ,? , ?  ,  ?, ?  , ? , ? , ? )";
			// 第一个月
			ps2 = conn.prepareStatement(sql);
			// 重新生成主键
			PK_ID = SysUtil.getId();
			ps2.setString(1, PK_ID);
			DbUtil.setObject(ps2, 2, Types.VARCHAR, vo.opt("SCHEMNO"));
			DbUtil.setObject(ps2, 3, Types.VARCHAR, vo.opt("SCHEMNAME"));
			DbUtil.setObject(ps2, 4, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps2, 5, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps2, 6, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps2, 7, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps2, 8, Types.VARCHAR, FK_ID);
			DbUtil.setObject(ps2, 9, Types.VARCHAR, vo.opt("TARGET_MONTH1"));
			DbUtil.setObject(ps2, 10, Types.VARCHAR, vo.opt("TARGET_SALESTR1"));
			DbUtil.setObject(ps2, 11, Types.VARCHAR, vo.opt("TARGET_PSNCODE1"));
			DbUtil.setObject(ps2, 12, Types.VARCHAR, vo.opt("TARGET_PSNNAME1"));
			ps2.execute();
			ps2.close();
			// 关闭资源

			// 第二个月
			ps2 = conn.prepareStatement(sql);
			// 重新生成主键
			PK_ID = SysUtil.getId();
			ps2.setString(1, PK_ID);
			DbUtil.setObject(ps2, 2, Types.VARCHAR, vo.opt("SCHEMNO"));
			DbUtil.setObject(ps2, 3, Types.VARCHAR, vo.opt("SCHEMNAME"));
			DbUtil.setObject(ps2, 4, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps2, 5, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps2, 6, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps2, 7, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps2, 8, Types.VARCHAR, FK_ID);
			DbUtil.setObject(ps2, 9, Types.VARCHAR, vo.opt("TARGET_MONTH2"));
			DbUtil.setObject(ps2, 10, Types.VARCHAR, vo.opt("TARGET_SALESTR2"));
			DbUtil.setObject(ps2, 11, Types.VARCHAR, vo.opt("TARGET_PSNCODE2"));
			DbUtil.setObject(ps2, 12, Types.VARCHAR, vo.opt("TARGET_PSNNAME2"));
			ps2.execute();
			ps2.close();
			// 第三个月
			ps2 = conn.prepareStatement(sql);
			// 重新生成主键
			PK_ID = SysUtil.getId();
			ps2.setString(1, PK_ID);
			DbUtil.setObject(ps2, 2, Types.VARCHAR, vo.opt("SCHEMNO"));
			DbUtil.setObject(ps2, 3, Types.VARCHAR, vo.opt("SCHEMNAME"));
			DbUtil.setObject(ps2, 4, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps2, 5, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps2, 6, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps2, 7, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps2, 8, Types.VARCHAR, FK_ID);
			DbUtil.setObject(ps2, 9, Types.VARCHAR, vo.opt("TARGET_MONTH3"));
			DbUtil.setObject(ps2, 10, Types.VARCHAR, vo.opt("TARGET_SALESTR3"));
			DbUtil.setObject(ps2, 11, Types.VARCHAR, vo.opt("TARGET_PSNCODE3"));
			DbUtil.setObject(ps2, 12, Types.VARCHAR, vo.opt("TARGET_PSNNAME3"));
			ps2.execute();
			ps2.close();
			// 关闭资源
			// 提交事务
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			conn.commit();
			System.out.println(f);

		}
		conn.commit();
		DbUtil.closeConnection(conn);
	}

	/**
	 * 片组奖励分配比例
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_cust_employee_standrad(List<JSONObject> voList,
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
		// 先循环删除所有相同的方案号

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// 先删除

			sql = "DELETE WB_ERP.sbt_cust_employee_standrad where SECHNO=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SECHNO"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		conn.commit();

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			if (Double.parseDouble(vo.get("XS").toString()) < 0) {
				throw new Exception("分配比例不允许小于0");
			}
			if (Double.parseDouble(vo.get("STATUS").toString()) < 0) {
				throw new Exception("在岗系数不允许小于0");
			}
			if (Double.parseDouble(vo.get("BL").toString()) < 0) {
				throw new Exception("预发比例不允许小于0");
			}

			if (Double.parseDouble(vo.get("XS").toString()) > 1) {
				throw new Exception("分配比例不允许大于1");
			}
			if (Double.parseDouble(vo.get("STATUS").toString()) > 1) {
				throw new Exception("在岗系数不允许大于1");
			}
			if (Double.parseDouble(vo.get("BL").toString()) > 1) {
				throw new Exception("预发比例不允许大于1");
			}

			sql = "insert into wb_erp.sbt_cust_employee_standrad (ID,SECHNO,	SECHNAME,	JSMONTH,	ZHANGQ,	YXB,	SALESTR,	PSNCODE,	PSNNAME,	STATUS,	XS,	BL) "
					+ " values (?,?,?,	?,	?,	?,	?,	?,	?,	?,	?,	?)";

			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SECHNO"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("SECHNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("JSMONTH"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("SALESTR"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PSNCODE"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("PSNNAME"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("STATUS"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("XS"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("BL"));

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
	 * 客户主管-奖励方案1
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_SBT_CUST_EMPLOYEE_SCHEME(List<JSONObject> voList,
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
		// 先循环删除所有相同的方案号

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// 先删除

			sql = "DELETE from  WB_ERP.SBT_CUST_EMPLOYEE_SCHEME  where SECHNO=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SECHNO"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		conn.commit();
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			System.out.println(f);
			sql = "insert into WB_ERP.SBT_CUST_EMPLOYEE_SCHEME (ID,SECHNO,	SECHNAME,	JSDATE,	ZHANGQ,	YXB,	SALESTR,	UNIT,	STANDARD,	BASENUMBER,	HIGHERNUMBER,	LOWERNUMBER,	BASEGDPRICE,	BASEPRICE,	XS,	HIGHERPRICE,	HIGHERMAXPRICE) "
					+ " values (?,?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?)";

			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SECHNO"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("SECHNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("JSDATE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("SALESTR"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("UNIT"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("STANDARD"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("BASENUMBER"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("HIGHERNUMBER"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("LOWERNUMBER"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("BASEGDPRICE"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("BASEPRICE"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("XS"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("HIGHERPRICE"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("HIGHERMAXPRICE"));

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
	 * 客户销量与利润目标
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_saletarget_customer(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response,
			String imptype)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		int result2 = 0;
		String strcusttype = "";
		String strcpxx = "";
		String strcustcode = "";
		String strzhangq = "";
		String stryxb = "";
		ResultSet rSet = null;
		PreparedStatement ps1 = null;

		// 判断是否有审核，审核的数据均不允许修改
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			strcusttype = vo.get("CUSTTYPE").toString().trim();

			sql = "select status as CT from wb_erp.saletarget_customer a  where monthyear=? and rownum=1";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTHYEAR"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 1) {
				throw new Exception("除标题外第" + f + "行所在销量期间已审核，不允许新增或修改！");
			}
			DbUtil.closeStatement(ps1);
		}
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			strcusttype = vo.get("CUSTTYPE").toString().trim();
			strcustcode = vo.get("CUSTCODE").toString().trim();
			strzhangq = vo.get("ZHANGQ").toString().trim();
			stryxb = vo.get("YXB").toString().trim();
			strcpxx = vo.get("CPXX").toString().trim();

			if (strcusttype.equals("潜力型价值客户") || strcusttype.equals("维护型价值客户")
					|| strcusttype.equals("一般客户")
					|| strcusttype.equals("新开价值客户")) {
				System.out.println(f + "校验通过");
				if (strcusttype.equals("潜力型价值客户")
						|| strcusttype.equals("维护型价值客户")) {
					if (strcustcode.equals("")) {
						throw new Exception("第" + f + "行客户编码不允许为空");
					} else {
						sql = "select count(1) as CT from wb_erp.bd_custcode_view  where custcode1=?";
						ps1 = conn.prepareStatement(sql);
						DbUtil.setObject(ps1, 1, Types.VARCHAR, vo
								.get("CUSTCODE"));
						rSet = ps1.executeQuery();
						if (rSet.next()) {
							result2 = rSet.getInt("CT");
						}
						if (result2 == 0) {
							throw new Exception("除标题外第" + f + 1 + "行"
									+ vo.get("CUSTCODE").toString()
									+ "一级客户编码不存在，请检查后重新导入！");
						}
						DbUtil.closeStatement(ps1);
					}
				}
			}

			else {
				System.out.print(strcusttype);
				throw new Exception("第" + f + "行客户分类校验不通过");
			}
			if (!strzhangq.equals("") && !stryxb.equals("")) {
				sql = "select count(1) as CT from WB_ERP.org_orgs a  where name  in (?,?)";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ZHANGQ"));
				DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("YXB"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 != 2) {
					throw new Exception("除标题外第" + f + 1
							+ "行战区或营销部不存在，请检查后重新导入！");
				}
				DbUtil.closeStatement(ps1);
			} else {
				throw new Exception("第" + f + "行战区或营销部不允许为空");
			}

			if (strcpxx.equals("分销") || strcpxx.equals("直销")
					|| strcpxx.equals("禽料") || strcpxx.equals("鱼料")
					|| strcpxx.equals("OEM") || strcpxx.equals("预混料"))
				System.out.println("校验通过");
			else
				throw new Exception("第" + f + "行产品行销校验不通过");

			// 先删除
			if ("10".equals(imptype))
				sql = "DELETE WB_ERP.saletarget_customer WHERE custcode = ? and MONTHYEAR =? and cpxx=? and salestr=?";
			if ("11".equals(imptype))
				sql = "DELETE WB_ERP.costtarget_customer WHERE custcode = ? and MONTHYEAR =? and cpxx=? and salestr=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("CUSTCODE"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("MONTHYEAR"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("CPXX"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("SALESTR"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
			conn.commit();
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			if (vo.get("CUSTTYPE").toString().trim().equals("新开价值客户")) {
				if (vo.get("SALESTR").toString().trim().equals("")) {
					throw new Exception("第" + f + "行：新开价值客户的片组为必填项，请填写后重新导入！");
				}
			}
			if ("10".equals(imptype))
				sql = "insert into WB_ERP.saletarget_customer (PK_SALETARGET, CUSTCODE, MONTHYEAR, TARGET , COPERATOR,ZHANGQ,YXB,CUSTTYPE,MONEYOFPROMTION,MONEYOFBANK,CPXX,TS,SALESTR) "
						+ " values (?, ?, ?, ?, ?,?,?,?,?,?,?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?)";
			if ("11".equals(imptype))
				sql = "insert into WB_ERP.costtarget_customer (PK_SALETARGET, CUSTCODE, MONTHYEAR, TARGET , COPERATOR,ZHANGQ,YXB,CUSTTYPE,MONEYOFPROMTION,MONEYOFBANK,CPXX,ts) "
						+ " values (?, ?, ?, ?, ?,?,?,?,?,?,?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'))";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("MONTHYEAR"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("TARGET"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("CUSTTYPE"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("MONEYOFPROMTION"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("MONEYOFBANK"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("CPXX"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("SALESTR"));
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
	 * 特殊客户-不结算列表
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_cust_specialcust(List<JSONObject> voList,
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
		// 先循环删除所有相同的方案号
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// 先删除
			sql = "DELETE WB_ERP.sbt_cust_specialcust WHERE custcode = ? and startmonth =? and endmonth=? and type=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("CUSTCODE"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("STARTMONTH"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ENDMONTH"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("TYPE"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		conn.commit();
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			sql = "insert into WB_ERP.sbt_cust_specialcust (ID, CUSTCODE, CUSTNAME, STARTMONTH, ENDMONTH, RATE, TYPE, TS, COPERATOR) "
					+ " values (?, ?, ?, ?, ?, ?, ?, TO_CHAR(SYSDATE,'yyyy-mm-dd hh24:mi:ss'), ?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("STARTMONTH"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ENDMONTH"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("RATE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, request
					.getAttribute("sys.userName"));

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
	 * 客户年终奖
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_cust_ht_year(List<JSONObject> voList,
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
		// 先循环删除所有相同的方案号
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// 先删除
			sql = "DELETE WB_ERP.sbt_cust_ht_year WHERE htno = ?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("HTNO"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		conn.commit();
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// 校验存货组合
			sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVBADOCGROUP"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("INVBADOCGROUP").toString()
						+ "存货组合不存在，请检查后重新导入！");
			}
			DbUtil.closeStatement(ps1);
			// 校验销量标吨是否存在
			if (!vo.get("SALETYPE").equals("销量")
					&& !vo.get("SALETYPE").equals("标吨")) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("SALETYPE").toString() + "销量类型不存在，请检查后重新导入！");
			}

			// System.out.println(f);
			// 校验标准是否存在

			sql = "select  count(1) as CT FROM WB_ERP.SBT_CUST_STANDRAD B WHERE B.STANDRADNAME=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SCHEMEGROUP"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("SCHEMEGROUP").toString()
						+ "奖励标准不存在，请检查后重新导入！");
			}
			DbUtil.closeStatement(ps1);

			sql = "INSERT INTO  WB_ERP.sbt_cust_ht_year(id,HTNO,  HTTYPE,  CUSTCODE,  CUSTNAME,  STARTDATE,  ENDDATE,  INVBADOCGROUP,  SCHEMEGROUP,  TYPE0,  TYPE7,  TYPE6,  TYPE5,  TYPE1,  TYPE2,  TYPE3,  TYPE4,	TYPE8,	TYPE9,	TYPE10,	SALETYPE) "
					+ " VALUES(?,?,	?,	?,	?,	?,	?,	(select id from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?),	(select  id FROM WB_ERP.SBT_CUST_STANDRAD B WHERE B.STANDRADNAME=? and rownum=1),	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("HTNO"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("HTTYPE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("STARTDATE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ENDDATE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("INVBADOCGROUP"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("SCHEMEGROUP"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("TYPE0"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("TYPE7"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("TYPE6"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("TYPE5"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("TYPE1"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("TYPE2"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("TYPE3"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("TYPE4"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("TYPE8"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("TYPE9"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("TYPE10"));
			DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("SALETYPE"));

			/*
			 * DbUtil.setObject(ps, 7, Types.VARCHAR, request
			 * .getAttribute("sys.userName"));
			 */

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
	 * 专销合同签定情况表
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_cust_ht_sign(List<JSONObject> voList,
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
		// 先循环删除所有相同的方案号
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// 先删除
			sql = "DELETE wb_erp.sbt_cust_ht_sign WHERE htno = ?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("HTNO"));
			ps1.executeUpdate();
			conn.commit();
			DbUtil.closeStatement(ps1);
			ps1.close();
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// 校验合同是否存在

			sql = "select count(1) as CT from wb_erp.sbt_cust_ht where HTNO=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("HTNO"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("HTNO").toString() + "合同编号不存在，请检查后重新导入！");
			}
			ps1.close();
			rSet.close();

			sql = "insert into  WB_ERP.sbt_cust_ht_sign(id,HTNO,  HTTYPE,  CUSTCODE,  CUSTNAME,  STARTDATE,  ENDDATE,  TYPE1,  TYPE2,	TYPE3,	TYPE4,	TYPE5,	TYPE6,TYPE7,TYPE8) "
					+ " values(?,?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("HTNO"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("HTTYPE"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("STARTDATE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ENDDATE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("TYPE1"));
			/*
			 * DbUtil.setObject(ps, 7, Types.VARCHAR, request
			 * .getAttribute("sys.userName"));
			 */
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("TYPE2"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("TYPE3"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("TYPE4"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("TYPE5"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("TYPE6"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("TYPE7"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("TYPE8"));
			ps.execute();
			// 提交事务
			// 关闭资源
			// DbUtil.closeStatement(ps1);
			ps.close();
			DbUtil.closeStatement(ps);
			System.out.println(f);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 专销合同
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_SBT_CUST_HT(List<JSONObject> voList,
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
		// 先循环删除所有相同的方案号
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// 先删除
			sql = "DELETE WB_ERP.SBT_CUST_HT WHERE HTNO = ?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("HTNO"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// 校验存货组合
			sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVBADOCGROUP"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("PK_PRODLINE").toString()
						+ "存货组合不存在，请检查后重新导入！");
			}
			// 校验标准是否存在

			sql = "select  count(1) as CT FROM WB_ERP.SBT_CUST_STANDRAD B WHERE B.STANDRADNAME=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SCHEMEGROUP"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("SCHEMEGROUP").toString()
						+ "奖励标准不存在，请检查后重新导入！");
			}

			// 校验公司是否存在
			if (!vo.get("PK_CORP").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("PK_CORP"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行不存在，请检查后重新导入！");
				}
			}

			sql = "insert into wb_erp.sbt_cust_ht (ID, HTNO, JSZQ, PK_CORP, INVBADOCGROUP, SCHEMEGROUP, JSFS, TS, COPERATOR,SALETYPE) "
					+ " values (?,?, ?, (select pk_corp FROM wb_erp.bd_corp p where p.memo=?), (select id from WB_ERP.sbt_cust_invbasdocgroup b  where b.nvbasdocgroup=?), (select id from WB_ERP.sbt_cust_standrad b  where b.standradname=? and rownum=1), ?, to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'), ?,?)";

			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("HTNO"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("JSZQ"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("PK_CORP"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("INVBADOCGROUP"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("SCHEMEGROUP"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("JSFS"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("SALETYPE"));
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
	 * 客户奖励方案
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_SBT_CUST_SCHEME(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		String strtmp = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;
		// 先循环删除所有相同的方案号
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// 先删除
			sql = "DELETE wb_erp.sbt_cust_scheme WHERE schemeno = ?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SCHEMENO"));
			ps1.executeUpdate();
			conn.commit();
			DbUtil.closeStatement(ps1);
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// 客户组合

			sql = "select count(1) as CT from wb_erp.SBT_CUST_CUSTGROUP A WHERE A.CUSTGROUPNAME =?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("CUSTOMERGROUP"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("CUSTOMERGROUP").toString()
						+ "客户组合不存在，请检查后重新导入！");
			}
			// 校验存货组合
			sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVBADOCGROUP"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("INVBADOCGROUP").toString()
						+ "存货组合不存在，请检查后重新导入！");
			}

			// 校验存货组合1
			if (!vo.get("INVBADOCGROUP1").equals("")) {
				sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo
						.get("INVBADOCGROUP1"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + 1 + "行"
							+ vo.get("INVBADOCGROUP1").toString()
							+ "产品库1名称不存在，请检查后重新导入！");
				}
			}

			// 校验存货组合2
			if (!vo.get("INVBADOCGROUP2").equals("")) {
				sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo
						.get("INVBADOCGROUP2"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + 1 + "行"
							+ vo.get("INVBADOCGROUP2").toString()
							+ "产品库2名称不存在，请检查后重新导入！");
				}
			}
			// 校验存货组合3
			if (!vo.get("INVBADOCGROUP3").equals("")) {
				sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo
						.get("INVBADOCGROUP3"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + 1 + "行"
							+ vo.get("INVBADOCGROUP3").toString()
							+ "产品库3名称不存在，请检查后重新导入！");
				}
			}
			// 校验存货组合4
			if (!vo.get("INVBADOCGROUP4").equals("")) {
				sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo
						.get("INVBADOCGROUP4"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + 1 + "行"
							+ vo.get("INVBADOCGROUP4").toString()
							+ "产品库4名称不存在，请检查后重新导入！");
				}
			}

			// 校验销量标吨是否存在
			if (!vo.get("DJTYPE").equals("销量")
					&& !vo.get("DJTYPE").equals("标吨")) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("DJTYPE").toString() + "定级标准类型不存在，请检查后重新导入！");
			}

			// 校验客户计量方式
			if (!vo.get("CONDITIONJLFS").equals("")
					&& !vo.get("CONDITIONJLFS").equals("猪料总量")
					&& !vo.get("CONDITIONJLFS").equals("按产品组合执行")
					&& !vo.get("CONDITIONJLFS").equals("禽料总量")
					&& !vo.get("CONDITIONJLFS").equals("猪料总量加预混料")
					&& !vo.get("CONDITIONJLFS").equals("猪料总量加S4011")
					&& !vo.get("CONDITIONJLFS").equals("直销猪料")
					&& !vo.get("CONDITIONJLFS").equals("直配猪料")

			) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("CONDITIONJLFS").toString()
						+ "计量方式不存在，请检查后重新导入！");
			}

			// 校验标准是否存在

			sql = "select  count(1) as CT FROM WB_ERP.SBT_CUST_STANDRAD B WHERE B.STANDRADNAME=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SCHEMEGROUP"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("SCHEMEGROUP").toString()
						+ "奖励标准不存在，请检查后重新导入！");
			}

			/*
			 * 校验四个计量方式
			 */
			for (int i = 0; i < 5; i++) {
				if (i != 3 && i != 6) {
					strtmp = vo.get("CONDITION" + i).toString();
					if (!strtmp.equals("")) {
						if (!strtmp.equals("计量计奖") && !strtmp.equals("计量不计奖")
								&& !strtmp.equals("不计量不计奖")) {
							throw new Exception("除标题外第" + f + 1 + "行"
									+ vo.get("CONDITION" + i).toString()
									+ "不存在");

						}
					}
				}
			}

			// 校验公司是否存在
			if (!vo.get("PK_CORP").toString().equals("")) {
				sql = "select  count(1) as CT FROM WB_ERP.bd_corp B WHERE memo=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("PK_CORP"));
				rSet = ps1.executeQuery();
				if (rSet.next()) {
					result2 = rSet.getInt("CT");
				}
				if (result2 == 0) {
					throw new Exception("除标题外第" + f + "行公司不存在，请检查后重新导入！");
				}
			}

			sql = "insert into wb_erp.sbt_cust_scheme(ID,SCHEMENO,  SCHEMENAME,  PK_CORP,  TYPE,  STARTDATE,  ENDDATE,  DETAILNAME,  CUSTOMERGROUP,  SCHEMEGROUP,  DJTYPE,  INVBADOCGROUP,  INVBADOCGROUP1,  INVBADOCGROUP2,  INVBADOCGROUP3,  INVBADOCGROUP4,  CONDITIONJLFS,  CONDITION0,  CONDITION1,  CONDITION2,  CONDITION4,	CONDITION5,TS,COPERATOR,flag,STARTDATE_JS,ENDDATE_JS,RATE,RATE_CONDITION) "
					+ "values(?,?,  ?,  (SELECT PK_CORP FROM WB_ERP.BD_CORP WHERE MEMO=? ),  ?,  ?,  ?,  ?,  (SELECT ID FROM WB_ERP.SBT_CUST_CUSTGROUP A WHERE A.CUSTGROUPNAME =?),  (SELECT ID FROM WB_ERP.SBT_CUST_STANDRAD B WHERE B.STANDRADNAME=? and rownum=1),  ?,  (SELECT ID FROM WB_ERP.SBT_CUST_INVBASDOCGROUP B WHERE B.NVBASDOCGROUP=?),  (SELECT ID FROM WB_ERP.SBT_CUST_INVBASDOCGROUP B WHERE B.NVBASDOCGROUP=?),  (SELECT ID FROM WB_ERP.SBT_CUST_INVBASDOCGROUP B WHERE B.NVBASDOCGROUP=?),  (SELECT ID FROM WB_ERP.SBT_CUST_INVBASDOCGROUP B WHERE B.NVBASDOCGROUP=?),  (SELECT ID FROM WB_ERP.SBT_CUST_INVBASDOCGROUP B WHERE B.NVBASDOCGROUP=?),  ?,  ?,  ?,  ?,  ?,	?,TO_CHAR(SYSDATE,'YYYY-MM-DD HH24:MI:SS'),?,0,?,?,?,?)";

			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SCHEMENO"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("SCHEMENAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("PK_CORP"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("STARTDATE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ENDDATE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("DETAILNAME"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("CUSTOMERGROUP"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("SCHEMEGROUP"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("DJTYPE"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("INVBADOCGROUP"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("INVBADOCGROUP1"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("INVBADOCGROUP2"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("INVBADOCGROUP3"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("INVBADOCGROUP4"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("CONDITIONJLFS"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("CONDITION0"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("CONDITION1"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("CONDITION2"));
			DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("CONDITION4"));
			DbUtil.setObject(ps, 22, Types.VARCHAR, vo.opt("CONDITION5"));
			DbUtil.setObject(ps, 23, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 24, Types.VARCHAR, vo.opt("STARTDATE_JS"));
			DbUtil.setObject(ps, 25, Types.VARCHAR, vo.opt("ENDDATE_JS"));
			DbUtil.setObject(ps, 26, Types.VARCHAR, vo.opt("RATE"));
			DbUtil.setObject(ps, 27, Types.VARCHAR, vo.opt("RATE_CONDITION"));
			ps.execute();
			conn.commit();
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
	 * 开发人员奖励结算
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_employee_scheme_kf(List<JSONObject> voList,
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
		// 先循环删除所有相同的方案号
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// 先删除
			sql = "DELETE wb_erp.sbt_employee_scheme_kf WHERE SECHEMNO = ?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SECHEMNO"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// 校验存货产品线

			sql = "select count(1) as CT from wb_erp.bd_prodline a where a.prodlinename=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("PK_PRODLINE"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("PK_PRODLINE").toString()
						+ "存货产品线不存在，请检查后重新导入！");
			}
			// 校验存货组合
			sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("INVBASDOCID"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("PK_PRODLINE").toString()
						+ "存货组合不存在，请检查后重新导入！");
			}
			// 校验销量标吨是否存在
			if (!vo.get("SALETYPE").equals("销量")
					&& !vo.get("SALETYPE").equals("标吨")) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("SALETYPE").toString() + "销量类型不存在，请检查后重新导入！");
			}
			sql = "insert into wb_erp.sbt_employee_scheme_kf (ID, SECHEMNAME, ZHANGQ, YXB, PK_PRODLINE, STARTMONTH, ENDMONTH, COUNTMONTH, SALETYPE, SALESTART, SALEEND, INVBASDOCID, MONEY_YWY, MONEY_ZZ, SECHEMNO)"
					+ " values (?, ?, ?, ?, (select PK_PRODLINE from  wb_erp.bd_prodline where prodlinename=?), ?, ?, ?, ?, ?,?, (select id from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?), ?, ?, ?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SECHEMNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ZHANGQ"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PK_PRODLINE"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("STARTMONTH"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ENDMONTH"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("COUNTMONTH"));
			/*
			 * DbUtil.setObject(ps, 7, Types.VARCHAR, request
			 * .getAttribute("sys.userName"));
			 */
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("SALETYPE"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("SALESTART"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("SALEEND"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("INVBASDOCID"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("MONEY_YWY"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("MONEY_ZZ"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("SECHEMNO"));
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
	 * 开发人员奖励结算稽核条件，三个导入同一个入口
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_employee_scheme_kf_check(
			List<JSONObject> voList, HttpServletRequest request,
			HttpServletResponse response, String imptype)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;
		int result2 = 0;
		ResultSet rSet = null;
		PreparedStatement ps1 = null;

		// 先循环删除所有相同的方案号
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// 先删除
			if (imptype.equals("2"))
				sql = "DELETE wb_erp.sbt_employee_scheme_kf_check1 where SECHEMNO=?";
			if (imptype.equals("3"))
				sql = "DELETE wb_erp.sbt_employee_scheme_kf_check2 where SECHEMNO=?";
			if (imptype.equals("4"))
				sql = "DELETE wb_erp.sbt_employee_scheme_kf_check3 where SECHEMNO=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("SECHEMNO"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			// 校验存货产品线

			sql = "select count(1) as CT from wb_erp.bd_prodline a where a.prodlinename=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("PK_PRODLINE"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("PK_PRODLINE").toString()
						+ "存货产品线不存在，请检查后重新导入！");
			}
			// 校验存货组合
			if (vo.get("PK_INVBASDOCGROUP") != null) {
				if (!vo.get("PK_INVBASDOCGROUP").toString().equals("")) {
					sql = "select count(1) as CT from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?";
					ps1 = conn.prepareStatement(sql);
					DbUtil.setObject(ps1, 1, Types.VARCHAR, vo
							.get("PK_INVBASDOCGROUP"));
					rSet = ps1.executeQuery();
					if (rSet.next()) {
						result2 = rSet.getInt("CT");
					}
					if (result2 == 0) {
						throw new Exception("除标题外第" + f + 1 + "行"
								+ vo.get("PK_PRODLINE").toString()
								+ "存货组合不存在，请检查后重新导入！");
					}
				}
			}
			// 校验销量标吨是否存在
			if (!vo.get("SALETYPE").equals("销量")
					&& !vo.get("SALETYPE").equals("标吨")) {
				throw new Exception("除标题外第" + f + 1 + "行"
						+ vo.get("SALETYPE").toString() + "销量类型不存在，请检查后重新导入！");
			}
			// 是否存在方案
			sql = "select count(1) as CN from wb_erp.sbt_employee_scheme_kf  t where t.zhangq=? and  pk_prodline=(select pk_prodline from wb_erp.bd_prodline where prodlinename=?) and t.startmonth=? and t.endmonth=?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ZHANGQ"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("PK_PRODLINE"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("STARTMONTH"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ENDMONTH"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CN");
			}
			if (result2 == 0) {
				throw new Exception("除标题外第" + f + 1
						+ "行找不到对应的奖励方案，请优先导入开发奖励方案，再导入稽核条件表 ");
			}

			if (imptype.equals("2")) {
				// sbt_employee_scheme_kf_check1 受益期内稽核条件写入数据库 by wenshixian
				// 2017-11-20
				sql = "insert into wb_erp.sbt_employee_scheme_kf_check1 (ID,  SCHEMNAME,ZHANGQ, YXB, PK_PRODLINE, STARTMONTH, ENDMONTH, COUNTNUMBER, PK_INVBASDOCGROUP, SALETYPE, COUNTMONTH, RATE, SECHEMNO)"
						+ " values (?, ?, ?, ?,  (select a.pk_prodline from wb_erp.bd_prodline a where prodlinename=?),?, ?,?, (select id from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?), ?, ?, ?, ?)";
				ps = conn.prepareStatement(sql);
				String PK_ID = SysUtil.getId();
				ps.setString(1, PK_ID);
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SECHEMNAME"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ZHANGQ"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("YXB"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PK_PRODLINE"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("STARTMONTH"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ENDMONTH"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("COUNTMONTH"));
				/*
				 * DbUtil.setObject(ps, 7, Types.VARCHAR, request
				 * .getAttribute("sys.userName"));
				 */
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo
						.opt("PK_INVBASDOCGROUP"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("SALETYPE"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("COUNTMONTH"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("RATE"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("SECHEMNO"));
				ps.execute();
			}
			if (imptype.equals("3")) {
				// sbt_employee_scheme_kf_check2 受益期次月稽核条件写入数据库 by wenshixian
				// 2017-11-20
				sql = "insert into wb_erp.sbt_employee_scheme_kf_check2 (ID,  SCHEMNAME,ZHANGQ, YXB, PK_PRODLINE, STARTMONTH, ENDMONTH, COUNTNUMBER, PK_INVBASDOCGROUP, SALETYPE, COUNTMONTH, RATE, SECHEMNO,LOWERRATE)"
						+ " values (?, ?, ?, ?,  (select a.pk_prodline from wb_erp.bd_prodline a where a.prodlinename=?),?, ?,?, (select id from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?), ?, ?, ?, ?,?)";
				ps = conn.prepareStatement(sql);
				String PK_ID = SysUtil.getId();
				ps.setString(1, PK_ID);
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SECHEMNAME"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ZHANGQ"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("YXB"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PK_PRODLINE"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("STARTMONTH"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ENDMONTH"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("COUNTMONTH"));
				/*
				 * DbUtil.setObject(ps, 7, Types.VARCHAR, request
				 * .getAttribute("sys.userName"));
				 */
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo
						.opt("PK_INVBASDOCGROUP"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("SALETYPE"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("COUNTMONTH"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("RATE"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("SECHEMNO"));
				DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("LOWERRATE"));
				ps.execute();
			}
			if (imptype.equals("4")) {
				// sbt_employee_scheme_kf_check3 受益期次两月稽核条件写入数据库 by wenshixian
				// 2017-11-20
				sql = "insert into wb_erp.sbt_employee_scheme_kf_check3 (ID,  SCHEMNAME,ZHANGQ, YXB, PK_PRODLINE, STARTMONTH, ENDMONTH, COUNTNUMBER, PK_INVBASDOCGROUP, SALETYPE, TYPE, RATE, SECHEMNO,KHNUMBER,MEMO)"
						+ " values (?, ?, ?, ?,  (select a.pk_prodline from wb_erp.bd_prodline a where a.prodlinename=?),?, ?,?, (select id from wb_erp.sbt_cust_invbasdocgroup a where a.nvbasdocgroup=?), ?, ?, ?, ?,?,?)";
				ps = conn.prepareStatement(sql);
				String PK_ID = SysUtil.getId();
				ps.setString(1, PK_ID);
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SECHEMNAME"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ZHANGQ"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("YXB"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PK_PRODLINE"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("STARTMONTH"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ENDMONTH"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("COUNTMONTH"));
				/*
				 * DbUtil.setObject(ps, 7, Types.VARCHAR, request
				 * .getAttribute("sys.userName"));
				 */
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo
						.opt("PK_INVBASDOCGROUP"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("SALETYPE"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("TYPE"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("RATE"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("SECHEMNO"));
				DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("KHNUMBER"));
				DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("MEMO"));
				ps.execute();
			}
			// 提交事务

			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 基础_奖励标准
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_cust_standrad(List<JSONObject> voList,
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
		// 先循环将标准名称唯一值找出来
		List<String> STANDRADNAME = new ArrayList();
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			if (!STANDRADNAME.contains(vo.get("STANDRADNAME").toString())) {
				STANDRADNAME.add(vo.get("STANDRADNAME").toString());
			}
		}
		// 循环标准名称看系统中是否已存在，已存在先删除
		for (int i = 0; i < STANDRADNAME.size(); i++) {
			sql = "DELETE WB_ERP.sbt_cust_standrad WHERE STANDRADNAME ='"
					+ STANDRADNAME.get(i) + "'";
			ps1 = conn.prepareStatement(sql);
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);
		}
		/*
		 * for (int f = 0; f < voList.size(); f++) { JSONObject vo =
		 * voList.get(f); // 先删除 sql =
		 * "DELETE WB_ERP.sbt_cust_specialcust WHERE custcode = ? and startmonth =? and endmonth=? and type=?"
		 * ; ps1 = conn.prepareStatement(sql); DbUtil.setObject(ps1, 1,
		 * Types.VARCHAR, vo.get("CUSTCODE")); DbUtil.setObject(ps1, 2,
		 * Types.VARCHAR, vo.get("STARTMONTH")); DbUtil.setObject(ps1, 3,
		 * Types.VARCHAR, vo.get("ENDMONTH")); DbUtil.setObject(ps1, 4,
		 * Types.VARCHAR, vo.get("TYPE")); ps1.executeUpdate();
		 * DbUtil.closeStatement(ps1); } conn.commit();
		 */
		for (int i = 0; i < STANDRADNAME.size(); i++) {
			String PK_ID = SysUtil.getId();
			for (int f = 0; f < voList.size(); f++) {
				JSONObject vo = voList.get(f);
				if (vo.opt("STANDRADNAME").equals(STANDRADNAME.get(i))) {
					sql = "insert into WB_ERP.sbt_cust_standrad (ID, STANDRADNAME, STARTNUM, ENDNUM, PRICE, TS, COPERATOR) "
							+ " values (?, ?, ?, ?, ?, TO_CHAR(SYSDATE,'yyyy-mm-dd hh24:mi:ss'), ?)";
					ps = conn.prepareStatement(sql);
					ps.setString(1, PK_ID);
					DbUtil.setObject(ps, 2, Types.VARCHAR, vo
							.opt("STANDRADNAME"));
					// DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("TYPE"));
					DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("STARTNUM"));
					DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ENDNUM"));
					DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PRICE"));
					// DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("FLAG"));
					DbUtil.setObject(ps, 6, Types.VARCHAR, request
							.getAttribute("sys.userName"));

					ps.execute();
					// 提交事务

					// 关闭资源
					// DbUtil.closeStatement(ps1);
					DbUtil.closeStatement(ps);
				}
			}
		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}
	
	/**
	 * 预警目标（战区）
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_app_main_yjmb_zq(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		sql = "DELETE wb_erp.app_main_yjmb_zq WHERE MONTH in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_yjmb_zq (zhanq,month,mbnnumber)"
					+ "values (?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("ZHANQ"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("MBNNUMBER"));
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
	 * 预警目标（行销）
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_app_main_yjmb_xx(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		sql = "DELETE wb_erp.app_main_yjmb_xx WHERE MONTH in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_yjmb_xx (cpx,month,mbnnumber)"
					+ "values (?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("MBNNUMBER"));
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
	 * 预警目标（战区）
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_app_main_yjmb_yxb(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		sql = "DELETE wb_erp.app_main_yjmb_yxb WHERE MONTH in "+Times;
		PreparedStatement ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		DbUtil.closeStatement(ps1);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			
			sql = "insert into wb_erp.app_main_yjmb_yxb (zhanq,YXB,month,mbnnumber)"
					+ "values (?,?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("ZHANQ"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("MBNNUMBER"));
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
