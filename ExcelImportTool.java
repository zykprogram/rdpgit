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
public class ExcelImportTool {
	public static void getFile(HttpServletRequest request,
			HttpServletResponse response) throws Exception {

		InputStream in = (InputStream) request.getAttribute("uploadFile");
		String fileName = request.getAttribute("uploadFile__name").toString();
		String fileType = fileName.substring(fileName.lastIndexOf(".") + 1,
				fileName.length());
		String imptype = request.getAttribute("imptype").toString();
		Map<String, String> map = new HashMap<String, String>();
		if ("1".equals(imptype)) { // 市场范围维护 导入

			map.put("人员编码", "USERCODE");
			map.put("姓名", "USERNAME");
			map.put("省", "PROVINCE");
			map.put("市", "CITY");
			map.put("区/县", "AREA");
			map.put("产品线", "PRODUCTLINE");
			map.put("失效日期", "SXRQ"); //因杨建强增加 付辉平 2018-03-26
			map.put("是否失效", "SFSX"); //因杨建强增加 付辉平 2018-03-26
		} 
		else if ("2".equals(imptype)) { // 业务员与客户对应关系维护 导入
			map.put("客户编码", "CUSTCODE");
			map.put("客户姓名", "CUSTNAME");
			map.put("存货产品线", "RELATIONTYPE");
			map.put("营销部", "YXB");
			map.put("小区", "XQ");
			map.put("片组", "SALESTR");
			
			map.put("客户经理编码", "PSNCODE_CUSTMANAGER");
			map.put("客户经理人员姓名", "PSNNAME_CUSTMANAGER");
			
			map.put("受益组长编码", "MANAGERCODE");
			map.put("受益组长", "MANAGERNAME");

			map.put("受益开发代表1编码", "PSNCODE");
			map.put("受益开发代表1", "PSNNAME");
			map.put("开发代表1受益比例", "MEMO1");

			map.put("受益开发代表2编码", "PSNCODE2");
			map.put("受益开发代表2", "PSNNAME2");
			map.put("开发代表2受益比例", "MEMO2");

			map.put("受益开发代表3编码", "PSNCODE3");
			map.put("受益开发代表3", "PSNNAME3");
			map.put("开发代表3受益比例", "MEMO3");

			map.put("技术经理编码", "PSNCODEMANA_TECHNICAL");
			map.put("技术经理姓名", "PSNNAMEMANA_TECHNICAL");
			
			map.put("技术人员1编码", "PSNCODE1_TECHNICAL");
			map.put("技术人员1姓名", "PSNNAME1_TECHNICAL");
			map.put("技术人员1受益比例", "TECHNICAL1_RATE");

			map.put("技术人员2编码", "PSNCODE2_TECHNICAL");
			map.put("技术人员2姓名", "PSNNAME2_TECHNICAL");
			map.put("技术人员2受益比例", "TECHNICAL2_RATE");

			map.put("是否业绩", "TYPE");
			map.put("备注", "MEMO");

			//map.put("老客户激活日期", "MONTH1");
			map.put("老客户帮扶月份", "MONTH2");
			map.put("母猪头数", "MPIG");
			map.put("肉猪头数", "RPIG");
			//
			Connection conn = DbUtil.getConnection();
			PreparedStatement ps1 = null;
			ResultSet rSet = null;
			int result = 0;
			// 查询是否存在
			String sql = "select count(*) CT from wb_erp.APP_ZBXX_USER where MANGERCODE = ?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result = rSet.getInt("CT");
			}
			// 关闭资源
			DbUtil.closeResultSet(rSet);
			DbUtil.closeStatement(ps1);
			DbUtil.closeConnection(conn);
			if (result == 0) {
				throw new Exception("当前登录人，无权限导入！");
			}

		} else if ("3".equals(imptype)) {
			map.put("客户编码", "KHBM");
			map.put("营销部", "YXB");
			map.put("客户姓名", "KHXM");
			map.put("客户电话", "KHDH");
			map.put("注册时间", "ZCSJ");
			map.put("近3月均标吨", "JBD3");
			map.put("近两月标吨", "JBD2");
			map.put("分销贷款余额", "FXDKYE");
			map.put("直销贷款余额", "ZXDKYE");
			map.put("信用卡余额", "XYKYE");
			map.put("分销授信额度", "FXSXED");
			map.put("当前可贷额度", "DQKDED");
		} else if ("4".equals(imptype)) {
			map.put("客户标识", "KHBS");
			map.put("面签银行", "MQYH");
			map.put("大区", "DQ");
			map.put("省区", "SQ");
			map.put("营销部", "YXB");
			map.put("工厂名称", "GCMC");
			map.put("客户类型", "KHLX");
			map.put("客户编码", "KHBM");
			map.put("客户姓名", "KHXM");
			map.put("所在县市", "SZXS");
			map.put("交易总额（万元）", "JYZE");
			map.put("推荐额度（万元）", "TJED");
			map.put("交易额期间", "JYEQJ");
			map.put("电话", "DH");
			map.put("身份证号", "SFZH");
			map.put("合作年限（系统时间）", "HZNX");
			map.put("营业执照号", "YYZZH");
			map.put("现场面签结果", "XCMQJG");
			map.put("原因说明", "YYSM");
			map.put("现场面签人", "XCMQR");
			map.put("面签日期", "MQRQ");
			map.put("材料齐全日期", "CLQQRQ");
			map.put("信贷属性", "XDSX");
			map.put("问题类别", "WTLB");
			map.put("信贷审核说明", "XDSHSM");
			map.put("推荐日期", "TJRQ");
			map.put("推荐函编号", "TJHBH");
			map.put("还款卡号", "HKKH");
			map.put("放款日期", "FKRQ");
			map.put("放款金额", "FKJE");
			map.put("放款利率", "FKLL");
			map.put("放款期数/月", "FKQS");
			map.put("NC收款单日期", "NCSKDRQ");
			map.put("是否电话通知提货", "SFDHTZTH");
			map.put("外呼人", "WHR");
			map.put("外呼日期", "WHRQ");
			map.put("备注", "BZ");
			map.put("应放款时间", "YFKSJ");
			map.put("超出天数", "CCTS");
			map.put("放款耗时", "FKHS");
			map.put("提前还款时间", "TQHKSJ");
			map.put("应还款时间", "YHKSJ");
			map.put("当前余额", "DQYE");
		} else if ("5".equals(imptype)) {
			map.put("客户类型", "KHLX");
			map.put("贷款银行", "DKYH");
			map.put("所属市场", "SSSC");
			map.put("大区", "DQ");
			map.put("基地", "JD");
			map.put("营销部", "YXB");
			map.put("工厂名称", "GCMC");
			map.put("客户编码", "KHBM");
			map.put("客户姓名", "KHXM");
			map.put("客户电话", "KHDH");
			map.put("所在县市", "SZXS");
			map.put("首次提货日期", "SCTHRQ");
			map.put("身份证号", "SFZH");
			map.put("存栏母猪头数", "CLMZTS");
			map.put("养猪成绩/头", "YZCJ");
			map.put("养猪年限/年", "YZNX");
			map.put("推荐额度（万元）", "TJED");
			map.put("系统注册时间", "XTZCSJ");
			map.put("近一年交易额（万元）", "JYNJYE");
			map.put("经产母猪头数", "JCMZTS");
			map.put("存栏肉猪头数", "CLRZTS");
			map.put("贷前银行负债", "DQYHFZ");
			map.put("贷前民间借款", "DQMJFZ");
			map.put("协议挂量吨数（标吨/年）", "XYGLDS");
			map.put("接收OA日期", "JSOARQ");
			map.put("贷审会日期", "DSHRQ");
			map.put("面签人员", "MQRY");
			map.put("核实面签日期", "HSMQRQ");
			map.put("信贷属性", "XDSX");
			map.put("问题类别", "WTLB");
			map.put("信贷审核说明", "XDSHSM");
			map.put("推荐日期", "TJRQ");
			map.put("推荐函编号", "TJHBH");
			map.put("是否补材料", "SFBQCL");
			map.put("材料齐全日期", "CLQQRQ");
			map.put("是否特批", "SFTP");
			map.put("特批分类", "TPFL");
			map.put("特批说明", "TPSM");
			map.put("扣减天数", "KJTS");
			map.put("需求满足天数", "XQMZTS");
			map.put("进度预警", "JDYJ");
			map.put("放款金额", "FKJE");
			map.put("放款日期", "FKRQ");
			map.put("NC下账日期", "NCXZRQ");
			map.put("是否交保证金", "SFJBZJ");
			map.put("放款期数/月", "FKQS");
			map.put("物流商编码", "WLSBM");
			map.put("物流商名称", "WLSMC");
			map.put("电话", "DH");
			map.put("物流商身份证号", "WLSSFZH");
			map.put("银行负债额", "YHFZE");
			map.put("基总", "JZ");
			map.put("部总", "BZ");
			map.put("业务员", "YWY");
			map.put("应还本金日期", "YHBJRQ");
			map.put("实际结清本金日期", "SJJQBJRQ");
			map.put("还款卡号", "HKKH");
			map.put("贷款利率", "DKLL");
			map.put("信贷推进月份", "XDTJYF");

		} else if ("6".equals(imptype)) {
			map.put("办理银行", "BLYH");
			map.put("营销部", "YXB");
			map.put("客户编码", "KHBM");
			map.put("客户姓名", "KHXM");
			map.put("放款金额", "FKJE");
			map.put("销量基数区间", "XLJSQJ");
			map.put("销量基数", "XLJS");
			map.put("增量目标", "ZLMB");
			map.put("当前实际销量", "DQSJXL");
			map.put("全额补息销量", "QEBXXL");

		} else if ("7".equals(imptype)) {
			map.put("补息方案", "BXFA");
			map.put("补息月份", "BXYF");
			map.put("办理银行", "BLYH");
			map.put("营销部", "YXB");
			map.put("工厂名称", "GCMC");
			map.put("客户编码", "KHBM");
			map.put("客户姓名", "KHXM");
			map.put("电话", "DH");
			map.put("身份证号", "SFZH");
			map.put("面签日期", "MQRQ");
			map.put("放款日期", "FKRQ");
			map.put("放款金额", "FKJE");
			map.put("贷款利息", "DKLX");
			map.put("还款完毕时间", "HKWBSJ");
			map.put("面签前一年销量", "MQQYNXL");
			map.put("面签后一年销量", "MQHYNXL");
			map.put("补息率", "BXL");
			map.put("补息金额", "BXJE");
			map.put("补息说明", "BXSM");

		} else if ("8".equals(imptype)) {
			map.put("客户类型", "KHLX");
			map.put("营销部", "YXB");
			map.put("贷款银行", "DKYH");
			map.put("客户编码", "KHBM");
			map.put("客户姓名", "KHXM");
			map.put("客户电话", "KHDH");
			map.put("应还时间", "YHSJ");
			map.put("应还金额", "YHJE");
			map.put("还款卡号", "HKKH");
		} else if ("9".equals(imptype)) {
			map.put("营销部", "YXB");
			map.put("客户编码", "KHBM");
			map.put("客户姓名", "KHXM");
			map.put("前月销量", "QYXL");
			map.put("上月销量", "SYXL");
			map.put("当月销量", "DYXL");
			map.put("贷销比", "DXB");
			map.put("预警原因", "YJYY");
			map.put("解除预警目标", "JCYJMB");
			map.put("预警次数", "YJCS");
		} else if ("10".equals(imptype)) {
			map.put("大区", "DQ");
			map.put("基地", "JD");
			map.put("营销部", "YXB");
			map.put("直销市场", "ZXSC");
			map.put("金融区域", "JRQY");
			map.put("金融细分市场", "JRXFSC");
			map.put("人员编码", "USERCODE");
			map.put("姓名", "USERNAME");
			map.put("挂点领导编码", "GDLDBM");
			map.put("挂点领导", "GDLDXM");
			map.put("金融区域专员编码", "JRZYBM");
			map.put("金融区域专员", "JRZYXM");
		} else if ("11".equals(imptype)) {// 金融逾期信息导入
			map.put("客户类型", "KHLX");
			map.put("客户编码", "KHBM");
			map.put("战区", "ZQ");
			map.put("营销部", "YXB");
			map.put("客户姓名", "KHXM");
			map.put("客户电话", "KHDH");
			map.put("所在县市", "SZXS");
			map.put("办理银行", "BLYH");
			map.put("贷款额度（万元）", "DKED");
			map.put("还款卡号", "HKKH");
			map.put("逾期类别", "YQLB");
			map.put("逾期开始时间", "YQKSSJ");
			map.put("逾期天数", "YQTS");
			map.put("逾期金额", "YQJE");
			map.put("跟进情况", "GJQK");
			map.put("权益缓发核实情况", "QYHFQK");
			map.put("是否追回", "SFZH");
			map.put("结清日期", "JQRQ");
			map.put("追回金额", "ZHJE");
			map.put("逾期余额", "YQYE");
		} else if ("12".equals(imptype)) {// 金融放贷导入
			map.put("推荐函号", "TJHBH");
			map.put("客户编码", "KHBM");
			map.put("客户姓名", "KHXM");
			map.put("放款日期", "FKRQ");
			map.put("放款金额", "FKJE");
			map.put("还款卡号", "HKKH");
			map.put("NC下账日期", "NCXZRQ");

		} else if ("13".equals(imptype)) {// 直销客户主管客户移交
			map.put("战区", "DQ");
			map.put("营销部", "YXB");
			map.put("客户编码", "CUSTCODE");
			map.put("客户姓名", "CUSTNAME");
			map.put("移交/新开", "TRANSFERTYPE");
			map.put("移交月份/新开月份", "MONTH");
			map.put("移交前受益人员片组", "SALESTRUBEFORE");
			map.put("移交前受益人员编码", "PSNCODEBEFORE");
			map.put("移交前受益人员姓名", "PSNNAMEBEFORE");
			map.put("移交后受益人员片组", "SALESTRUAFTER");
			map.put("移交后受益人员编码", "PSNCODEAFTER");
			map.put("移交后受益人员姓名", "PSNNAMEAFTER");

		} else if ("14".equals(imptype))// 客户挂靠关系维护
		{
			map.put("模式分类", "TYPE");
			map.put("客户编码", "CUSTCODE");
			map.put("客户姓名", "CUSTNAME");
			map.put("实际所属经销商编码", "FACTCUSTCODE");
			map.put("实际所属经销商姓名", "FACTCUSTNAME");
		} else if ("15".equals(imptype))// 人员名单/人员组织关系
		{
			map.put("年月", "MONTH");
			// map.put("系统", "SYS");
			// map.put("行政组织", "ORGNAME");
			// map.put("组织简称", "ORGSHORTNAME");
			map.put("人员编码", "PSNCODE");
			map.put("姓名", "PSNNAME");
			// map.put("进入集团时间", "ONJOBTIME");
			map.put("所在部门", "DEPTNAME");
			map.put("岗位", "POSTNAME");
			map.put("离职日期", "OUTTIME");
			map.put("产品线", "CPX");
			map.put("一级部门", "ORG1");
			map.put("二级部门", "ORG2");
			map.put("三级部门", "ORG3");
			map.put("四级分销片", "ORG4");
			map.put("人员状态", "STATUS");
			map.put("新老业务", "BUSSTYPE");
		} else if ("16".equals(imptype) || "161".equals(imptype))// 人力包明细-福利社保
																	// 人力包-薪酬-其他
		{
			map.put("费用归属月份", "MONTH");
			map.put("产品线", "CPX");
			map.put("一级部门", "ORG1");
			map.put("二级部门", "ORG2");
			map.put("三级部门", "ORG3");
			map.put("四级部门", "ORG4");
			map.put("人员编码", "PSNCODE");
			map.put("人员姓名", "PSNNAME");
			map.put("管理一级科目", "KM1");
			map.put("管理二级科目", "KM2");
			map.put("管理三级科目", "KM3");
			map.put("管理四级科目", "KM4");
			map.put("费用金额", "MONEY");
			map.put("费用类型", "TYPE");
			map.put("新老业务", "BUSSTYPE");

		} else if ("17".equals(imptype))// 人力包明细-薪酬
		{
			map.put("费用归属月份", "MONTH");
			map.put("产品线", "CPX");
			map.put("一级部门", "ORG1");
			map.put("二级部门", "ORG2");
			map.put("三级部门", "ORG3");
			map.put("四级部门", "ORG4");
			map.put("人员编码", "PSNCODE");
			map.put("姓名", "PSNNAME");
			map.put("固定工资", "MONEYBASE");
			map.put("绩效奖励", "MONEYKPI");
			map.put("差旅费", "MONEYTRAVEL");
			map.put("市场招待费", "MONEYSERVE1");
			map.put("业务招待费", "MONEYSERVE2");
			map.put("项目奖励", "MONEYPROJECT");
			map.put("战略补贴", "MONEYFILLPOST");
			map.put("奖罚", "MONEYKK");
			map.put("新老业务", "BUSSTYPE");
		} else if ("18".equals(imptype))// 业务包明细-客户
		{
			map.put("费用归属月份", "MONTH");
			map.put("产品线", "CPX");
			map.put("一级部门", "ORG1");
			map.put("二级部门", "ORG2");
			map.put("三级部门", "ORG3");
			map.put("四级部门", "ORG4");
			map.put("客户编码", "CUSTCODE");
			map.put("客户名称", "CUSTNAME");
			map.put("管理一级科目", "KM1");
			map.put("管理二级科目", "KM2");
			map.put("管理三级科目", "KM3");
			map.put("管理四级科目", "KM4");
			map.put("费用金额", "MONEY");
			map.put("费用说明", "MEMO");
			map.put("促销类型", "TYPE");
			map.put("公司", "UNITSHORTNAME");
			map.put("促销存货", "INVNAME");
			map.put("促销销量", "NNUMBER");
			map.put("系统开票类型", "SYSTYPE");
			map.put("系统开票备注", "SYSMEMO");
			map.put("修正大促销编号", "HDMB");
			map.put("新老业务", "BUSSTYPE");
		}

		else if ("19".equals(imptype))// 人力包明细-营销奖金1-开发奖金
		{
			map.put("结算计提月份", "MONTH");
			map.put("销售月份", "XSMONTH");
			map.put("产品线", "CPX");
			map.put("一级部门", "ORG1");
			map.put("二级部门", "ORG2");
			map.put("三级部门", "ORG3");
			map.put("四级部门", "ORG4");
			map.put("客户编码", "CUSTCODE");
			map.put("客户姓名", "CUSTNAME");
			map.put("开发月份", "KFMONTH");
			map.put("开发奖励", "MONEY");
			map.put("受益期后稽核扣回", "MONEYKH");
			map.put("备注", "MEMO");
			map.put("新老业务", "BUSSTYPE");
		} else if ("20".equals(imptype))// 人力包明细-营销奖金2-客户主管
		{
			map.put("结算计提月份", "MONTH");
			map.put("产品线", "CPX");
			map.put("一级部门", "ORG1");
			map.put("二级部门", "ORG2");
			map.put("三级部门", "ORG3");
			map.put("四级部门", "ORG4");
			map.put("跑面片", "SALESTR");
			map.put("人员编码", "PSNCODE");
			map.put("人员姓名", "PSNNAME");
			map.put("客户主管奖励", "MONEY");
			map.put("料型", "TYPE");
			map.put("新老业务", "BUSSTYPE");
		} else if ("21".equals(imptype))// 人力包明细-营销奖金3-其他营销奖金
		{
			map.put("结算计提月份", "MONTH");
			map.put("产品线", "CPX");
			map.put("一级部门", "ORG1");
			map.put("二级部门", "ORG2");
			map.put("三级部门", "ORG3");
			map.put("四级部门", "ORG4");
			map.put("人员编码", "PSNCODE");
			map.put("人员姓名", "PSNNAME");
			map.put("开发月份", "KFMONTH");
			map.put("营销员奖金合计", "SUMMONEY");
			map.put("管理奖励", "MONEYGL");
			map.put("预混料奖励", "MONEYYHL");
			map.put("OEM奖励", "MONEYOEM");
			map.put("驻场技术员奖励", "MONEYZCJS");
			map.put("其他奖励", "MONEYQT");
			map.put("补发补扣", "MONEYBF");
			map.put("新老业务", "BUSSTYPE");
		} else if ("22".equals(imptype))// 管理帐明细
		{
			map.put("费用归属月份", "MONTH");
			map.put("OA号", "REQUESTID");
			map.put("ID号", "WORKFLOWID");
			map.put("流程名称", "REQUESTNAME");
			map.put("发起人编码", "PSNCODE");
			map.put("发起人", "PSNNAME");
			map.put("金额", "MONEY");
			map.put("支出事由", "MEMO");
			map.put("凭证号", "ISTRUE");
			map.put("报账公司", "UNITSHORTNAME");
			map.put("一级科目", "KM1");
			map.put("二级科目", "KM2");
			map.put("三级科目", "KM3");
			map.put("四级科目", "KM4");
			map.put("产品线", "CPX");
			map.put("一级部门", "ORG1");
			map.put("二级部门", "ORG2");
			map.put("三级部门", "ORG3");
			map.put("四级部门", "ORG4");
			map.put("此单据是否取数", "ISBILL");
			map.put("新老业务", "BUSSTYPE");

		} else if ("23".equals(imptype))// 员工培训费
		{
			map.put("费用归属月份", "MONTH");
			map.put("产品线", "CPX");
			map.put("一级部门", "ORG1");
			map.put("二级部门", "ORG2");
			map.put("三级部门", "ORG3");
			map.put("四级部门", "ORG4");
			map.put("管理一级科目", "KM1");
			map.put("管理二级科目", "KM2");
			map.put("管理三级科目", "KM3");
			map.put("管理四级科目", "KM4");
			map.put("费用金额", "MONEY");
			map.put("费用类型", "TYPE");
			map.put("新老业务", "BUSSTYPE");

		} else if ("24".equals(imptype))// 特殊费用-OA调整
		{
			map.put("调整月份", "MONTH");
			map.put("产品线", "CPX");
			map.put("一级部门", "ORG1");
			map.put("二级部门", "ORG2");
			map.put("三级部门", "ORG3");
			map.put("四级部门", "ORG4");
			map.put("管理一级科目", "KM1");
			map.put("管理二级科目", "KM2");
			map.put("管理三级科目", "KM3");
			map.put("管理四级科目", "KM4");
			map.put("费用金额", "MONEY");
			map.put("费用说明", "MEMO");
			map.put("新老业务", "BUSSTYPE");
			map.put("报账公司", "UNITNAME");
			map.put("凭证号", "PZH");
		} else if ("25".equals(imptype))// 试验补贴
		{
			map.put("费用归属月份", "MONTH");
			map.put("产品线", "CPX");
			map.put("一级部门", "ORG1");
			map.put("二级部门", "ORG2");
			map.put("三级部门", "ORG3");
			map.put("四级部门", "ORG4");
			map.put("管理一级科目", "KM1");
			map.put("管理二级科目", "KM2");
			map.put("管理三级科目", "KM3");
			map.put("管理四级科目", "KM4");
			map.put("费用金额", "MONEY");
			map.put("费用类型", "TYPE");
			map.put("新老业务", "BUSSTYPE");

		} else if ("26".equals(imptype))// 其他费用
		{
			map.put("费用归属月份", "MONTH");
			map.put("报账日期", "DBILLDATE");
			map.put("摘要", "MEMO");
			map.put("公司或部门", "DEPT");
			map.put("台帐一级", "KM1TZ");
			map.put("台帐二级", "KM2TZ");
			map.put("借方", "JF");
			map.put("贷方", "DF");
			map.put("新老业务", "BUSSTYPE");
			map.put("产品线", "CPX");
			map.put("一级部门", "ORG1");
			map.put("二级部门", "ORG2");
			map.put("三级部门", "ORG3");
			map.put("四级部门", "ORG4");
			map.put("一级科目", "KM1");
			map.put("二级科目", "KM2");
			map.put("三级科目", "KM3");
			map.put("四级科目", "KM4");
			map.put("金额", "MONEY");

		} else if ("27".equals(imptype))// 部门校验
		{
			map.put("费用归属月份", "MONTH");
			map.put("产品线", "CPX");
			map.put("新老业务", "TYPE");
			map.put("一级部门", "ORG1");
			map.put("二级部门", "ORG2");
			map.put("三级部门", "ORG3");
			map.put("四级部门", "ORG4");
		} else if ("28".equals(imptype))// 人力包明细-薪酬汇总
		{
			map.put("核算年月", "HSMONTH");
			map.put("费用归属月份", "MONTH");
			map.put("产品线", "CPX");
			map.put("一级部门", "ORG1");
			map.put("二级部门", "ORG2");
			map.put("三级部门", "ORG3");
			map.put("四级部门", "ORG4");
			map.put("固定工资", "MONEYBASE");
			map.put("绩效奖励", "MONEYKPI");
			map.put("差旅费", "MONEYTRAVEL");
			map.put("业务招待费", "MONEYSERVE2");
			map.put("项目奖励", "MONEYPROJECT");
			map.put("战略补贴", "MONEYFILLPOST");
			map.put("奖罚", "MONEYKK");
			map.put("新老业务", "BUSSTYPE");
		} else if ("29".equals(imptype))//费用导入
		{
			map.put("费用归属月份", "MONTH");
			map.put("人员姓名", "PSNNAME");
			map.put("人员编码", "PSNCODE");
			map.put("产品行销", "CPXX");
			map.put("一级部门", "ORG1");
			map.put("二级部门", "ORG2");
			map.put("三级部门", "ORG3");
			map.put("四级部门", "ORG4");
			map.put("金额", "Price");
			map.put("管理一级科目", "KM1");
			map.put("管理二级科目", "KM2");
			map.put("管理三级科目", "KM3");
			map.put("管理四级科目", "KM4");
		}else if("30".equals(imptype))
		{
			map.put("组织", "ZZNAME");
			map.put("目标月份","MONTH");
			map.put("目标类型", "TYPE");
			map.put("目标", "TARGET");
			map.put("销售分类", "MEASURE");
			map.put("业务类型", "PRODLINENAME");
		}
		 else if ("31".equals(imptype))//固定
			{
				map.put("费用归属月份", "MONTH");
				map.put("产品线", "CPX");
				map.put("一级部门", "ORG1");
				map.put("二级部门", "ORG2");
				map.put("三级部门", "ORG3");
				map.put("四级部门", "ORG4");
				map.put("金额", "PRICE");
				map.put("管理一级科目", "KM1");
				map.put("管理二级科目", "KM2");
				map.put("管理三级科目", "KM3");
				map.put("管理四级科目", "KM4");
				map.put("凭证号", "PZH");
			}
		 else if("32".equals(imptype)) {
			 map.put("省", "PROVICE");
			 map.put("市", "CITY");
			 map.put("区县", "COUNTY");
			 map.put("养殖场户名称", "NAME");
			 map.put("畜种", "TYPE");
			 map.put("详细地址", "ADDRESS");
			 map.put("联系人", "LINKMAN");
			 map.put("手机", "PHONE");
			 map.put("总存栏", "TOTALCL");
			 map.put("能繁母畜存栏", "MCCL");
			 map.put("出栏", "CL");
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
								if (cell == null&&!"30".equals(request
										.getAttribute("imptype")
										.toString())) {
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
									/*
									 * 客户对应关系导入允许字段为空 by wensx 2016-07-20
									 */
									if (cellVal == null
											&& "2".equals(request.getAttribute(
													"imptype").toString())
											&& (j >= 6 && j!= 30))
										cellVal = "";
									if (cellVal == null
											&& "15".equals(request
													.getAttribute("imptype")
													.toString()) && j == 9)
										cellVal = "";
									if (cellVal == null
											&& "16".equals(request
													.getAttribute("imptype")
													.toString()) && j == 6)
										cellVal = "";
									if (cellVal == null
											&& "30".equals(request
													.getAttribute("imptype")
													.toString()) )
										cellVal = "";
									if (cellVal == null
											&& "2".equals(request.getAttribute(
													"imptype").toString()))
										throw new Exception("第" + i + "行" + j
												+ "列为空，请填写");
								}
						/*	System.out.print(cellVal);
								System.out.println(j + "列");
							System.out.println(i + "行");*/

//								System.out.println(map.get(headRow.getCell(j)
//										.getStringCellValue().toString().trim()
//										.replaceAll("\r|\n", "")));
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
			if ("3".equals(imptype)) {
				imp_app_jr_kde(list, request, response);
			} else if ("4".equals(imptype)) {
				imp_APP_JR_FX(list, request, response);
			} else if ("5".equals(imptype)) {
				imp_APP_JR_ZX(list, request, response);
			} else if ("6".equals(imptype)) {
				imp_APP_JR_TXXLMB(list, request, response);
			} else if ("7".equals(imptype)) {
				imp_APP_JR_TXQK(list, request, response);
			} else if ("8".equals(imptype)) {
				imp_APP_JR_HB(list, request, response);
			} else if ("9".equals(imptype)) {
				imp_APP_JR_YJ(list, request, response);
			} else if ("10".equals(imptype)) {
				imp_APP_JR_GDLDorJRZY(list, request, response);
			} else if ("11".equals(imptype)) {
				imp_APP_JR_YQXX(list, request, response);
			} else if ("12".equals(imptype)) {
				imp_APP_JR_FKXX(list, request, response);
			} else if ("13".equals(imptype)) {
				imp_Customertransfer(list, request, response);
			} else if ("14".equals(imptype)) {
				imp_SuperiorMerchants(list, request, response);
			} else if ("15".equals(imptype)) {
				imp_PersonAndOrg(list, request, response);
			} else if ("16".equals(imptype) || "161".equals(imptype)) {
				//福利社保与人力包-薪酬其他类似，
				imp_DetailsCharges(list, request, response,imptype);
			} else if ("17".equals(imptype) || "28".equals(imptype)) {
				imp_Salarydetails(list, request, response, imptype);
			} else if ("18".equals(imptype)) {
				imp_DetailsChargesCust(list, request, response);
			} else if ("19".equals(imptype)) {
				imp_detailschargekfjl(list, request, response);
			} else if ("20".equals(imptype)) {
				imp_detailschargekhjl(list, request, response);
			} else if ("21".equals(imptype)) {
				imp_detailschargeqtjl(list, request, response);
			} else if ("22".equals(imptype)) {
				imp_TotalCost(list, request, response);
			} else if ("23".equals(imptype)) {
				imp_StaffTraining(list, request, response);
			} else if ("24".equals(imptype)) {
				imp_specialexpenses(list, request, response);
			} else if ("25".equals(imptype)) {
				imp_Testsubsidy(list, request, response);
			} else if ("26".equals(imptype)) {
				imp_DetailsChargestz(list, request, response);
			} else if ("27".equals(imptype)) {
				imp_sbt_deptcheck(list, request, response);
			}else if ("29".equals(imptype)) {
				imp_sbt_costimport(list, request, response);
			}else if("30".equals(imptype)) {
				imp_saletarget(list,request,response);
			}else if("31".equals(imptype)) {
				imp_sbt_guding(list,request,response);
			}else if("32".equals(imptype)) {
				imp_sbt_chda(list,request,response);
			}
			else {
				for (int f = 0; f < list.size(); f++) {
					if ("1".equals(imptype)) {
						imp_scfwwh(list.get(f), request, response);
					} else if ("2".equals(imptype)) {

						imp_CUSTMOERRELATSALESMAN(list.get(f), request,
								response, sBuffer, f + 2);
						System.out.println(f);
					}
				}
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
	
	
	private static void imp_sbt_chda(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		 PreparedStatement ps1 = null;
		int result2 = 0;
		ResultSet rSet = null;
		/*for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// 先删除
				sql = "DELETE wb_erp.sbt_deptcheck WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);	
		}*/
		

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
	


			// System.out.println(f);
			sql = "INSERT INTO wb_erp.app_main_archives "
					+ "(ID, Provice,City, County, Name, Type, Address, Linkman, Phone, Totalcl, MCCL,CL) "
					+ "VALUES " + "(?, ?,?, ?, ?, ?, ?, ?, ?, ?, ?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("PROVICE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("CITY"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("COUNTY"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("NAME"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ADDRESS"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("LINKMAN"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("PHONE"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("TOTALCL"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MCCL"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("CL"));
			ps.execute();
			// 提交事务

			// 关闭资源
	
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 27部门校验
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_deptcheck(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		 PreparedStatement ps1 = null;
		int result2 = 0;
		ResultSet rSet = null;
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// 先删除
				sql = "DELETE wb_erp.sbt_deptcheck WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);	
		}
		

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
	


			// System.out.println(f);
			sql = "insert into wb_erp.sbt_deptcheck (MONTH,  CPX,  ORG1,  ORG2,  ORG3,	ORG4,		TS,	COPERATOR,ID,TYPE) "
					+ "values (?,	?,	?,	?,	?,	?,	to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(8, PK_ID);
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("TYPE"));
			ps.execute();
			// 提交事务

			// 关闭资源
	
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 26其他费用
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_DetailsChargestz(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		// String PK_ID = null;
		String sql = "";
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		PreparedStatement ps1 = null;
		int result2 = 0;
		ResultSet rSet = null;
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// 先删除
				sql = "DELETE wb_erp.SBT_DetailsChargestz WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("第" + f + "行" + vo.get("ORG4").toString()
						+ "部门不存在，请检查后重新导入！注：从0行算第一行");

			}

		
			// System.out.println(f);
			sql = "insert into wb_erp.SBT_DetailsChargestz (MONTH,  CPX,  ORG1,  ORG2,  ORG3,	ORG4,	KM1TZ,	KM2TZ,	KM1,	KM2,	KM3,	KM4,	MONEY,	DEPT,	TS,	COPERATOR,ID,BUSSTYPE,JF,DF,MEMO,DBILLDATE) "
					+ "values (?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("KM1TZ"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("KM2TZ"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("DEPT"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(16, PK_ID);
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("BUSSTYPE"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("JF"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("DF"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("MEMO"));
			DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("DBILLDATE"));

			ps.execute();
			// 提交事务

			// 关闭资源
			 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 25试验补贴
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_Testsubsidy(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// 先删除
				sql = "DELETE wb_erp.sbt_detailsTestsubsidy WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("第" + f + "行"+ "部门不存在，请检查后重新导入！");
			}
	

			// System.out.println(f);
			sql = "insert into wb_erp.sbt_detailsTestsubsidy (MONTH,  CPX,  ORG1,  ORG2,  ORG3,	ORG4,		KM1,	KM2,	KM3,	KM4,	MONEY,	TYPE,	TS,	COPERATOR,ID,BUSSTYPE) "
					+ "values (?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(14, PK_ID);
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("BUSSTYPE"));

			ps.execute();
			// 提交事务

			// 关闭资源
			 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 24特殊费用-OA调整
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_specialexpenses(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// 先删除
				sql = "DELETE wb_erp.sbt_detailsspecialexpenses WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);	
		}
	
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("第" + f + "行" + vo.get("ORG4").toString()
						+ "部门不存在，请检查后重新导入！");
			}
		
			// System.out.println(f);
			sql = "insert into wb_erp.sbt_detailsspecialexpenses (MONTH,  CPX,  ORG1,  ORG2,  ORG3,	ORG4,		KM1,	KM2,	KM3,	KM4,	MONEY,	MEMO,	TS,	COPERATOR,ID,BUSSTYPE,UNITNAME,PZH) "
					+ "values (?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("MEMO"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(14, PK_ID);
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("BUSSTYPE"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("UNITNAME"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("PZH"));
			ps.execute();
			// 提交事务

			// 关闭资源
			 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 23员工培训费
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_StaffTraining(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// 先删除
				sql = "DELETE wb_erp.sbt_detailsStaffTraining WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("第" + f + "行" + vo.get("ORG4").toString()
						+ "部门不存在，请检查后重新导入！");
			}
		

			// System.out.println(f);
			sql = "insert into wb_erp.sbt_detailsStaffTraining (MONTH,  CPX,  ORG1,  ORG2,  ORG3,	ORG4,		KM1,	KM2,	KM3,	KM4,	MONEY,	TYPE,	TS,	COPERATOR,ID,BUSSTYPE) "
					+ "values (?,	?,	?,		?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(14, PK_ID);
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("BUSSTYPE"));

			ps.execute();
			// 提交事务

			// 关闭资源
			 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 22管理账单明细
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_TotalCost(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// 先删除
				sql = "DELETE wb_erp.SBT_DETAILSCHARGEGLZ WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);	
		}
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("第" + f + "行" + vo.get("ORG4").toString()
						+ "部门不存在，请检查后重新导入！");
			}
			DbUtil.closeStatement(ps1);
		
			sql = "insert into WB_ERP.SBT_DETAILSCHARGEGLZ (MONTH, REQUESTID, WORKFLOWID, REQUESTNAME, PSNCODE, PSNNAME, MONEY, MEMO, ISTRUE, UNITSHORTNAME, KM1, KM2, KM3, KM4, CPX, ORG1, ORG2, ORG3, ORG4, ISBILL, TS, ID,COPERATOR,BUSSTYPE) "
					+ "values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,  to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?, ?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("REQUESTID"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("WORKFLOWID"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("REQUESTNAME"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PSNCODE"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("PSNNAME"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("MEMO"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("ISTRUE"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("UNITSHORTNAME"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("ISBILL"));

			String PK_ID = SysUtil.getId();
			ps.setString(21, PK_ID);
			DbUtil.setObject(ps, 22, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 23, Types.VARCHAR, vo.opt("BUSSTYPE"));

			ps.execute();
			// 提交事务

			// 关闭资源
			DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);
			System.out.println(f);
		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 21其他营销奖金
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_detailschargeqtjl(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// 先删除
				sql = "DELETE wb_erp.sbt_detailschargeqtjl WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);
		}
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("第" + f + "行" + vo.get("ORG4").toString()
						+ "部门不存在，请检查后重新导入！");
			}
		

			sql = "insert into wb_erp.sbt_detailschargeqtjl (MONTH, CPX, ORG1, ORG2, ORG3, ORG4, PSNCODE, PSNNAME, SUMMONEY, MONEYGL, MONEYYHL, MONEYZCJS, MONEYOEM, MONEYQT, MONEYBF, TS,COPERATER , ID, MEMO,BUSSTYPE) "
					+ "values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,  to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'), ?,?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PSNCODE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PSNNAME"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("SUMMONEY"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("MONEYGL"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MONEYYHL"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("MONEYZCJS"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("MONEYOEM"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("MONEYQT"));

			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("MONEYBF"));
			// DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("TS"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(17, PK_ID);
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("MEMO"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("BUSSTYPE"));
			ps.execute();
			// 提交事务
			/*System.out.println( vo.opt("PSNCODE").toString());
			if(vo.opt("PSNCODE").toString().equals("008422"))
			{
				System.out.println(vo.opt("PSNCODE").toString());
			}*/
			// 关闭资源
		 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 20客户主管奖励
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_detailschargekhjl(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// 先删除
				sql = "DELETE wb_erp.SBT_DETAILSCHARGEKHZG WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);
		}
	
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=?";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("第" + f + "行" + vo.get("ORG4").toString()
						+ "部门不存在，请检查后重新导入！");
			}
		

			sql = "insert into wb_erp.SBT_DETAILSCHARGEKHZG (MONTH, CPX, ORG1, ORG2, ORG3, ORG4,  PSNCODE, PSNNAME,SALESTR, TS, COPERATOR, ID, MEMO, TYPE, MONEY,BUSSTYPE) "
					+ "values (?, ?, ?, ?, ?, ?, ?, ?, ?, to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'), ?, ?, ?, ?, ?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PSNCODE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PSNNAME"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("SALESTR"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(11, PK_ID);
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("MEMO"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("BUSSTYPE"));
			ps.execute();
			// 提交事务

			// 关闭资源
			 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 19开发奖励
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_detailschargekfjl(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// 先删除
				sql = "DELETE wb_erp.SBT_DETAILSCHARGEkfjl WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);
				
		}
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("第" + f + "行" + vo.get("ORG4").toString()
						+ "部门不存在，请检查后重新导入！");
			}
		
			sql = "insert into wb_erp.SBT_DETAILSCHARGEkfjl (MONTH, CPX, ORG1, ORG2, ORG3, ORG4, CUSTCODE, CUSTNAME, KFMONTH, MONEY, MONEYKH, TS, COPERATOR, ID, MEMO,XSMONTH,BUSSTYPE) "
					+ "values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'), ?, ?, ?, ?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KFMONTH"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MONEYKH"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(13, PK_ID);
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("MEMO"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("XSMONTH"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("BUSSTYPE"));

			ps.execute();
			// 提交事务

			// 关闭资源
			 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 18业务包明细-客户
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_DetailsChargesCust(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

				// 先删除
				sql = "DELETE wb_erp.sbt_detailschargesCust WHERE month = ? ";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
				conn.commit();
				DbUtil.closeStatement(ps1);	
		}
	
		
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;

			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("第" + f + "行" + vo.get("ORG4").toString()
						+ "部门不存在，请检查后重新导入！");
			}

		

			sql = "insert into wb_erp.sbt_detailschargesCust (MONTH, CPX, ORG1, ORG2, ORG3, ORG4, CUSTCODE, CUSTNAME, KM1, KM2, KM3, KM4, MONEY, TYPE, TS, COPERATOR, ID, MEMO, UNITSHORTNAME, INVNAME, NNUMBER, SYSTYPE, SYSMEMO, HDMB,BUSSTYPE)"
					+ "values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'), ?, ?, ?, ?, ?, ?, ?, ?, ?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("CUSTCODE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("CUSTNAME"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(16, PK_ID);
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("MEMO"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("UNITSHORTNAME"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("INVNAME"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("NNUMBER"));
			DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("SYSTYPE"));
			DbUtil.setObject(ps, 22, Types.VARCHAR, vo.opt("SYSMEMO"));
			DbUtil.setObject(ps, 23, Types.VARCHAR, vo.opt("HDBM"));
			DbUtil.setObject(ps, 24, Types.VARCHAR, vo.opt("BUSSTYPE"));

			ps.execute();
			// 提交事务

			// 关闭资源
			 DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 17薪酬明细
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_Salarydetails(List<JSONObject> voList,
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

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			if ( imptype.equals("17")) {
				// 先删除
				sql = "DELETE wb_erp.sbt_Salarydetails WHERE month = ?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
			}
			if (imptype.equals("28")) {
				// 先删除
				sql = "DELETE wb_erp.sbt_Salarydetails_total WHERE month = ?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				ps1.executeUpdate();
			}
			conn.commit();
			DbUtil.closeStatement(ps1);
		}

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("第" + f + "行" + vo.get("ORG4").toString()
						+ "部门不存在，请检查后重新导入！");
			}
			
			DbUtil.closeStatement(ps1);

			if ("28".equals(imptype)) {
				sql = "insert into wb_erp.sbt_Salarydetails_total (HSMONTH,MONTH,	CPX,	ORG1,	ORG2,	ORG3,	ORG4,	MONEYBASE,	MONEYKPI,	MONEYTRAVEL,	MONEYSERVE2,	MONEYPROJECT,	MONEYFILLPOST,TS,COPERATOR,ID,MONEYKK,BUSSTYPE)"
						+ "values (?,	?,	?,	?,	?,	?,	?,	?,		?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?,?)";
				ps = conn.prepareStatement(sql);
				DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("HSMONTH"));
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("MONTH"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("CPX"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG1"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG2"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG3"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ORG4"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("MONEYBASE"));
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("MONEYKPI"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("MONEYTRAVEL"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MONEYSERVE2"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("MONEYPROJECT"));
				DbUtil
						.setObject(ps, 13, Types.VARCHAR, vo
								.opt("MONEYFILLPOST"));
				DbUtil.setObject(ps, 14, Types.VARCHAR, request
						.getAttribute("sys.userName"));
				String PK_ID = SysUtil.getId();
				ps.setString(15, PK_ID);
				DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("MONEYKK"));
				DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("BUSSTYPE"));

				ps.execute();
			}

			if ("17".equals(imptype)) {
				sql = "insert into wb_erp.sbt_Salarydetails (MONTH,	CPX,	ORG1,	ORG2,	ORG3,	ORG4,	PSNCODE,	PSNNAME,	MONEYBASE,	MONEYKPI,	MONEYTRAVEL,	MONEYSERVE1,	MONEYSERVE2,	MONEYPROJECT,	MONEYFILLPOST,TS,COPERATOR,ID,MONEYKK,BUSSTYPE)"
						+ "values (?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?,?)";
				ps = conn.prepareStatement(sql);
				DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PSNCODE"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PSNNAME"));
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("MONEYBASE"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("MONEYKPI"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MONEYTRAVEL"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("MONEYSERVE1"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("MONEYSERVE2"));
				DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("MONEYPROJECT"));
				DbUtil
						.setObject(ps, 15, Types.VARCHAR, vo
								.opt("MONEYFILLPOST"));
				DbUtil.setObject(ps, 16, Types.VARCHAR, request
						.getAttribute("sys.userName"));
				String PK_ID = SysUtil.getId();
				ps.setString(17, PK_ID);
				DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("MONEYKK"));
				DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("BUSSTYPE"));

				ps.execute();
			}
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));

			// 提交事务
			System.out.println(f);
			// 关闭资源
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		DbUtil.closeConnection(conn);

	}

	/**
	 * 16社保福利
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_DetailsCharges(List<JSONObject> voList,
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
		// PreparedStatement ps1 = null;
		int result2 = 0;
		ResultSet rSet = null;
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
		if ("16".equals(imptype)) {
			sql = "DELETE wb_erp.SBT_DetailsCharges WHERE month in "+Times;
		}
		if ("161".equals(imptype)) {
			sql = "DELETE wb_erp.SBT_Salarydetails_other WHERE month in "+Times;;
		}
		ps1 = conn.prepareStatement(sql);
		ps1.executeUpdate();
		conn.commit();
		DbUtil.closeStatement(ps1);
	

		for ( int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			 ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("第" + f + "行" + vo.get("ORG4").toString()
						+ "部门不存在，请检查后重新导入！注：从0行算第一行");

			}

			// System.out.println(f);
			if ("16".equals(imptype)) {
				sql = "insert into wb_erp.SBT_DetailsCharges (MONTH,  CPX,  ORG1,  ORG2,  ORG3,	ORG4,	PSNCODE,	PSNNAME,	KM1,	KM2,	KM3,	KM4,	MONEY,	TYPE,	TS,	COPERATOR,ID,BUSSTYPE) "
						+ "values (?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?)";
			}
			if ("161".equals(imptype)) {
				sql = "insert into wb_erp.SBT_Salarydetails_other (MONTH,  CPX,  ORG1,  ORG2,  ORG3,	ORG4,	PSNCODE,	PSNNAME,	KM1,	KM2,	KM3,	KM4,	MONEY,	TYPE,	TS,	COPERATOR,ID,BUSSTYPE) "
						+ "values (?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?)";
			}
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PSNCODE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PSNNAME"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("MONEY"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(16, PK_ID);
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("BUSSTYPE"));
			System.out.println(f);
			ps.execute();
			// 提交事务
//System.out.println(f);
			// 关闭资源
		  DbUtil.closeStatement(ps1);
			conn.commit();
			DbUtil.closeStatement(ps);

		}

		DbUtil.closeConnection(conn);

	}

	/**
	 * 15人员组织关系
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_PersonAndOrg(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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


		for ( int f = 0; f < voList.size(); f++) {
			//
			JSONObject vo = voList.get(f);
			sql = "DELETE wb_erp.SBT_PERSONANDORG WHERE month = ? ";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
			ps1.executeUpdate();
			conn.commit();
			DbUtil.closeStatement(ps1);
		}
		

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);

			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? ";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("第" + f + "行" + vo.get("ORG4").toString()
						+ "部门不存在，请检查后重新导入！注：从0行算第一行");

			}

		
			sql = "insert into wb_erp.SBT_PERSONANDORG (MONTH, PSNCODE, PSNNAME,  DEPTNAME, POSTNAME, OUTTIME, CPX, ORG1, ORG2, ORG3, ORG4, STATUS,TS,COPERATOR,id,BUSSTYPE)"
					+ "values (?,  ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?,?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.opt("MONTH"));
			// DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("SYS"));
			// DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("ORGNAME"));
			// DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORGSHORTNAME"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("PSNCODE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("PSNNAME"));
			// DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ONJOBTIME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("DEPTNAME"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("POSTNAME"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("OUTTIME"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("STATUS"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			// DbUtil.setObject(ps, 7, Types.VARCHAR,
			// request.getAttribute("sys.userName"));
			String PK_ID = SysUtil.getId();
			ps.setString(14, PK_ID);
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("BUSSTYPE"));

			ps.execute();
			// 提交事务

			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);

		}
		conn.commit();
		// 关闭资源
		DbUtil.closeConnection(conn);

	}

	/**
	 * 14客户挂靠关系导入
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_SuperiorMerchants(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
	// TODO Auto-generated method stub
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result = 0;
		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		PreparedStatement ps1 = null;
		ResultSet rSet = null;

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			int result2 = 0;
			// 检查是否存在客户信息
			sql = "select count(1) as CT from wb_erp.bd_cubasdoc b where  b.custcode in (?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("CUSTCODE"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("FACTCUSTCODE"));

			rSet = ps.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result2 < 2) {
				throw new Exception("第" + f
						+ "行客户编码或实际所属经销商编码不存在或实际所属经销商不能与客户编码相同，请重填");
			}

			// 校验是否存在此客户，如果存在则更新，如果不存在则插入

			sql = "select count(1) as CT from wb_erp.SuperiorMerchants where custcode=?";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("CUSTCODE"));
			rSet = ps.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result2 > 0) {

				sql = "update wb_erp.SuperiorMerchants "
						+ "  set CUSTNAME = ?, FACTCUSTCODE = ?, FACTCUSTNAME = ?, TYPE = ?,operator=?,ts=to_char(sysdate,'yyyy-mm-dd hh24:mi:ss')"
						+ " where custcode = ?";

				ps = conn.prepareStatement(sql);
				DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("CUSTNAME"));
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("FACTCUSTCODE"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.get("FACTCUSTNAME"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.get("TYPE"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, request
						.getAttribute("sys.userName"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.get("CUSTCODE"));
				result = ps.executeUpdate();
				// 提交事务
				conn.commit();

			}

			// 插入
			else {
				sql = "insert into wb_erp.SuperiorMerchants (ID, CUSTCODE, CUSTNAME, FACTCUSTCODE, FACTCUSTNAME, MEMO, TYPE, OPERATOR, TS)"
						+ " values (?, ?, ?, ?, ?, '批量导入', ?, ?, to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'))";

				ps = conn.prepareStatement(sql);
				PK_ID = SysUtil.getId();
				ps.setString(1, PK_ID);
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("CUSTCODE"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("CUSTNAME"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("FACTCUSTCODE"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("FACTCUSTNAME"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("TYPE"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, request
						.getAttribute("sys.userName"));

				result = ps.executeUpdate();
				// 提交事务
				conn.commit();
			}
		}

		// 关闭资源
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);
	}

	/**
	 * 13直销客户主管客户移交
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_Customertransfer(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result;

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		PreparedStatement ps1 = null;
		ResultSet rSet = null;

		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			int result2 = 0;
			// 检查是否存在客户信息
			sql = "select count(1) as CT from wb_erp.bd_cubasdoc b where  b.custcode=?";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("CUSTCODE"));
			rSet = ps.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result2 == 0) {
				throw new Exception("第" + f + "行客户编码不存在，请重填");
			}

			// 检查是否存业务员信息
			sql = "select count(1) as CT from wb_erp.view_ddm_userinfo b where  b.USER_NAME IN (?,?)";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("PSNCODEBEFORE"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("PSNCODEAFTER"));
			rSet = ps.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result2 < 1) {
				throw new Exception("第" + f + 1 + "行业务员编码不存在，请重填");
			}

			// 校验是否存在此客户，如果存在则更新，如果不存在则插入

			sql = "select count(1) as CT from wb_erp.Customertransfer where custcode=?";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("CUSTCODE"));
			rSet = ps.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result2 > 0) {

				sql = "update wb_erp.Customertransfer a "
						+ " set DQ       = ?,"
						+ "   YXB      = ?,"
						+ "   CUSTNAME = ?,"
						+ "   TRANSFERTYPE   = ?,"
						+ "   MONTH          = ?,"
						+ "   SALESTRUBEFORE = ?,"
						+ "   PSNCODEBEFORE  = ?,"
						+ "   PSNNAMEBEFORE  = ?,"
						+ "   SALESTRUAFTER  = ?,"
						+ "   PSNCODEAFTER   = ?,"
						+ "   PSNNAMEAFTER   = ?,"
						+ "   COPERATOR      =?,"
						+ "   TS             = to_char(SYSDATE, 'YYYY-MM-DD HH24:mi:ss')"
						+ " where custcode=?";

				ps = conn.prepareStatement(sql);
				DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("DQ"));
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("YXB"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.get("CUSTNAME"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.get("TRANSFERTYPE"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.get("MONTH"));
				DbUtil
						.setObject(ps, 6, Types.VARCHAR, vo
								.get("SALESTRUBEFORE"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.get("PSNCODEBEFORE"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.get("PSNNAMEBEFORE"));
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo.get("SALESTRUAFTER"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.get("PSNCODEAFTER"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.get("PSNNAMEAFTER"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, request
						.getAttribute("sys.userName"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.get("CUSTCODE"));
				result = ps.executeUpdate();
				// 提交事务
				conn.commit();

			}

			// 插入
			else {
				sql = "insert into wb_erp.Customertransfer "
						+ "(ID,DQ,YXB,CUSTCODE, CUSTNAME,  TRANSFERTYPE, MONTH, SALESTRUBEFORE, PSNCODEBEFORE, PSNNAMEBEFORE, SALESTRUAFTER, PSNCODEAFTER, PSNNAMEAFTER, COPERATOR, TS) "
						+ "values (?,?,?, ?,?, ?, ?, ?, ?,?, ?, ?, ?,?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'))";

				ps = conn.prepareStatement(sql);
				PK_ID = SysUtil.getId();
				ps.setString(1, PK_ID);
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("DQ"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("YXB"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("CUSTCODE"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CUSTNAME"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("TRANSFERTYPE"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("MONTH"));
				DbUtil
						.setObject(ps, 8, Types.VARCHAR, vo
								.opt("SALESTRUBEFORE"));
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("PSNCODEBEFORE"));
				DbUtil
						.setObject(ps, 10, Types.VARCHAR, vo
								.opt("PSNNAMEBEFORE"));
				DbUtil
						.setObject(ps, 11, Types.VARCHAR, vo
								.opt("SALESTRUAFTER"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("PSNCODEAFTER"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("PSNNAMEAFTER"));
				DbUtil.setObject(ps, 14, Types.VARCHAR, request
						.getAttribute("sys.userName"));

				result = ps.executeUpdate();
				// 提交事务
				conn.commit();
			}
		}

		// 关闭资源
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * 导入市场范围维护 yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_scfwwh(JSONObject vo, HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		String PK_ID = null;
		String sql = "";
		int result = 0;
		if (null != vo) {
			Connection conn = DbUtil.getConnection();
			DbUtil.startTrans(conn, "");
			PreparedStatement ps = null;
			// PreparedStatement ps1 = null;

			// 先删除
			// sql= "DELETE wb_erp.APP_WB_MARKET_ZX WHERE USERCODE = ?";
			// ps1 = conn.prepareStatement(sql);
			// DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("USERCODE"));
			// result=ps1.executeUpdate();

			// 插入
			sql = "INSERT INTO wb_erp.APP_WB_MARKET_ZX "
					+ "(PK_ID, USERCODE, USERNAME, PROVINCE, CITY, AREA, PRODUCTLINE, SXRQ, SFSX) "
					+ "VALUES " + " (?, ?, ?, ?, ?, ?, ?, ?, ?)";
			ps = conn.prepareStatement(sql);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("USERCODE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("USERNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("PROVINCE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CITY"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("AREA"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PRODUCTLINE"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("SXRQ"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("SFSX"));
			result = ps.executeUpdate();

			// 提交事务
			conn.commit();

			// 关闭资源
			// DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);
			DbUtil.closeConnection(conn);
		}

	}

	/**
	 * 导入业务员与客户对应关系维护
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_CUSTMOERRELATSALESMAN(JSONObject vo,
			HttpServletRequest request, HttpServletResponse response,
			StringBuffer sBuffer, int row) throws Exception {
		String PK_ID = null;
		String sql = "";
		int result = 0;
		if (null != vo) {
			Connection conn = DbUtil.getConnection();
			DbUtil.startTrans(conn, "");
			PreparedStatement ps = null;
			PreparedStatement ps1 = null;
			ResultSet rSet = null;
			// 查询是否同一部门
			
		//修改放开权限进行导入数据
		/*	sql = "select 1 as CT "
					+ "from dual where (select distinct zsj.zzname "
					+ "       from wb_erp.zsj_ddm_userinfo zsj "
					+ "      where zsj.user_name in (?,?)) in "
					+ "    (select zu.deptname "
					+ "       from wb_erp.APP_ZBXX_USER zu "
					+ "      where zu.MANGERCODE = ?) or '集团' in (select zu.deptname from wb_erp.APP_ZBXX_USER zu where zu.MANGERCODE = ?)";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("PSNCODE"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("PSNCODE_CUSTMANAGER"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, request
					.getAttribute("sys.userName"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result != 1) {
				sBuffer.append("第" + row + "行" + vo.get("PSNNAME") + "("
						+ vo.get("PSNCODE") + "),");
			}
			*/

			// 校验是否存在部、片组
			result = 0;
			sql = "select count(1) as CT from ( "
					+ "select 1 from wb_erp.org_dept b where name like ? and rownum=1 "
					+ " union all  "
					+ "select 1 from wb_erp.org_dept b where name like ? and rownum=1 "
					+ "union all  "
					+ "select 1 from wb_erp.org_dept b where name like ? and rownum=1)";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("YXB"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("XQ"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("SALESTR"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result != 3) {
				System.out.println(sql);
				throw new Exception("第" + row + "行营销部/小区/片组与系统匹配！请重新导入");
			}

			// 校验产品线字段
			result = 0;
			sql = "select 1 AS CT " + "  from nc.bd_prodline@sbtnc a "
					+ "  where isseal = 'N'  " + "and dr = 0 "
					+ "and prodlinename = ?  ";

			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("RELATIONTYPE"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result != 1) {
				throw new Exception("第" + row + "行产品线错误，请重填");
			}
			/*
			// 校验客户激活时间
			if(!vo.get("MONTH1").equals("")){
			result = 0;
			sql = "SELECT 1 AS CT FROM DUAL  WHERE TO_CHAR(SYSDATE,'YYYY-MM')=?";

			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH1"));
			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result != 1) {
				throw new Exception("第" + row + "行老客户激活月份必须为当月，请重填");
			}
			}
			*/
			
			// 校验客户经理 与开发组长+开发人员必填其中一个以上
			if(vo.get("PSNCODE_CUSTMANAGER").toString().equals(""))
			{
				if(vo.get("MANAGERCODE").toString().equals("")&&vo.get("PSNCODE").toString().equals("") &&vo.get("MEMO1").toString().equals(""))
				{
					throw new Exception("第" + row + "行（客户经理）或（开发组长+受益业务员）至少一个为必填");
				}
				else
				{
					// 校验是否存在部、片组
					Double d = 0.0;
					Double d1 = 0.0;
					Double d2 = 0.0;
					Double d3 = 0.0;
					if(vo.get("MEMO1")==null)
					{
						throw new Exception("第" + row + "行受益比例不允许为空");
					}
					d1 = Double.parseDouble(vo.get("MEMO1").toString());
					if (vo.get("MEMO2") == null
							|| vo.get("MEMO2").toString().equals(""))
						d2 = 0.0;
					else
						d2 = Double.parseDouble(vo.get("MEMO2").toString());

					if (vo.get("MEMO3") == null
							|| vo.get("MEMO3").toString().equals(""))
						d3 = 0.0;
					else
						d3 = Double.parseDouble(vo.get("MEMO3").toString());

					d = d1 + d2 + d3;
					if (d != 1) {
						throw new Exception("第" + row + "行奖励分配比例相加必须为1");
					}
				}
			}
				
			
			// 查询是否存在
			result = 0;
			sql = "select count(*) CT from wb_erp.zsj_so_custmoerrelatsalesman a where a.CUSTCODE= ? and a.RELATIONTYPE= (select a.pk_prodline from nc.bd_prodline@sbtnc a where  isseal='N' and dr=0 and prodlinename= ?) and a.TYPE= ? ";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("CUSTCODE"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("RELATIONTYPE"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("TYPE"));
			rSet = ps1.executeQuery();
			//System.out.println(vo.get("CUSTCODE"));
			if (rSet.next()) {
				result = rSet.getInt("CT");
			}
			if (result > 0) {
				// 修改
				sql = "update wb_erp.zsj_so_custmoerrelatsalesman "
						+ " set PSNCODE         = ?, "
						+ "  PSNNAME         = ?, "
						+ "   JOBNAME     = (select du.postname "
						+ "   from wb_erp.zsj_ddm_userinfo du "
						+ "    where du.user_name = ?), "
						+ "  VSALESTRUNAME_3 = (select du.zzname "
						+ "      from wb_erp.zsj_ddm_userinfo du "
						+ "     where du.user_name = ?), "
						+ "       VSALESTRUNAME   = (select du.bmname "
						+ "      from wb_erp.zsj_ddm_userinfo du "
						+ "                           where du.user_name = ?), "
						+ "       COPERATORID     = ?, "
						+ "       FSTATUS         = 1, "
						+ "       TS              = to_char(SYSDATE, 'YYYY-MM-DD HH24:mi:ss'), "
						+ "       PSNCODEMANA_TECHNICAL   = ?, "
						+ "       MEMO1           = ?, "
						+ "       PSNCODE2        = ?, "
						+ "       PSNNAME2        = ?, "
						+ "       MEMO2           = ?, "
						+ "       PSNCODE3        = ?, "
						+ "       PSNNAME3        = ?, "
						+ "       MEMO3           = ?, "
						+ "       MANAGERCODE     = ?, "
						+ "       MANAGERNAME     = ?, "
						+ "       MEMO            = ?, "
						+ "       YXB             = ? ,"
						+ "       salestr         = ? ," 
						+ "       XQ = ? ,"
						+ "       PSNNAMEMANA_TECHNICAL = ? ," 
						+ "       MONTH2 = ? ,"
						+ "       MPIG = nvl(?,MPIG) ," 
						+ "      RPIG = nvl(?,RPIG),  "
						
						+ "       PSNCODE_CUSTMANAGER = ? ,"
						+ "       PSNNAME_CUSTMANAGER = ? ," 
						+ "       PSNCODE1_TECHNICAL = ? ,"
						+ "       PSNNAME1_TECHNICAL = ? ," 
						+ "       TECHNICAL1_RATE = ? ,"
						+ "       PSNCODE2_TECHNICAL = ? ,"
						+ "       PSNNAME2_TECHNICAL = ? ," 
						+ "       TECHNICAL2_RATE = ? "
						
						+ " where CUSTCODE = ? "
						+ " and RELATIONTYPE = (select a.pk_prodline "
						+ "  from nc.bd_prodline@sbtnc a "
						+ "  where isseal = 'N' " + " and dr = 0 "
						+ "  and prodlinename = ?) " + "   and TYPE = 'Y' ";
				ps = conn.prepareStatement(sql);
				DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("PSNCODE"));
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("PSNNAME"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.get("PSNCODE"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.get("PSNCODE"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.get("PSNCODE"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, request
						.getAttribute("sys.userName"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.get("PSNCODEMANA_TECHNICAL"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.get("MEMO1"));
				DbUtil.setObject(ps, 9, Types.VARCHAR, vo.get("PSNCODE2"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.get("PSNNAME2"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.get("MEMO2"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.get("PSNCODE3"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.get("PSNNAME3"));
				DbUtil.setObject(ps, 14, Types.VARCHAR, vo.get("MEMO3"));
				DbUtil.setObject(ps, 15, Types.VARCHAR, vo.get("MANAGERCODE"));
				DbUtil.setObject(ps, 16, Types.VARCHAR, vo.get("MANAGERNAME"));
				DbUtil.setObject(ps, 17, Types.VARCHAR, vo.get("MEMO"));
				DbUtil.setObject(ps, 18, Types.VARCHAR, vo.get("YXB"));
				DbUtil.setObject(ps, 19, Types.VARCHAR, vo.get("SALESTR"));
				DbUtil.setObject(ps, 20, Types.VARCHAR, vo.get("XQ"));
				DbUtil.setObject(ps, 21, Types.VARCHAR, vo.get("PSNNAMEMANA_TECHNICAL"));
				DbUtil.setObject(ps, 22, Types.VARCHAR, vo.get("MONTH2"));
				DbUtil.setObject(ps, 23, Types.VARCHAR, vo.get("MPIG"));
				DbUtil.setObject(ps, 24, Types.VARCHAR, vo.get("RPIG"));
				
				DbUtil.setObject(ps, 25, Types.VARCHAR, vo.get("PSNCODE_CUSTMANAGER"));
				DbUtil.setObject(ps, 26, Types.VARCHAR, vo.get("PSNNAME_CUSTMANAGER"));
				DbUtil.setObject(ps, 27, Types.VARCHAR, vo.get("PSNCODE1_TECHNICAL"));
				DbUtil.setObject(ps, 28, Types.VARCHAR, vo.get("PSNNAME1_TECHNICAL"));
				DbUtil.setObject(ps, 29, Types.VARCHAR, vo.get("TECHNICAL1_RATE"));
				DbUtil.setObject(ps, 30, Types.VARCHAR, vo.get("PSNCODE2_TECHNICAL"));
				DbUtil.setObject(ps, 31, Types.VARCHAR, vo.get("PSNNAME2_TECHNICAL"));
				DbUtil.setObject(ps, 32, Types.VARCHAR, vo.get("TECHNICAL2_RATE"));
				
				DbUtil.setObject(ps, 33, Types.VARCHAR, vo.get("CUSTCODE"));
				DbUtil.setObject(ps, 34, Types.VARCHAR, vo.get("RELATIONTYPE"));
				result = ps.executeUpdate();

				// 修改多个产品线的客户，实际经销商只允许一个
				/*
				 * sql="update wb_erp.zsj_so_custmoerrelatsalesman a set a.managercustcode=?,a.managercustname=?,a.managertype=? where custcode=?"
				 * ; ps = conn.prepareStatement(sql); DbUtil.setObject(ps, 1,
				 * Types.VARCHAR, vo.get("MANAGERCUSTCODE"));
				 * DbUtil.setObject(ps, 2, Types.VARCHAR,
				 * vo.get("MANAGERCUSTNAME")); DbUtil.setObject(ps, 3,
				 * Types.VARCHAR, vo.get("MANAGERTYPE")); DbUtil.setObject(ps,
				 * 4, Types.VARCHAR, vo.get("CUSTCODE")); result =
				 * ps.executeUpdate();
				 */
			} else {
				// 插入
				sql = "insert into wb_erp.zsj_so_custmoerrelatsalesman "
						+ "(PK_DZ,BILLCODE,DATEVERSION,CUSTCODE,CUSTNAME, PSNCODE,PSNNAME,JOBNAME,VSALESTRUNAME_3, "
						+ "VSALESTRUNAME,ISRECENTLY,COPERATORID,DBILLDATE,CAPPROVEID,DAPPROVEDATE,FSTATUS, "
						+ "DR,TS,TYPE,RELATIONTYPE,SALESTR,MEMO,YXB,MANAGERCODE,MANAGERNAME,XQ,PSNCODEMANA_TECHNICAL,MEMO2,MEMO3,MEMO1,PSNCODE2,PSNNAME2,PSNCODE3,PSNNAME3,PSNNAMEMANA_TECHNICAL,MONTH2,MPIG,RPIG,  PSNCODE_CUSTMANAGER ,PSNNAME_CUSTMANAGER,PSNCODE1_TECHNICAL, PSNNAME1_TECHNICAL, TECHNICAL1_RATE, PSNCODE2_TECHNICAL, PSNNAME2_TECHNICAL, TECHNICAL2_RATE) "
						+ "values "
						+ "(?,'',to_char(sysdate, 'YYYY-MM-DD'),?,?,?,?, "
						+ "(select du.postname from wb_erp.zsj_ddm_userinfo du where du.user_name = ?), "
						+ "(select du.zzname from wb_erp.zsj_ddm_userinfo du where du.user_name = ?), "
						+ "(select du.bmname from wb_erp.zsj_ddm_userinfo du where du.user_name = ?), "
						+ "'Y',?,to_char(SYSDATE, 'YYYY-MM-DD HH24:mi:ss'),'','',0,0,to_char(SYSDATE, 'YYYY-MM-DD HH24:mi:ss'), "
						+ "?,(select a.pk_prodline from nc.bd_prodline@sbtnc a where  isseal='N' and dr=0 and prodlinename= ?),"
						+ "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
				ps = conn.prepareStatement(sql);
				PK_ID = SysUtil.getId();
				ps.setString(1, PK_ID);
				DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("CUSTCODE"));
				DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("CUSTNAME"));
				DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("PSNCODE"));
				DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("PSNNAME"));
				DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("PSNCODE"));
				DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PSNCODE"));
				DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PSNCODE"));
				DbUtil.setObject(ps, 9, Types.VARCHAR, request
						.getAttribute("sys.userName"));
				DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("TYPE"));
				DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("RELATIONTYPE"));
				DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("SALESTR"));
				DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("MEMO"));
				DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("YXB"));
				DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("MANAGERCODE"));
				DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("MANAGERNAME"));
				DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("XQ"));
				DbUtil
						.setObject(ps, 18, Types.VARCHAR, vo
								.opt("PSNCODEMANA_TECHNICAL"));
				DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("MEMO2"));
				DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("MEMO3"));
				DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("MEMO1"));
				DbUtil.setObject(ps, 22, Types.VARCHAR, vo.opt("PSNCODE2"));
				DbUtil.setObject(ps, 23, Types.VARCHAR, vo.opt("PSNNAME2"));
				DbUtil.setObject(ps, 24, Types.VARCHAR, vo.opt("PSNCODE3"));
				DbUtil.setObject(ps, 25, Types.VARCHAR, vo.opt("PSNNAME3"));
				DbUtil.setObject(ps, 26, Types.VARCHAR, vo.opt("PSNNAMEMANA_TECHNICAL"));
				DbUtil.setObject(ps, 27, Types.VARCHAR, vo.opt("MONTH2"));
				DbUtil.setObject(ps, 28, Types.VARCHAR, vo.opt("MPIG"));
				DbUtil.setObject(ps, 29, Types.VARCHAR, vo.opt("RPIG"));
				
				DbUtil.setObject(ps, 30, Types.VARCHAR, vo.get("PSNCODE_CUSTMANAGER"));
				DbUtil.setObject(ps, 31, Types.VARCHAR, vo.get("PSNNAME_CUSTMANAGER"));
				DbUtil.setObject(ps, 32, Types.VARCHAR, vo.get("PSNCODE1_TECHNICAL"));
				DbUtil.setObject(ps, 33, Types.VARCHAR, vo.get("PSNNAME1_TECHNICAL"));
				DbUtil.setObject(ps, 34, Types.VARCHAR, vo.get("TECHNICAL1_RATE"));
				DbUtil.setObject(ps, 35, Types.VARCHAR, vo.get("PSNCODE2_TECHNICAL"));
				DbUtil.setObject(ps, 36, Types.VARCHAR, vo.get("PSNNAME2_TECHNICAL"));
				DbUtil.setObject(ps, 37, Types.VARCHAR, vo.get("TECHNICAL2_RATE"));

				result = ps.executeUpdate();
			}
			// 提交事务
			conn.commit();

			// 关闭资源
			DbUtil.closeResultSet(rSet);
			DbUtil.closeStatement(ps1);
			DbUtil.closeStatement(ps);
			DbUtil.closeConnection(conn);
		}

	}

	/**
	 * 导入可贷额
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_app_jr_kde(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// 先删除
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);
		// result=ps1.executeUpdate();

		// 插入
		sql = "insert into wb_erp.app_jr_kde "
				+ "(PK_ID, KHBM, YXB, KHXM, KHDH, ZCSJ, JBD2, JBD3, FXDKYE, ZXDKYE, XYKYE, FXSXED, DQKDED,TS) "
				+ "values " + "(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,sysdate)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("KHBM"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("KHXM"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("KHDH"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ZCSJ"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("JBD2"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("JBD3"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("FXDKYE"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("ZXDKYE"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("XYKYE"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("FXSXED"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("DQKDED"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}

		result = ps.executeBatch();

		// 提交事务
		conn.commit();

		// 关闭资源
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * 导入分销查询 yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_FX(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// 先删除
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);
		// result=ps1.executeUpdate();

		// 插入
		sql = "INSERT INTO wb_erp.APP_JR_FX "
				+ "(PK_ID, KHBS, MQYH, DQ, SQ, YXB, GCMC, KHLX, KHBM, KHXM, SZXS, JYZE, TJED, JYEQJ, DH, SFZH, HZNX, YYZZH, XCMQJG, YYSM, XCMQR, MQRQ, CLQQRQ, XDSX, WTLB, XDSHSM, TJRQ, TJHBH, HKKH, FKRQ, FKJE, FKLL, FKQS, NCSKDRQ, SFDHTZTH, WHR, WHRQ, BZ, YFKSJ, CCTS, FKHS, TQHKSJ, YHKSJ, DQYE, TS) "
				+ "VALUES "
				+ "(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, sysdate)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("KHBS"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("MQYH"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("DQ"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("SQ"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("GCMC"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("KHLX"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KHBM"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KHXM"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("SZXS"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("JYZE"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("TJED"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("JYEQJ"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("DH"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("SFZH"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("HZNX"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("YYZZH"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("XCMQJG"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("YYSM"));
			DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("XCMQR"));
			DbUtil.setObject(ps, 22, Types.VARCHAR, vo.opt("MQRQ"));
			DbUtil.setObject(ps, 23, Types.VARCHAR, vo.opt("CLQQRQ"));
			DbUtil.setObject(ps, 24, Types.VARCHAR, vo.opt("XDSX"));
			DbUtil.setObject(ps, 25, Types.VARCHAR, vo.opt("WTLB"));
			DbUtil.setObject(ps, 26, Types.VARCHAR, vo.opt("XDSHSM"));
			DbUtil.setObject(ps, 27, Types.VARCHAR, vo.opt("TJRQ"));
			DbUtil.setObject(ps, 28, Types.VARCHAR, vo.opt("TJHBH"));
			DbUtil.setObject(ps, 29, Types.VARCHAR, vo.opt("HKKH"));
			DbUtil.setObject(ps, 30, Types.VARCHAR, vo.opt("FKRQ"));
			DbUtil.setObject(ps, 31, Types.VARCHAR, vo.opt("FKJE"));
			DbUtil.setObject(ps, 32, Types.VARCHAR, vo.opt("FKLL"));
			DbUtil.setObject(ps, 33, Types.VARCHAR, vo.opt("FKQS"));
			DbUtil.setObject(ps, 34, Types.VARCHAR, vo.opt("NCSKDRQ"));
			DbUtil.setObject(ps, 35, Types.VARCHAR, vo.opt("SFDHTZTH"));
			DbUtil.setObject(ps, 36, Types.VARCHAR, vo.opt("WHR"));
			DbUtil.setObject(ps, 37, Types.VARCHAR, vo.opt("WHRQ"));
			DbUtil.setObject(ps, 38, Types.VARCHAR, vo.opt("BZ"));
			DbUtil.setObject(ps, 39, Types.VARCHAR, vo.opt("YFKSJ"));
			DbUtil.setObject(ps, 40, Types.VARCHAR, vo.opt("CCTS"));
			DbUtil.setObject(ps, 41, Types.VARCHAR, vo.opt("FKHS"));
			DbUtil.setObject(ps, 42, Types.VARCHAR, vo.opt("TQHKSJ"));
			DbUtil.setObject(ps, 43, Types.VARCHAR, vo.opt("YHKSJ"));
			DbUtil.setObject(ps, 44, Types.VARCHAR, vo.opt("DQYE"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}
		result = ps.executeBatch();

		// 提交事务
		conn.commit();

		// 关闭资源
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * 导入直销查询 yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_ZX(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// 先删除
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);
		// result=ps1.executeUpdate();

		// 插入
		sql = "INSERT INTO wb_erp.APP_JR_ZX "
				+ "(PK_ID, KHLX, DKYH, SSSC, DQ, JD, YXB, GCMC, KHBM, KHXM, KHDH, SZXS, SCTHRQ, SFZH, CLMZTS, YZCJ, YZNX, TJED, XTZCSJ, JYNJYE, JCMZTS, CLRZTS, DQYHFZ, DQMJFZ, XYGLDS, JSOARQ, DSHRQ, MQRY, HSMQRQ, XDSX, WTLB, XDSHSM, TJRQ, TJHBH, SFBQCL, CLQQRQ, SFTP, TPFL, TPSM, KJTS, XQMZTS, JDYJ, FKJE, FKRQ, NCXZRQ, SFJBZJ, FKQS, WLSBM, WLSMC, DH, WLSSFZH, YHFZE, JZ, BZ, YWY, YHBJRQ, SJJQBJRQ, HKKH, DKLL, XDTJYF, XTZCSJ2, JYNJYE2, TS) "
				+ "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, SYSDATE)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("KHLX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("DKYH"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("SSSC"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("DQ"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("JD"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("GCMC"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KHBM"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KHXM"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("KHDH"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("SZXS"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("SCTHRQ"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("SFZH"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("CLMZTS"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("YZCJ"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("YZNX"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("TJED"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("XTZCSJ"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("JYNJYE"));
			DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("JCMZTS"));
			DbUtil.setObject(ps, 22, Types.VARCHAR, vo.opt("CLRZTS"));
			DbUtil.setObject(ps, 23, Types.VARCHAR, vo.opt("DQYHFZ"));
			DbUtil.setObject(ps, 24, Types.VARCHAR, vo.opt("DQMJFZ"));
			DbUtil.setObject(ps, 25, Types.VARCHAR, vo.opt("XYGLDS"));
			DbUtil.setObject(ps, 26, Types.VARCHAR, vo.opt("JSOARQ"));
			DbUtil.setObject(ps, 27, Types.VARCHAR, vo.opt("DSHRQ"));
			DbUtil.setObject(ps, 28, Types.VARCHAR, vo.opt("MQRY"));
			DbUtil.setObject(ps, 29, Types.VARCHAR, vo.opt("HSMQRQ"));
			DbUtil.setObject(ps, 30, Types.VARCHAR, vo.opt("XDSX"));
			DbUtil.setObject(ps, 31, Types.VARCHAR, vo.opt("WTLB"));
			DbUtil.setObject(ps, 32, Types.VARCHAR, vo.opt("XDSHSM"));
			DbUtil.setObject(ps, 33, Types.VARCHAR, vo.opt("TJRQ"));
			DbUtil.setObject(ps, 34, Types.VARCHAR, vo.opt("TJHBH"));
			DbUtil.setObject(ps, 35, Types.VARCHAR, vo.opt("SFBQCL"));
			DbUtil.setObject(ps, 36, Types.VARCHAR, vo.opt("CLQQRQ"));
			DbUtil.setObject(ps, 37, Types.VARCHAR, vo.opt("SFTP"));
			DbUtil.setObject(ps, 38, Types.VARCHAR, vo.opt("TPFL"));
			DbUtil.setObject(ps, 39, Types.VARCHAR, vo.opt("TPSM"));
			DbUtil.setObject(ps, 40, Types.VARCHAR, vo.opt("KJTS"));
			DbUtil.setObject(ps, 41, Types.VARCHAR, vo.opt("XQMZTS"));
			DbUtil.setObject(ps, 42, Types.VARCHAR, vo.opt("JDYJ"));
			DbUtil.setObject(ps, 43, Types.VARCHAR, vo.opt("FKJE"));
			DbUtil.setObject(ps, 44, Types.VARCHAR, vo.opt("FKRQ"));
			DbUtil.setObject(ps, 45, Types.VARCHAR, vo.opt("NCXZRQ"));
			DbUtil.setObject(ps, 46, Types.VARCHAR, vo.opt("SFJBZJ"));
			DbUtil.setObject(ps, 47, Types.VARCHAR, vo.opt("FKQS"));
			DbUtil.setObject(ps, 48, Types.VARCHAR, vo.opt("WLSBM"));
			DbUtil.setObject(ps, 49, Types.VARCHAR, vo.opt("WLSMC"));
			DbUtil.setObject(ps, 50, Types.VARCHAR, vo.opt("DH"));
			DbUtil.setObject(ps, 51, Types.VARCHAR, vo.opt("WLSSFZH"));
			DbUtil.setObject(ps, 52, Types.VARCHAR, vo.opt("YHFZE"));
			DbUtil.setObject(ps, 53, Types.VARCHAR, vo.opt("JZ"));
			DbUtil.setObject(ps, 54, Types.VARCHAR, vo.opt("BZ"));
			DbUtil.setObject(ps, 55, Types.VARCHAR, vo.opt("YWY"));
			DbUtil.setObject(ps, 56, Types.VARCHAR, vo.opt("YHBJRQ"));
			DbUtil.setObject(ps, 57, Types.VARCHAR, vo.opt("SJJQBJRQ"));
			DbUtil.setObject(ps, 58, Types.VARCHAR, vo.opt("HKKH"));
			DbUtil.setObject(ps, 59, Types.VARCHAR, vo.opt("DKLL"));
			DbUtil.setObject(ps, 60, Types.VARCHAR, vo.opt("XDTJYF"));
			DbUtil.setObject(ps, 61, Types.VARCHAR, vo.opt("XTZCSJ2"));
			DbUtil.setObject(ps, 62, Types.VARCHAR, vo.opt("JYNJYE2"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}
		result = ps.executeBatch();

		// 提交事务
		conn.commit();

		// 关闭资源
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * 导入贴息销量目标查询 yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_TXXLMB(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// 先删除
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);
		// result=ps1.executeUpdate();

		// 插入
		sql = "INSERT INTO wb_erp.APP_JR_TXXLMB "
				+ "(PK_ID, BLYH, YXB, KHBM, KHXM, FKJE, XLJSQJ, XLJS, ZLMB, DQSJXL, QEBXXL, TS) "
				+ "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, sysdate)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("BLYH"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("KHBM"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("KHXM"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("FKJE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("XLJSQJ"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("XLJS"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("ZLMB"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("DQSJXL"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("QEBXXL"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}

		result = ps.executeBatch();

		// 提交事务
		conn.commit();

		// 关闭资源
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * 导入贴息情况查询 yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_TXQK(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// 先删除
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);
		// result=ps1.executeUpdate();

		// 插入
		sql = "INSERT INTO wb_erp.APP_JR_TXQK "
				+ "(PK_ID, BXFA, BXYF, BLYH, YXB, GCMC, KHBM, KHXM, DH, SFZH, MQRQ, FKRQ, FKJE, DKLX, HKWBSJ, MQQYNXL, MQHYNXL, BXL, BXJE, BXSM, TS) "
				+ "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, sysdate)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("BXFA"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("BXYF"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("BLYH"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("GCMC"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("KHBM"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("KHXM"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("DH"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("SFZH"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("MQRQ"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.get("FKRQ"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("FKJE"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("DKLX"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("HKWBSJ"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("MQQYNXL"));
			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("MQHYNXL"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("BXL"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("BXJE"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("BXSM"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}

		result = ps.executeBatch();

		// 提交事务
		conn.commit();

		// 关闭资源
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * 导入还本查询 yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_HB(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// 先删除
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);
		// result=ps1.executeUpdate();

		// 插入
		sql = "INSERT INTO wb_erp.APP_JR_HB "
				+ "(PK_ID, KHLX, YXB, DKYH, KHBM, KHXM, KHDH, YHSJ, YHJE, HKKH, TS) "
				+ "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, sysdate)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("KHLX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("DKYH"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("KHBM"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("KHXM"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("KHDH"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("YHSJ"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("YHJE"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("HKKH"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}

		result = ps.executeBatch();

		// 提交事务
		conn.commit();

		// 关闭资源
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * 导入预警查询 yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_YJ(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// 先删除
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);
		// result=ps1.executeUpdate();

		// 插入
		sql = "INSERT INTO wb_erp.APP_JR_YJ "
				+ "(PK_ID, YXB, KHBM, KHXM, QYXL, SYXL, DYXL, DXB, YJYY, JCYJMB, YJCS, TS) "
				+ "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, sysdate)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("YXB"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("KHBM"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("KHXM"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("QYXL"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("SYXL"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("DYXL"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("DXB"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("YJYY"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("JCYJMB"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("YJCS"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}

		result = ps.executeBatch();

		// 提交事务
		conn.commit();

		// 关闭资源
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * 12导入放款信息
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_FKXX(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		ResultSet rSet = null;

		// 修改
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// 校验是否存在异常推荐编号客户
			int result2 = 0;
			sql = "select count(*) as CT from wb_erp.APP_JR_ZX where  TJHBH=? and khbm=? and KHXM=?";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("TJHBH"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("KHBM"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.get("KHXM"));
			rSet = ps.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			DbUtil.closeResultSet(rSet);
			if (result2 == 0) {
				// DbUtil.closeStatement(ps1);
				DbUtil.closeStatement(ps);
				DbUtil.closeConnection(conn);
				throw new Exception("推荐编号：" + vo.get("TJHBH") + "与客户不匹配！请重新导入");
			}

			sql = "update wb_erp.APP_JR_ZX set  FKRQ=?,FKJE=?,HKKH=?,NCXZRQ=? where TJHBH=?";
			ps = conn.prepareStatement(sql);
			DbUtil.setObject(ps, 1, Types.VARCHAR, vo.get("FKRQ"));
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("FKJE"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.get("HKKH"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.get("NCXZRQ"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.get("TJHBH"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}
		result = ps.executeBatch();

		// 提交事务
		conn.commit();

		// 关闭资源
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * 导入金融逾期信息 boyang
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_YQXX(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// 先删除
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);1426
		// result=ps1.executeUpdate();

		// 插入
		sql = "INSERT INTO wb_erp.APP_JR_YQXX "
				+ "(PK_YQXX, KHLX, KHBM, ZQ, YXB, KHXM, KHDH, SZXS, BLYH, DKED, HKKH, YQLB, YQKSSJ, YQTS,YQJE,GJQK,QYHFQK,SFZH,JQRQ,ZHJE,YQYE) "
				+ "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?,?,?,?,?,?,?)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("KHLX"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("KHBM"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ZQ"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("KHXM"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("KHDH"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("SZXS"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("BLYH"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("DKED"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("HKKH"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("YQLB"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("YQKSSJ"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("YQTS"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, vo.opt("YQJE"));
			DbUtil.setObject(ps, 16, Types.VARCHAR, vo.opt("GJQK"));

			DbUtil.setObject(ps, 17, Types.VARCHAR, vo.opt("QYHFQK"));
			DbUtil.setObject(ps, 18, Types.VARCHAR, vo.opt("SFZH"));
			DbUtil.setObject(ps, 19, Types.VARCHAR, vo.opt("JQRQ"));
			DbUtil.setObject(ps, 20, Types.VARCHAR, vo.opt("ZHJE"));
			DbUtil.setObject(ps, 21, Types.VARCHAR, vo.opt("YQYE"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}

		result = ps.executeBatch();

		// 提交事务
		conn.commit();

		// 关闭资源
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	/**
	 * 导入对应挂点领导及金融专员 yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void imp_APP_JR_GDLDorJRZY(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String PK_ID = null;
		String sql = "";
		int result[];

		Connection conn = DbUtil.getConnection();
		DbUtil.startTrans(conn, "");
		PreparedStatement ps = null;
		// PreparedStatement ps1 = null;

		// 先删除
		// sql= "DELETE APP_JR_KDE WHERE 1 = 1";
		// ps1 = conn.prepareStatement(sql);
		// result=ps1.executeUpdate();

		// 插入
		sql = "INSERT INTO wb_erp.APP_JR_AUTH "
				+ "(PK_ID, DQ, JD, YXB, ZXSC, JRQY, JRXFSC, USERCODE, USERNAME, GDLDBM, GDLDXM, JRZYBM, JRZYXM, ISALL,TS) "
				+ "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0,sysdate)";
		ps = conn.prepareStatement(sql);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.get("DQ"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("JD"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("YXB"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ZXSC"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("JRQY"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("JRXFSC"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("USERCODE"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("USERNAME"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("GDLDBM"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("GDLDXM"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("JRZYBM"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("JRZYXM"));
			ps.addBatch();
			if ((f + 1) % 1000 == 0) {
				ps.executeBatch();
			}
		}

		result = ps.executeBatch();

		// 提交事务
		conn.commit();

		// 关闭资源
		// DbUtil.closeStatement(ps1);
		DbUtil.closeStatement(ps);
		DbUtil.closeConnection(conn);

	}

	public static void getBIFile(HttpServletRequest request,
			HttpServletResponse response) throws Exception {

		InputStream in = (InputStream) request.getAttribute("uploadFile");
		String fileName = request.getAttribute("uploadFile__name").toString();
		String fileType = fileName.substring(fileName.lastIndexOf(".") + 1,
				fileName.length());
		String imptype = request.getAttribute("imptype").toString();
		Map<String, String> map = new HashMap<String, String>();
		if ("1".equals(imptype)) { // 工厂、产品线、节点、客户、地区、产品、销量、标吨
			map.put("日期", "DEF");
			map.put("工厂", "DEF1");
			map.put("产品线", "DEF2");
			map.put("节点", "DEF3");
			map.put("客户", "DEF4");
			map.put("地区", "DEF5");
			map.put("产品", "DEF6");
			map.put("销量", "DEF7");
			map.put("标吨", "DEF8");
		} else if ("2".equals(imptype)) { // 产品线、节点、客户、地区、产品、促销编号,促销名称
			map.put("日期", "DEF");
			map.put("产品线", "DEF1");
			map.put("节点", "DEF2");
			map.put("客户", "DEF3");
			map.put("地区", "DEF4");
			map.put("产品", "DEF5");
			map.put("促销编号", "DEF6");
			map.put("促销名称", "DEF7");
		} else if ("3".equals(imptype)) { // 节点、标准金额
			map.put("日期", "DEF");
			map.put("节点", "DEF1");
			map.put("标准金额", "DEF2");
		} else if ("4".equals(imptype)) { // 产品线、节点、营销员、计提金额、调整金额
			map.put("日期", "DEF");
			map.put("产品线", "DEF1");
			map.put("节点", "DEF2");
			map.put("营销员", "DEF3");
			map.put("计提金额", "DEF4");
			map.put("调整金额", "DEF5");
		} else if ("5".equals(imptype)) { // 产品线、节点、营销员、客户、计提金额、调整金额
			map.put("日期", "DEF");
			map.put("产品线", "DEF1");
			map.put("节点", "DEF2");
			map.put("营销员", "DEF3");
			map.put("客户", "DEF4");
			map.put("计提金额", "DEF5");
			map.put("调整金额", "DEF6");
		} else if ("6".equals(imptype)) { // 产品线、节点
			map.put("日期", "DEF");
			map.put("产品线", "DEF1");
			map.put("节点", "DEF2");
		}

		readBI(in, fileType, map, request, response);

	}
	
	/**
	 * 29费用导入
	 * 
	 * @param vo
	 * @throws Exception
	 */

	private static void imp_sbt_costimport(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		sql = "DELETE wb_erp.SBT_COSTIMPORT WHERE MONTH in "+Times;
		PreparedStatement ps2 = conn.prepareStatement(sql);
		ps2.executeUpdate();
		DbUtil.closeStatement(ps2);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? AND MONTH=?";
			PreparedStatement ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));
			DbUtil.setObject(ps1, 5, Types.VARCHAR, vo.get("MONTH"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("第" + f + "行" + vo.get("ORG4").toString()
						+ "部门不存在，请检查后重新导入！注：从0行算第一行");
			}

			if (f == 0) {
				// 先删除,福利社保按照月份及科目四进行覆盖删除
				sql = "DELETE wb_erp.SBT_DetailsCharges WHERE month = ? and km4=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("KM4"));
				ps1.executeUpdate();
			}
			DbUtil.closeStatement(ps1);

			sql = "insert into wb_erp.SBT_COSTIMPORT (id,month,psnname,psncode,cpxx,ORG1,  ORG2,  ORG3,	ORG4,price,	KM1,	KM2,	KM3,	KM4,TS,	COPERATOR) "
					+ "values (?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("PSNNAME"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("PSNCODE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("CPXX"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 15, Types.VARCHAR, request
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
	
	private static void imp_sbt_guding(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		String Times = "(";
		DecimalFormat decimalFormat = new DecimalFormat("###################.###########");
		for(int f = 0; f < voList.size(); f++) {
			JSONObject obj = voList.get(f);
			if(Times.indexOf(decimalFormat.format(obj.get("MONTH")))==-1){
				if(f==0) 
				{
					Times += "'"+decimalFormat.format(obj.get("MONTH"))+"'";
				}
				else 
				{
					Times += ",'"+decimalFormat.format(obj.get("MONTH"))+"'";
				}
			}
		}
		Times += ")";
		//删除包含的历史数据
		sql = "DELETE wb_erp.sbt_detailsguding WHERE MONTH in "+Times;
		PreparedStatement ps2 = conn.prepareStatement(sql);
		ps2.executeUpdate();
		DbUtil.closeStatement(ps2);
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			// int result2 = 0;
			sql = "select count(1)  CT from  wb_erp.sbt_deptcheck where ORG1=? and ORG2=? and ORG3=? and ORG4=? AND MONTH=?";
			PreparedStatement ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ORG1"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("ORG2"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("ORG3"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("ORG4"));
			DbUtil.setObject(ps1, 5, Types.VARCHAR, vo.get("MONTH"));

			rSet = ps1.executeQuery();
			if (rSet.next()) {
				result2 = rSet.getInt("CT");
			}
			if (result2 == 0) {
				throw new Exception("第" + f + "行" + vo.get("ORG4").toString()
						+ "部门不存在，请检查后重新导入！注：从0行算第一行");
			}

			if (f == 0) {
				// 先删除,福利社保按照月份及科目四进行覆盖删除
				sql = "DELETE wb_erp.SBT_DetailsCharges WHERE month = ? and km4=?";
				ps1 = conn.prepareStatement(sql);
				DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("MONTH"));
				DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("KM4"));
				ps1.executeUpdate();
			}
			DbUtil.closeStatement(ps1);

			sql = "insert into wb_erp.sbt_detailsguding (id,month,cpx,ORG1,  ORG2,  ORG3,	ORG4,money,	KM1,	KM2,	KM3,	KM4,PZH,TS,	COPERATOR) "
					+ "values (?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,	?,to_char(sysdate,'yyyy-mm-dd hh24:mi:ss'),?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("CPX"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("ORG1"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("ORG2"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("ORG3"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("ORG4"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("PRICE"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("KM1"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("KM2"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("KM3"));
			DbUtil.setObject(ps, 12, Types.VARCHAR, vo.opt("KM4"));
			DbUtil.setObject(ps, 13, Types.VARCHAR, vo.opt("PZH"));
			DbUtil.setObject(ps, 14, Types.VARCHAR, request
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
	
	private static void imp_saletarget(List<JSONObject> voList,
			HttpServletRequest request, HttpServletResponse response)
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
		for (int f = 0; f < voList.size(); f++) {
			JSONObject vo = voList.get(f);
			//删除历史数据
			sql = "DELETE wb_erp.SALETARGET WHERE ZZNAME=? and TYPE =? and MONTH=? and MEASURE=? and PRODLINENAME=?";
			PreparedStatement ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.VARCHAR, vo.get("ZZNAME"));
			DbUtil.setObject(ps1, 2, Types.VARCHAR, vo.get("TYPE"));
			DbUtil.setObject(ps1, 3, Types.VARCHAR, vo.get("MONTH"));
			DbUtil.setObject(ps1, 4, Types.VARCHAR, vo.get("MEASURE"));
			DbUtil.setObject(ps1, 5, Types.VARCHAR, vo.get("PRODLINENAME"));
			ps1.executeUpdate();
			DbUtil.closeStatement(ps1);

			sql = "insert into wb_erp.SALETARGET (ID,ZZNAME,MONTH,TYPE,TARGET,MEASURE,PRODLINENAME)"
					+ "values (? ,	?,	?,	?,	?,	?,?)";
			ps = conn.prepareStatement(sql);
			String PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.VARCHAR, vo.opt("ZZNAME"));
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("MONTH"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("TYPE"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("TARGET"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("MEASURE"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("PRODLINENAME"));
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
	 * 读取文件
	 * 
	 * @param in
	 * @param fileType
	 * @throws Exception
	 */
	public static void readBI(InputStream in, String fileType,
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
								if (headRow.getCell(j) != null
										&& headRow.getCell(j)
												.getStringCellValue().contains(
														"日期")
										&& HSSFDateUtil
												.isCellDateFormatted(cell)) {
									cellVal = cell.getDateCellValue();
								} else {
									cellVal = ExcelObject.getCellValue(cell);
								}
								jsonObject.put(map.get(headRow.getCell(j)
										.getStringCellValue().toString()),
										cellVal);
							}
							list.add(jsonObject);
						}
					}
				}

			} else {
				throw new Exception("只支持2003版本的Excel导入！");
			}

			for (int f = 0; f < list.size(); f++) {
				impBI(list.get(f), request, response);
			}
			in.close();
		} catch (Exception e) {
			dqrow = dqrow + 1;
			throw e;

		}
	}

	/**
	 * 导入市场范围维护 yezq
	 * 
	 * @param vo
	 * @throws Exception
	 */
	public static void impBI(JSONObject vo, HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		String PK_ID = null;
		String sql = "";
		int result = 0;
		if (null != vo) {
			Connection conn = DbUtil.getConnection();
			DbUtil.startTrans(conn, "");
			PreparedStatement ps = null;
			PreparedStatement ps1 = null;
			String imptype = request.getAttribute("imptype").toString();

			// 先删除
			sql = "DELETE wb_erp.BI_EXCELTABLE WHERE IMP_TYPE = ?";
			ps1 = conn.prepareStatement(sql);
			DbUtil.setObject(ps1, 1, Types.CHAR, imptype);
			result = ps1.executeUpdate();

			// 插入
			sql = "INSERT INTO wb_erp.BI_EXCELTABLE "
					+ "(PK_ID, IMP_TYPE,DEF, DEF1, DEF2, DEF3, DEF4, DEF5, DEF6, DEF7, DEF8,ts) "
					+ "VALUES " + "(?, ?,?, ?, ?, ?, ?, ?, ?, ?, ?,sysdate)";
			ps = conn.prepareStatement(sql);
			PK_ID = SysUtil.getId();
			ps.setString(1, PK_ID);
			DbUtil.setObject(ps, 2, Types.CHAR, imptype);
			DbUtil.setObject(ps, 3, Types.VARCHAR, vo.opt("DEF"));
			DbUtil.setObject(ps, 4, Types.VARCHAR, vo.opt("DEF1"));
			DbUtil.setObject(ps, 5, Types.VARCHAR, vo.opt("DEF2"));
			DbUtil.setObject(ps, 6, Types.VARCHAR, vo.opt("DEF3"));
			DbUtil.setObject(ps, 7, Types.VARCHAR, vo.opt("DEF4"));
			DbUtil.setObject(ps, 8, Types.VARCHAR, vo.opt("DEF5"));
			DbUtil.setObject(ps, 9, Types.VARCHAR, vo.opt("DEF6"));
			DbUtil.setObject(ps, 10, Types.VARCHAR, vo.opt("DEF7"));
			DbUtil.setObject(ps, 11, Types.VARCHAR, vo.opt("DEF8"));
			result = ps.executeUpdate();

			// 提交事务
			conn.commit();

			// 关闭资源
			DbUtil.closeStatement(ps1);
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