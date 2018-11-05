package com.sbt.tool;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.text.CollationKey;
import java.text.Collator;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Map.Entry;
import java.util.concurrent.ConcurrentHashMap;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import org.json.JSONArray;
import org.json.JSONObject;

import com.webbuilder.common.Session;
import com.webbuilder.tool.PageInfo;
import com.webbuilder.utils.DbUtil;
import com.webbuilder.utils.StringUtil;
import com.webbuilder.utils.WebUtil;
/**
 * yangbo 客户网修改_2014-11
 * @author Administrator
 * @获取树形节点工具类
 * 
 */
public class TreeTool {
	private static List<Entry<String, String>> sortRole(Map<String, String> map) {
		List<Entry<String, String>> list = new ArrayList<Entry<String, String>>(
				map.entrySet());
		Collections.sort(list, new Comparator<Entry<String, String>>() {
			Collator collator = Collator.getInstance();

			public int compare(Entry<String, String> e1,
					Entry<String, String> e2) {
				CollationKey key1 = collator.getCollationKey(StringUtil
						.getValuePart(e1.getValue()).toLowerCase());
				CollationKey key2 = collator.getCollationKey(StringUtil
						.getValuePart(e2.getValue()).toLowerCase());
				return key1.compareTo(key2);
			}
		});
		return list;
	}
	
	/**
	 * 产品二开用户模块业务单元树
	 * 获取组织表数据集
	 * getZZBTree
	 * @param request
	 * @param response
	 * @throws Exception
	 */
	public static void getZZBTree(HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		String v, id, pid = request.getParameter("itemId");
		StringBuilder buf = new StringBuilder();
		ConcurrentHashMap<String, String> roles = TreeDb.getZzbsMap();
		List<Entry<String, String>> es = sortRole(roles);
		boolean isFirst = true, check = StringUtil.getBool(WebUtil.fetch(
				request, "check"));

		if (StringUtil.isEmpty(pid))
			pid = "-1";
		buf.append("{children:[");
		for (Entry<String, String> e : es) {
			v = e.getValue();
			if (StringUtil.getNamePart(v).equals(pid)) {
				if (isFirst)
					isFirst = false;
				else
					buf.append(',');
				id = e.getKey();
				buf.append("{text:");
				buf.append(StringUtil.quote(getValuePart(v)));
				buf.append(",CATSNAMEEN:");
				buf.append(StringUtil.quote(getValue(v)));
				buf.append(",itemId:\"");
				buf.append(id);
				buf.append("\",iconCls:\"product_icon\"");
				if (!hasSubRole(es, id))
					buf.append(",children:[]");
				if (check)
					buf.append(",checked:false");
				buf.append('}');
			}
		}
		buf.append("]}");
		WebUtil.response(response, buf);
	}
	
	/**
	 * by yangbo 组织单元树修改
	 * @param request
	 * @param response
	 * @throws Exception
	 */
	public static void UpdateZzbTree(HttpServletRequest request,
			HttpServletResponse response) throws Exception {

		DbUtil.update(request,
				"update WB_SBT_ZZB set ZZNAME={?name?}, STATE={?CATSNAMEEN?} where ZZBID={?id?}");

	}
	 /**
	  * 增加组织单元
	  * @param request
	  * @param response
	  * @throws Exception
	  */
	public static void appendZzbTree(HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		//ConcurrentHashMap<String, String> buffer = TreeDb.getRoleMap();
		String id = (String) request.getAttribute("sys.id");
		DbUtil.update(request,
				"insert into WB_SBT_ZZB (ZZBID ,ZZNAME, ZZBPID,STATE,CATSORDER) " +
				"values({?sys.id?},{?name?},{?parentId?},1,0)");
//		buffer.put(id, request.getParameter("parentId") + "="
//				+ request.getParameter("name"));
		WebUtil.response(response, StringUtil.concat("{id:'", id, "'}"));
	}
	/**
	 * 产品模块树
	 * 获取树形节点数据集
	 * 产品getCatsTree
	 * @param request
	 * @param response
	 * @throws Exception
	 */
	public static void getCatsTree(HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		String v, id, pid = request.getParameter("itemId");
		StringBuilder buf = new StringBuilder();
		ConcurrentHashMap<String, String> roles = TreeDb.getCatsMap();
		List<Entry<String, String>> es = sortRole(roles);
		boolean isFirst = true, check = StringUtil.getBool(WebUtil.fetch(
				request, "check"));

		if (StringUtil.isEmpty(pid))
			pid = "-1";
		buf.append("{children:[");
		for (Entry<String, String> e : es) {
			v = e.getValue();
			if (StringUtil.getNamePart(v).equals(pid)) {
				if (isFirst)
					isFirst = false;
				else
					buf.append(',');
				id = e.getKey();
				buf.append("{text:");
				buf.append(StringUtil.quote(getValuePart(v)));
				buf.append(",CATSNAMEEN:");
				buf.append(StringUtil.quote(getValue(v)));
				buf.append(",itemId:\"");
				buf.append(id);
				buf.append("\",iconCls:\"product_icon\"");
				if (!hasSubRole(es, id))
					buf.append(",children:[]");
				if (check)
					buf.append(",checked:false");
				buf.append('}');
			}
		}
		buf.append("]}");
		WebUtil.response(response, buf);
	}
	
	/**
	 * 获取首字母
	 * @param string
	 * @return
	 */
	public static String getValue(String string) {
		if (string == null)
			return "";
		int index = string.indexOf('@');

		if (index == -1)
			return "";
		else
			return string.substring(index + 1);
	}
	/**
	 * 获取内容
	 * @param string
	 * @return
	 */
	public static String getValuePart(String string) {
		if (string == null)
			return "";
		int indexF = string.indexOf('=');
		int indexE = string.indexOf('@');
		if (indexF == -1)
			return "";
		else
			return string.substring(indexF + 1,indexE);
	}
	 /**
	  * 增加产品分类
	  * @param request
	  * @param response
	  * @throws Exception
	  */
	public static void appendCatsTree(HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		//ConcurrentHashMap<String, String> buffer = TreeDb.getRoleMap();
		String id = (String) request.getAttribute("sys.id");
		DbUtil.update(request,
				"insert into cn_t_cats (CATSID ,CATSNAME,CATSNAMEEN ,CATSPID,STATE,CATSORDER) " +
				"values({?sys.id?},{?name?},{?CATSNAMEEN?},{?parentId?},1,0)");
//		buffer.put(id, request.getParameter("parentId") + "="
//				+ request.getParameter("name"));
		WebUtil.response(response, StringUtil.concat("{id:'", id, "'}"));
	}
	
	/**
	 * by yangbo 产品分类树修改
	 * @param request
	 * @param response
	 * @throws Exception
	 */
	public static void UpdateCatsTree(HttpServletRequest request,
			HttpServletResponse response) throws Exception {
	//	ConcurrentHashMap<String, String> buffer = Role.getRoleMap2();
//		String id = request.getParameter("id"), name = request
//				.getParameter("name"), val;
		DbUtil.update(request,
				"update CN_T_CATS set CATSNAME={?name?}, CATSNAMEEN={?CATSNAMEEN?} where CATSID={?id?}");
//		val = buffer.get(id);
//		if (val != null)
//			buffer.put(id, StringUtil.getNamePart(val) + "=" + name);
//		else
//			SysUtil.error(Str.format(request, "notExist", name));
	}
	
	/**
	 * 
	 * @param es
	 * @param id
	 * @return
	 */
	private static boolean hasSubRole(List<Entry<String, String>> es, String id) {
		for (Entry<String, String> e : es) {
			if (StringUtil.isEqual(StringUtil.getNamePart(e.getValue()), id))
				return true;
		}
		return false;
	}
	
	/**
	 * 获取组织下的用户
	 * @param request
	 * @param response
	 * @throws Exception
	 */
	public static void getUsers(HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		Connection conn = DbUtil.getConnection(request);
		StringBuilder buf = new StringBuilder();
		String orgFields[] = { "USER_NAME", "DISPLAY_NAME", "STATUS",
				"CREATE_DATE", "LOGIN_TIMES", "EMAIL", "LAST_LOGIN" };
		String mapFields[] = { "userName", "dispName", "status", "createDate",
				"loginTimes", "email", "lastLogin" };
		HashMap<String, JSONArray> roleMap = new HashMap<String, JSONArray>();
		JSONArray rows = new JSONArray(), ra;
		JSONObject jo;
		String uid, sortInfo[], orderBy, where, findName, findRole;
		PageInfo pageInfo;
		ResultSet userRs = null, roleRs = null;
		PreparedStatement roleSt = null, userSt = null;
		HttpSession session;
		ConcurrentHashMap<String, HttpSession> sessionList = Session.sessionList;
		int i, len, cp;
		boolean allRoles;
		HttpSession session2 = request.getSession(true), prevSession;
		 String userName=(String)session2.getAttribute("sys.userName");
		try {
			sortInfo = WebUtil.getSortInfo(request);
			if (sortInfo == null)
				orderBy = "a.STATUS desc,a.USER_NAME";
			else
				orderBy = "a."
						+ orgFields[StringUtil.indexOf(mapFields, sortInfo[0])]
						+ " " + sortInfo[1];

			findRole = request.getParameter("findRole");
			allRoles = StringUtil.isEqual(findRole, "-1");
			if (!allRoles && !StringUtil.isEmpty(findRole)) {
				userSt = conn
						.prepareStatement("select a.USER_ID,a.USER_NAME,a.DISPLAY_NAME,a.STATUS,a.CREATE_DATE,a.LOGIN_TIMES,a.EMAIL,a.LAST_LOGIN from WB_USER a,WB_USER_ROLE b,wb_sbt_zzb t1 where a.USER_ID=b.USER_ID and a.bm_pk=t1.zzbid and b.ROLE_ID=? order by "
								+ orderBy);
				userSt.setString(1, findRole);
			} else {
			  
				if (allRoles)
					findName = null;
				else
					findName = request.getParameter("findName");
				   String zzbid= request.getParameter("lb");
				if (StringUtil.isEmpty(findName))
				{
					where = "";
				}
					else
					{
			     
					where = " and a.USER_NAME like ?";
					}
					String st="select a.USER_ID,a.USER_NAME,a.DISPLAY_NAME,a.STATUS,a.CREATE_DATE,a.LOGIN_TIMES,a.EMAIL,a.LAST_LOGIN,a.ZCPK,a.BM_PK from WB_USER a, wb_sbt_zzb t1" +
							" where a.bm_pk=t1.zzbid and a.bm_pk ='"+zzbid+"' "//and a.ZCR='"+userName+"'
					+ where + " order by " + orderBy;
				userSt = conn
						.prepareStatement(st);
				if (!StringUtil.isEmpty(findName))
					
					userSt.setString(1, findName + "%");
				   
			}
			userRs = userSt.executeQuery();
			buf.append(",onlines:");
			buf.append(sessionList.size());
			buf.append(",totalUser:");
			buf.append(request.getAttribute("users.CT"));
			buf.append(",rows:");
			pageInfo = WebUtil.getPage(request);
			while (userRs.next()) {
				cp = WebUtil.checkPage(pageInfo);
				if (cp == 1)
					break;
				else if (cp == 2)
					continue;
				uid = userRs.getString(1);
				jo = new JSONObject();
				jo.put("userId", uid);
				jo.put("userName", userRs.getString(2));
				jo.put("dispName", userRs.getString(3));
				jo.put("status", userRs.getInt(4));
				jo.put("createDate", userRs.getTimestamp(5));
				jo.put("loginTimes", userRs.getInt(6));
				jo.put("email", userRs.getString(7));
				jo.put("lastLogin", userRs.getTimestamp(8));
				jo.put("ZCBH", userRs.getString(9));
				jo.put("DEPBH", userRs.getString(10));
				session = sessionList.get(uid);
				if (session != null) {
					jo.put("ip", session.getAttribute("sys.userIp"));
					jo.put("on", 1);
				}
				ra = new JSONArray();
				jo.put("roles", ra);
				roleMap.put(uid, ra);
				rows.put(jo);
			}
			len = roleMap.size();
			if (len == 0)
				buf.append("[]");
			else {
				roleSt = conn
						.prepareStatement("select a.USER_ID,a.ROLE_ID,b.ROLE_NAME from WB_USER_ROLE a, WB_ROLE b where a.ROLE_ID=b.ROLE_ID and a.USER_ID in("
								+ StringUtil.duplicate("?,", len - 1) + "?)");
				Set<Entry<String, JSONArray>> es = roleMap.entrySet();
				i = 1;
				for (Entry<String, JSONArray> e : es) {
					roleSt.setString(i++, e.getKey());
				}
				roleRs = roleSt.executeQuery();
				while (roleRs.next()) {
					ra = roleMap.get(roleRs.getString(1));
					ra.put(roleRs.getString(2) + "=" + roleRs.getString(3));
				}
				buf.append(rows);
			}
			buf.append('}');
			WebUtil.setTotal(buf, pageInfo);
			WebUtil.response(response, buf);
		} finally {
			DbUtil.closeResultSet(userRs);
			DbUtil.closeResultSet(roleRs);
			DbUtil.closeStatement(userSt);
			DbUtil.closeStatement(roleSt);
		}
	}

}
