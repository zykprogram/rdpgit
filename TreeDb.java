package com.sbt.tool;

import java.sql.Connection;
import java.sql.ResultSet;
import java.util.concurrent.ConcurrentHashMap;

import com.webbuilder.utils.DbUtil;

public class TreeDb {
	private static ConcurrentHashMap<String, String> buffer;
	private static ConcurrentHashMap<String, String> buffer2;

//	public static ConcurrentHashMap<String, String> getRoleMap()
//			throws Exception {
//		if (buffer == null)
//			initialize(false);
//		return buffer;
//	}
//
//	public static synchronized void initialize(boolean reload) throws Exception {
//		
//		if (!reload && buffer != null)
//			return;
//		buffer = new ConcurrentHashMap<String, String>();
//		Connection conn = null;
//		ResultSet rs = null;
//
//		try {
//			conn = DbUtil.getConnection();
//			rs = DbUtil.getResultSet(conn, "select * from WB_ROLE");
//			while (rs.next()) {
//				buffer.put(rs.getString(1), rs.getString(2) + "="
//						+ rs.getString(3));
//			}
//		} finally {
//			DbUtil.closeResultSet(rs);
//			DbUtil.closeConnection(conn);
//		}
//	}
	/**
	 *查询产品
	 */
	public static ConcurrentHashMap<String, String> getCatsMap()
	   throws Exception {
		buffer2 = new ConcurrentHashMap<String, String>();
		Connection conn = null;
		ResultSet rs = null;

		try {
			conn = DbUtil.getConnection();
			rs = DbUtil.getResultSet(conn, "select * from cn_T_Cats");
			while (rs.next()) {
				buffer2.put(rs.getString(1),rs.getString(4)+ "="
						+ rs.getString(2)+"@"+rs.getString(3));
			}
		} finally {
			DbUtil.closeResultSet(rs);
			DbUtil.closeConnection(conn);
		}
		return buffer2;
	}
	/**
	 *查询业务单元
	 */
	public static ConcurrentHashMap<String, String> getZzbsMap()
	   throws Exception {
		buffer= new ConcurrentHashMap<String, String>();
		Connection conn = null;
		ResultSet rs = null;

		try {
			conn = DbUtil.getConnection();
			rs = DbUtil.getResultSet(conn, "select * from WB_SBT_ZZB");
			while (rs.next()) {
				buffer.put(rs.getString(1),rs.getString(3)+ "="
						+ rs.getString(2)+"@"+rs.getString(4));
			}
		} finally {
			DbUtil.closeResultSet(rs);
			DbUtil.closeConnection(conn);
		}
		return buffer;
	}
}
