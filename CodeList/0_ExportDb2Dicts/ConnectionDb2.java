package com.amarscf.app.lrsong.export;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;


/*
 * @author lrsong
 * 模仿制作没有设计模式的数据字典导出工具
 * 
 * */
public class ConnectionDb2 {
	private String sqlDriver = "com.ibm.db2.jcc.DB2Driver";
	private String url = "jdbc:db2://10.125.188.81:60000/amarscf";//url为连接字符串
	private String user = "amarscf";//数据库用户名
	private String pwd = "amarscf";//数据库密码

    private static Connection conn = null;
	private static Statement stmt = null;
	private static ResultSet rs = null;
	
	// 数据库连接
	public ConnectionDb2(){
		try{
			Class.forName(sqlDriver);
		}catch(ClassNotFoundException e1){
			e1.printStackTrace();
			System.err.println(e1.getMessage()+"数据库连接问题");
		}
	}
	// 数据查询
	public ResultSet executeQuery(String sql) throws SQLException{
		conn = DriverManager.getConnection(url, user, pwd);  
        Statement stmt = conn.createStatement();  
        rs = stmt.executeQuery(sql);  
        return rs;  
	}
	//数据更新 -- 基本用不到
	public void executeUpdate(String sql) throws SQLException{
        conn = DriverManager.getConnection(url, user, pwd);
        Statement stmt = conn.createStatement();
        stmt.executeUpdate(sql);
    }
    public void close() throws SQLException{
        if (rs != null)
            rs.close();
        if (stmt != null)
            stmt.close();
        if (conn != null)
            conn.close();
    }
    /** 
     * @param args 
     * 测试 oracle数据连接是否联通 
     */  
    public static void main(String[] args) {  

    }
}
