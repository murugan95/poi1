package com.howtodoinjava.demo.poi;


import java.sql.*;  
public class Sampledb{  
public static void main(String args[]){  
try{  
//step1 load the driver class  
Class.forName("oracle.jdbc.driver.OracleDriver");  
  
//step2 create  the connection object  
Connection con=DriverManager.getConnection(  
"jdbc:oracle:thin:@10.1.4.70:1521:OD12CR11","CATS_TEST_20190613","CATS_TEST_20190613");  
  
//step3 create the statement object  
Statement stmt=con.createStatement();
  
//step4 execute query  
ResultSet rs=stmt.executeQuery("select * from student ");  
while(rs.next())  
System.out.println(rs.getInt(1)+"  "+rs.getString(2)+"  "+rs.getString(3));  
  
//step5 close the connection object  
con.close();  
  
}catch(Exception e){ System.out.println(e);}  
  
}  
}  