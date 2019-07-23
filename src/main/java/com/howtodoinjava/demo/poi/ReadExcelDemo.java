package com.howtodoinjava.demo.poi;

import java.io.File;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelDemo {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		 try

	        {
			 Class.forName ("oracle.jdbc.OracleDriver"); 
             Connection conn = DriverManager.getConnection("jdbc:oracle:thin:@10.1.4.70:1521:OD12CR11", "CATS_TEST_20190613", "CATS_TEST_20190613");
             PreparedStatement sql_statement = null;
             String jdbc_insert_sql = "INSERT INTO STUDENT(ID,FIRSTNAME,LASTNAME) VALUES(?,?,?)";
             sql_statement = conn.prepareStatement(jdbc_insert_sql);

	            FileInputStream file = new FileInputStream(new File("Lakshmi.xlsx"));

	 

	            //Create Workbook instance holding reference to .xlsx file

	            XSSFWorkbook workbook = new XSSFWorkbook(file);

	 

	            //Get first/desired sheet from the workbook

	            XSSFSheet sheet = workbook.getSheetAt(0);

	 

	            //Iterate through each rows one by one

	            Iterator<Row> rowIterator = sheet.iterator();

	            while (rowIterator.hasNext()) 

	            {

	                Row row = rowIterator.next();

	                //For each row, iterate through all the columns

	                Iterator<Cell> cellIterator = row.cellIterator();

	                 

	                while (cellIterator.hasNext()) 

	                {

	                    Cell cell = cellIterator.next();

	                    //Check the cell type and format accordingly

	                    switch (cell.getCellType()) 

	                    {

	                        case Cell.CELL_TYPE_NUMERIC:

	                            System.out.print(cell.getStringCellValue() + "\t");
	                            
	                            break;

	                        case Cell.CELL_TYPE_STRING:

	                            System.out.print(cell.getStringCellValue() + "\t");
	                            
	                            break;
	                            
	                         
	                    }
	                   
	                }
	               // sql_statement.executeUpdate();
	                System.out.println("");

	            }

	            file.close();

	        } 

	        catch (Exception e) 

	        {

	            e.printStackTrace();

	        }

	}

}
