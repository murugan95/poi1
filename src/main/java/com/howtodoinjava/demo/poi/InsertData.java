package com.howtodoinjava.demo.poi;

import java.sql.*  ;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.HashMap;
import java.util.Iterator;




import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.Statement;
import java.util.Iterator;
import java.util.Vector;

import javax.swing.plaf.synth.SynthOptionPaneUI;

//import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
 
 
public class InsertData{
 
        public static void main(String[] args) {
 
             //   String fileName = "C:\\excelFile.xls";
            //  String fileName ="C:/cyril/For Rajkumar//expense.xls";
            String fileName ="Lakshmi.xlsx";
                Vector dataHolder = ReadCSV(fileName);
                printCellDataToConsole(dataHolder);
        }
 
        public static Vector ReadCSV(String fileName) {
                Vector cellVectorHolder = new Vector();
 
                try {
                        FileInputStream myInput = new FileInputStream(fileName);
                       // System.out.println("Input:"+myInput);
                        Workbook wb = WorkbookFactory.create(myInput);
 
                        Sheet sheet = wb.getSheetAt(0);
 
                        Iterator rowIter = sheet.rowIterator();
 
                        while (rowIter.hasNext()) {
                                Row myRow = (Row) rowIter.next();
                                Iterator cellIter = myRow.cellIterator();
                                Vector cellStoreVector = new Vector();
                                while (cellIter.hasNext()) {
                                        Cell myCell = (Cell) cellIter.next();
                                        cellStoreVector.addElement(myCell);
                                }
                                cellVectorHolder.addElement(cellStoreVector);
                        }
                } catch (Exception e) {
                        e.printStackTrace();
                }
                return cellVectorHolder;
        }
 
        @SuppressWarnings("deprecation")
		private static void printCellDataToConsole(@SuppressWarnings("rawtypes") Vector dataHolder) {
        	try{
        		 /* Create Connection objects */
                Class.forName ("oracle.jdbc.OracleDriver"); 
                Connection conn = DriverManager.getConnection("jdbc:oracle:thin:@10.1.4.70:1521:OD12CR11", "CATS_TEST_20190613", "CATS_TEST_20190613");
                PreparedStatement sql_statement = null;
                String jdbc_insert_sql = "INSERT INTO STUDENT"
                                + "(ID,NAME,LASTNAME) VALUES"
                                + "(?,?,?)";
                sql_statement = conn.prepareStatement(jdbc_insert_sql);
                /* We should now load excel objects and loop through the worksheet data */
                FileInputStream input_document = new FileInputStream(new File("Lakshmi.xlsx"));
                /* Load workbook */
                @SuppressWarnings("resource")
				XSSFWorkbook my_xls_workbook = new  XSSFWorkbook(input_document);
                /* Load worksheet */
                XSSFSheet my_worksheet = my_xls_workbook.getSheetAt(0);
                // we loop through and insert data
                Iterator<Row> rowIterator = my_worksheet.iterator(); 
                while(rowIterator.hasNext()) {
                        Row row = rowIterator.next(); 
                        Iterator<Cell> cellIterator = row.cellIterator();
                                while(cellIterator.hasNext()) {
                                        Cell cell = cellIterator.next();
                                        switch(cell.getCellType()) { 
                                        case Cell.CELL_TYPE_STRING: //handle string columns
                                                sql_statement.setString(1, cell.getStringCellValue());                                                                                     
                                                break;
                                        case Cell.CELL_TYPE_NUMERIC: //handle double data
                                                sql_statement.setDouble(2,cell.getNumericCellValue() );
                                                break;
                                        }
System.out.println("some thing");                              }
                //we can execute the statement before reading the next row
                sql_statement.executeUpdate();
                }
                /* Close input stream */
                input_document.close();
                /* Close prepared statement */
                sql_statement.close();
                /* COMMIT transaction */
                conn.commit();
                /* Close connection */
                conn.close();
        	 

     
                    }
             catch(Exception e){
             	}   
        }
        	 
}

 