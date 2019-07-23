package com.howtodoinjava.demo.poi;
import java.io.FileInputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Vector;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToDb {
    public static void main( String [] args ) {
        String fileName="m.xlsx";
        Vector<List<XSSFCell>> dataHolder=read(fileName);
        saveToDatabase(dataHolder);
    }
    public static Vector<List<XSSFCell>> read(String fileName)    {
        Vector<List<XSSFCell>> cellVectorHolder = new Vector<List<XSSFCell>>();
        try{
            FileInputStream myInput = new FileInputStream(fileName);
            //POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
            @SuppressWarnings("resource")
			XSSFWorkbook myWorkBook = new XSSFWorkbook(myInput);
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);
            Iterator<?> rowIter = mySheet.rowIterator();
            while(rowIter.hasNext()){
                XSSFRow myRow = (XSSFRow) rowIter.next();
                Iterator<?> cellIter = myRow.cellIterator();
                //Vector cellStoreVector=new Vector();
                List<XSSFCell> list = new ArrayList<XSSFCell>();
                while(cellIter.hasNext()){
                    XSSFCell myCell = (XSSFCell) cellIter.next();
                    list.add(myCell);
                }
                cellVectorHolder.addElement(list);
            }
        }catch (Exception e){e.printStackTrace(); }
        return cellVectorHolder;
    }
    private static void saveToDatabase(Vector<List<XSSFCell>> dataHolder) {
    	 try {
    		 Class.forName ("oracle.jdbc.OracleDriver"); 
             Connection con = DriverManager.getConnection("jdbc:oracle:thin:@10.1.4.70:1521:OD12CR11", "CATS_TEST_20190613", "CATS_TEST_20190613");
             System.out.println("connection made...");
             PreparedStatement stmt=con.prepareStatement("INSERT INTO STUDENT(ID,FIRSTNAME,LASTNAME) VALUES(?,?,?)");
        
             
        String ID="";
        String FIRSTNAME="";
        String LASTNAME="";
       
        for(Iterator<List<XSSFCell>> iterator = dataHolder.iterator();iterator.hasNext();) {
            List<?> list = (List<?>) iterator.next();
            ID = list.get(0).toString();
            FIRSTNAME = list.get(1).toString();
            LASTNAME = list.get(2).toString();
                stmt.setString(1, ID);
                stmt.setString(2, FIRSTNAME);
                stmt.setString(3, LASTNAME);
               
                stmt.execute();
              
               
                
        
        }
      
        System.out.println(dataHolder);
        stmt.close();
        con.close();
        System.out.println("Data is inserted");
            } 
    	 catch (Exception e) {
                e.printStackTrace();
       }}}
    