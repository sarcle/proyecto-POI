package com.magnabyte.POI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
    	try {
    	     
    	    FileInputStream file = new FileInputStream(new File("C:/Users/Roberto/Documents/test.xlsx"));
    	     
    	    
    	    XSSFWorkbook workbook = new XSSFWorkbook (file);
    	     
    	    XSSFSheet sheet = workbook.getSheetAt(0);
    	     
    	    Iterator<Row> rowIterator = sheet.iterator();
    	     
    	    while(rowIterator.hasNext()) {
    	        Row row = rowIterator.next();
    	         
    	        Iterator<Cell> cellIterator = row.cellIterator();
    	        while(cellIterator.hasNext()) {
    	             
    	            Cell cell = cellIterator.next();
    	             
    	            switch(cell.getCellType()) {
    	                case Cell.CELL_TYPE_BOOLEAN:
    	                    System.out.print(cell.getBooleanCellValue() + "\t\t");
    	                    break;
    	                case Cell.CELL_TYPE_NUMERIC:
    	                    System.out.print(cell.getNumericCellValue() + "\t\t");
    	                    break;
    	                case Cell.CELL_TYPE_STRING:
    	                    System.out.print(cell.getStringCellValue() + "\t\t");
    	                    break;
    	            }
    	        }
    	        System.out.println("");
    	    }
    	    file.close();
    	    FileOutputStream out = 
    	        new FileOutputStream(new File("C:\\test.xls"));
    	    workbook.write(out);
    	    out.close();
    	     
    	} catch (FileNotFoundException e) {
    	    e.printStackTrace();
    	} catch (IOException e) {
    	    e.printStackTrace();
    	}
    }
}
