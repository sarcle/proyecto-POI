package com.magnabyte.POI;

import java.io.File;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcel 
{
    public static void main(String[] args) 
    {
        //Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook(); 
         
        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Employee Data");
          
        //This data needs to be written (Object[])
        
        Persona persona = new Persona();
        persona.setNombre("JUAN");
        persona.setApellido("PEREZ");
        persona.setEdad(12);
        Map<String, Persona> data = new HashMap<String, Persona>();
//        data.put("1", new Object[] {"ID", "NAME", "LASTNAME"});
//        data.put("2", new Object[] {1, "Amit", "Shukla"});
//        data.put("3", new Object[] {2, "Lokesh", "Gupta"});
//        data.put("4", new Object[] {3, "John", "Adwards"});
//        data.put("5", new Object[] {4, "Brian", "Schultz"});
        data.put("6", persona);
        data.put("7", persona);
        data.put("8", persona);
        data.put("9", persona);
        data.put("10", persona);
        
          
        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset)
        {
            Row row = sheet.createRow(rownum++);
            Persona objArr = data.get(key);
            int cellnum = 0;
            
            for(int a =0; a <=data.size(); a++)  {
            	Cell cell = row.createCell(0);
            	cell.setCellValue(objArr.getNombre());
            	cell = row.createCell(1);
            	cell.setCellValue(objArr.getApellido());
            	cell = row.createCell(2);
            	cell.setCellValue(objArr.getEdad());
            }
            
            
//			for (Object obj : objArr2)
//            {
//               Cell cell = row.createCell(cellnum++);
//               if(obj instanceof String)
//                    cell.setCellValue((String)obj);
//                else if(obj instanceof Integer)
//                    cell.setCellValue((Integer)obj);
//               System.out.println("cellNum ---> " + cellnum);
//            }
        }
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("C:/Users/Roberto/Documents/EJEMPLO-EXCEL.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
    }
}
