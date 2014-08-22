package com.magnabyte.POI;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * Ejemplo sencillo de cómo crear una hoja Excel con POI
 * 
 * @author chuidiang
 * 
 */
public class EjemploCrearExcel {

    /**
     * Crea una hoja Excel y la guarda.
     * 
     * @param args
     */
    public static void main(String[] args) {
        // Se crea el libro
        HSSFWorkbook libro = new HSSFWorkbook();

        // Se crea una hoja dentro del libro
        HSSFSheet hoja = libro.createSheet();

        // Se crea una fila dentro de la hoja
        HSSFRow fila = hoja.createRow(1);

        // Se crea una celda dentro de la fila
        HSSFCell celda = fila.createCell((short) 0);

        // Se crea el contenido de la celda y se mete en ella.
        HSSFRichTextString texto = new HSSFRichTextString("escribir nuevo dato");
        celda.setCellValue(texto);

        // Se salva el libro.
        try {
            FileOutputStream elFichero = new FileOutputStream("C:/Users/Roberto/Documents/holamundo.xls");
            libro.write(elFichero);
            elFichero.close();
            System.out.println("LISTOOOOOOO");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
