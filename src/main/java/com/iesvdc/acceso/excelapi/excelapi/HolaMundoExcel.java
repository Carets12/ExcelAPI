/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.iesvdc.acceso.excelapi.excelapi;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 *
 * @author Daniel Sierra Ráez
 */

public class HolaMundoExcel {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        
        SXSSFWorkbook wb = new SXSSFWorkbook();
    
        Sheet sh = wb.createSheet("HOLA MUNDO");

        for (int i = 0; i < 10; i++) {
            Row row = sh.createRow(i);
            for (int j = 0; j < 10 ; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue((char)('A'+j)+" "+(i+1));      
            }
        }   
         
        try {
            FileOutputStream out = new FileOutputStream("holaMundoExcel.xlsx");
            wb.write(out);
            out.close();       
            
        } catch (IOException ex) {
            System.out.println("ERROR al crear el archivo: "+
                    ex.getLocalizedMessage());
            
        } finally {
            wb.dispose();
        }         
    }    
}
