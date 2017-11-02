
package com.iesvdc.acceso.excelapi.excelapi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *Esta clase almacena información de libros para generar ficheros de Excel.
 * Un libro se compone de hojas.
 * 
 * @author Daniel Sierra Ráez
 */
public class Libro {
    private List<Hoja> hojas;
    private String nombreArchivo;
    private Hoja hoja = new Hoja();

    
    public Libro() {
        this.hojas = new ArrayList<>();
        this.nombreArchivo = "nuevo.xlsx";
    }

    public Libro(String nombreArchivo) {
        this.hojas = new ArrayList<>();
        this.nombreArchivo = nombreArchivo;      
    }

    public String getNombreArchivo() {
        return nombreArchivo;
    }

    public void setNombreArchivo(String nombreArchivo) {
        this.nombreArchivo = nombreArchivo;
    }
    
    /**
     * Añade una hoja
     * @param hoja 
     * @return  
     */    
    public boolean addHoja(Hoja hoja){
        return this.hojas.add(hoja);
    }
      
    /**
     * Borra una hoja
     * @param index
     * @return 
     * @throws com.iesvdc.acceso.excelapi.excelapi.ExcelAPIException 
     */   
    public Hoja removeHoja(int index) throws ExcelAPIException{
        if(index < 0 || index > this.hojas.size()){
            throw new ExcelAPIException("Libro()::removeHoja(): Posición no válida ");
        }
        return this.hojas.remove(index);
    }
    
    /**
     * Devuelve la posición de la hoja
     * @param index
     * @return 
     * @throws com.iesvdc.acceso.excelapi.excelapi.ExcelAPIException 
     */         
    public Hoja indexHoja(int index) throws ExcelAPIException{
        if(index < 0 || index > this.hojas.size()){
            throw new ExcelAPIException("Libro()::indexHoja(): Posición no válida ");
        }            
       return this.hojas.get(index);
    }
    
   /**
    * Carga el contenido de un libro con extensión "xlsx".
    * @throws com.iesvdc.acceso.excelapi.excelapi.ExcelAPIException 
    */
    public void load() throws ExcelAPIException{        
        //Inicializamos a null el valor de entrada   
        FileInputStream fichero = null;
           
        try {
            //Creamos un archivo con el nombreArchivo
            File archivo = new File(this.nombreArchivo);            
            fichero = new FileInputStream(archivo);
            
            //Busca el Libro instanciado por XLSX "archivo".            
            XSSFWorkbook myWorkBook = new XSSFWorkbook(fichero) ;
            
            //Si el contenido de las hojas esta lleno que lo borre y si no crea una nueva lista
            if (this.hojas != null){
                if( this.hojas.size() > 0 ){
                    this.hojas.clear();
                }
                
            } else {
                this.hojas = new ArrayList<>();
                
            }
            
            //Bucle que recorre las hojas
            for (int i = 0; i < myWorkBook.getNumberOfSheets();i++){
                Sheet sheetXlsx = myWorkBook.getSheetAt(i);
                
                int rowNum = sheetXlsx.getLastRowNum()+1;
                int colNum = 0;
                    //Bucle que recorre las filas
                    for ( int j =0; j < sheetXlsx.getLastRowNum(); j++){
                      Row rowXlsx = sheetXlsx.getRow(j);
                                //Si el numero de columnas es menor que la fila ** la iguala
                                if (colNum < rowXlsx.getLastCellNum()) {
                                    colNum = rowXlsx.getLastCellNum();
                                }
                    }
                    //Comprobamos que las hojas estan cargadas
                     System.out.println("libro.load():: datosDeHoja= " + sheetXlsx.getSheetName());
                     Hoja datosHoja = new Hoja(sheetXlsx.getSheetName(), rowNum, colNum);
                     
                     //Recorremos las filas
                     for (int j = 0; j < rowNum; j++){
                         Row rowXlsx = sheetXlsx.getRow(j);
                            //Recorremos las celdas
                            for (int k = 0; k < rowXlsx.getLastCellNum(); k++){
                                Cell cellXlsx = rowXlsx.getCell(k);
                                String datos = " ";
                                    //Si las celdas estan llenas comprobamos el tipo
                                    if(cellXlsx != null){                                        
                                          switch (cellXlsx.getCellType()){
                                              case Cell.CELL_TYPE_STRING:
                                              datos = cellXlsx.getStringCellValue();
                                              break;
                                              
                                              case Cell.CELL_TYPE_NUMERIC:
                                              datos += cellXlsx.getNumericCellValue();
                                              break;
                                              
                                              case Cell.CELL_TYPE_BOOLEAN:
                                              datos += cellXlsx.getBooleanCellValue();
                                               break;
                                              
                                              case Cell.CELL_TYPE_FORMULA:
                                              datos += cellXlsx.getCellFormula();
                                               break;
                                               
                                              default :
                                              datos = " ";
                                          }
                                          //Compruebo que el contenido de las hojas esta cargado
                                          System.out.println("libro.load  j= " + j +  " k= " + k + " datos= " + datos);
                                          datosHoja.setDatos(datos, j , k);                                         
                                          
                                    }                        
                              } 
                            //Añado los datos a Hoja
                            this.hojas.add(datosHoja);
                     }
            }
            
        } catch (IOException ex) {          
            throw new ExcelAPIException("libro IO Error cuando cargas archivo");
            
        } finally {
            try {
                    if (fichero != null){ 
                        fichero.close();
                    }
                
            } catch (IOException ex) {
                throw new ExcelAPIException("libro IO Error cuando cargas archivo");
            }
        }                 
    }
    
    public void load(String filename) throws ExcelAPIException{
        this.nombreArchivo = filename;
        this.load();
    }
    
    /**
     * Método que guarda el contenido de las hojas en el libro.
     * @throws  com.iesvdc.acceso.excelapi.excelapi.ExcelAPIException 
     */
    public void save() throws ExcelAPIException{       
        SXSSFWorkbook wb = new SXSSFWorkbook();
    
        for (Hoja hoja: this.hojas) {  
             Sheet sh = wb.createSheet(hoja.getTitulo());       
             
            for (int i = 0; i < hoja.getFilas(); i++) {
                Row row = sh.createRow(i);
                
                for (int j = 0; j < hoja.getColumnas(); j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(hoja.getDatos(i, j));                          
                }
            }   
        }            
                
        try (FileOutputStream out = new FileOutputStream(this.nombreArchivo)) {         
            wb.write(out);
                   
        } catch (IOException ex) {           
            throw new ExcelAPIException("Error al guardar el arhivo");
          
        } finally {
            wb.dispose();
        }
        
    }
    
    public void save(String filename) throws ExcelAPIException{
        this.nombreArchivo = filename;
        this.save();

    }
    
    
    /**
     * Método que comprueba la extensión del fichero "xlsx".
     * 
     */
    public void extension() { 
        String nombre = "";    
        String extensionPunto = ".xlsx";
        String extension = "xlsx";       
        int i = this.nombreArchivo.lastIndexOf('.');
        
        //Si es menor que 0 significa que el punto de la extensión no esta añadido y hay que añadirlo
            if(i <  0){
                nombreArchivo = this.nombreArchivo+extensionPunto;
                i  = this.nombreArchivo.lastIndexOf('.');                
            }
          
         nombre = this.nombreArchivo.substring(i+1);
         /*Despues de que este añadido el punto de la extensión
            compruebo que la extensión sea xlsx, si es otra o no esta añadida, la añado
         */
            if (!"xlsx".equals(nombre)){       
                nombreArchivo = this.nombreArchivo.substring(0,i+1) + extension;    
            }       
            
    }      
    
}