/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package exceltesting;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.scene.control.Cell;
import org.apache.poi.xssf.usermodel.*;

/**
 *
 * @author JOSEPH
 */
public class SpreadSheetLoader {
    Object [] objectArr = null;
    String name;
    Map <String, Object[]> worksheetInfo;
    int rowCount = 0;
    XSSFWorkbook workbook = null;
    XSSFSheet spreadsheet = null;
    public SpreadSheetLoader() {
        
    }
    
    public SpreadSheetLoader(String name,Object[] obj) {
        this.name = name;
        Initialize(obj);
    }
    
    

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public Map<String, Object[]> getWorksheetInfo() {
        return worksheetInfo;
    }

    public void setWorksheetInfo(Map<String, Object[]> worksheetInfo) {
        this.worksheetInfo = worksheetInfo;
    }
    
    
    
    private void Initialize(Object[] obj){
        worksheetInfo = new TreeMap < String, Object[] >();
        //create a blank workbook
        workbook = new XSSFWorkbook();
        //Create a blank sheet
        spreadsheet = workbook.createSheet("Employee Info");
        
        putNewRowIntoSpreadSheet(obj);
    }
    
    
    public void putNewRowIntoSpreadSheet(Object[] obj){
        //This is just putting info into the tree.
        worksheetInfo.put( Integer.toString(rowCount), obj);
        XSSFRow row;
        
        
        row = spreadsheet.createRow(rowCount);
        //get corresponding value in tree using the count
        
        objectArr = worksheetInfo.get(Integer.toString(rowCount));
        rowCount++;
        //This is used to itereate across the length of the object array.
        int count = 0;
        //The object array is taken here and its contents placed in the new row.
        for (Object objects : objectArr){
              XSSFCell cell = row.createCell(count);
              cell.setCellValue((String)objects);
              count++;
        }
    }
    
    public void CloseWorkBook(FileOutputStream out){
        try {
            workbook.write(out);
        } catch (IOException ex) {
            Logger.getLogger(SpreadSheetLoader.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
