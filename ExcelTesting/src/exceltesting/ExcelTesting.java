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
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.TreeMap;
import javafx.scene.control.Cell;
import org.apache.poi.xssf.usermodel.*;


/**
 *
 * @author JOSEPH
 */
public class ExcelTesting {
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args)throws SQLException, ClassNotFoundException, FileNotFoundException, IOException {
        
        
        SpreadSheetLoader spreadSheetLoader = new SpreadSheetLoader("ExcelWorksheetTest", new Object[]{"PROJECTID","PROJECTNAME","REVENUE"});
        
        SpreadSheetOperations spreadSheetOperations = new SpreadSheetOperations(spreadSheetLoader);
        
        spreadSheetOperations.FetchAllRecords();
        spreadSheetOperations.UpdateRecord(1, new Object[]{"3456","newProj",Integer.toString(789)});
        
        //Create file system using specific name
        FileOutputStream out = new FileOutputStream(
           new File(spreadSheetLoader.getName() + ".xlsx"));

        //write operation workbook using file out object
        spreadSheetLoader.CloseWorkBook(out);
        out.close();
        System.out.println(spreadSheetLoader.getName() + " written successfully");
   }
    
}
