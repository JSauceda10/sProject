/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package exceltesting;

import com.mysql.jdbc.CallableStatement;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;
/**
 *
 * @author JOSEPH
 */
public class SpreadSheetOperations {
    String pId = null;
    String pName = null;
    double rev = 0;
    SpreadSheetLoader spreadSheetLoader;
    public SpreadSheetOperations(SpreadSheetLoader spreadSheetLoader) {
        this.spreadSheetLoader = spreadSheetLoader;
    }
    
    public void FetchAllRecords(){
        try {
            Class.forName("com.mysql.jdbc.Driver");
            Connection conn = DriverManager.getConnection("jdbc:mysql://localhost/proj?autoReconnect=true&useSSL=false","root","0000");
        
        Statement st = conn.createStatement();
        ResultSet myRs = st.executeQuery("Select * from Projects;");
        int count =1;
        while(myRs.next()){
            count++;
            pId = myRs.getString("projectId");
            pName = myRs.getString("projectsName");
            rev = myRs.getDouble("revenue");
            System.out.println(pId);      
            System.out.println(pName);      
            System.out.println(rev);
            
            spreadSheetLoader.putNewRowIntoSpreadSheet(new Object[]{pId,pName,Double.toString(rev)});
                        
        }
        } catch (ClassNotFoundException ex) {
            Logger.getLogger(SpreadSheetOperations.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SQLException ex) {
            Logger.getLogger(SpreadSheetOperations.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        
    }
    
    public void UpdateRecord(int pos, Object[] obj){
        try {
            Class.forName("com.mysql.jdbc.Driver");
            Connection conn = DriverManager.getConnection("jdbc:mysql://localhost/proj?autoReconnect=true&useSSL=false","root","0000");
            
            CallableStatement cs = (CallableStatement) conn.prepareCall("{call UpdateProjectsTable(?,?,?)}");
            
            cs.setString(1, pId);
            cs.setString(2, pName);
            cs.setDouble(3, rev);
            
            cs.execute();
            
            cs.close();
            conn.close();
            
            Object [] objects = spreadSheetLoader.getWorksheetInfo().get(Integer.toString(pos));
            
            XSSFSheet sheet = spreadSheetLoader.getWorkbook().getSheetAt(0);
            XSSFRow row = sheet.getRow(pos);
            int count = 0;
            for (Object object : obj){
                
                XSSFCell cell = row.getCell(count);
                cell.setCellValue((String)object);
                if(objects != null)
                    objects[count] = (String)object;
                else
                    System.out.println("objects is null!!!");
                count++;
        }
            
            
        } catch (ClassNotFoundException ex) {
            Logger.getLogger(SpreadSheetOperations.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SQLException ex) {
            Logger.getLogger(SpreadSheetOperations.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
}
