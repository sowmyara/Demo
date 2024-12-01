package ExcelReader1;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

 public class XLUtility {
	
	
	
	public String sourceSheet;
	String filepath = null;
	public String getCellValue;
	
	XLUtility(String filepath)
	{
		
		 this.filepath = filepath;
		
	}
	
 

	public int getRowcount(String Sheet1) throws IOException {
		
		
		
		FileInputStream fi = new FileInputStream(filepath);
		XSSFWorkbook S = new XSSFWorkbook(fi);     
        XSSFSheet sourceSheet = S.getSheet(Sheet1);
      
        //creating variable to get rowcount and columncount
        
        int rowCount=sourceSheet.getLastRowNum();
     
        return rowCount;
	}
	
	public int getColumncount(String Sheet1 , int rowNum) throws IOException {
		
		
		FileInputStream fi = new FileInputStream(filepath);
		XSSFWorkbook S = new XSSFWorkbook(fi);     
        XSSFSheet sourceSheet = S.getSheet(Sheet1);
        int columnCount = sourceSheet.getRow(rowNum).getLastCellNum();
        return columnCount;
	}

	public String  getCellValue( String Sheet1, int rowNum, int colnum ) throws IOException 
	{
		FileInputStream fi = new FileInputStream(filepath);
		XSSFWorkbook S = new XSSFWorkbook(fi);     
        XSSFSheet sourceSheet = S.getSheet(Sheet1);
        XSSFRow row= sourceSheet.getRow(rowNum);
        XSSFCell cell = row.getCell(colnum);
        
        DataFormatter formatter = new DataFormatter();
        String data;
        data= formatter.formatCellValue(cell);
        
        return getCellValue;
        
		
		
		
	}
	
}
