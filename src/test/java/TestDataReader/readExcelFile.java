package TestDataReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class readExcelFile {
	public ArrayList<HashMap> readExcel(String filePath,String fileName,String sheetName) throws IOException{
	    File file =    new File(filePath+"//"+fileName);
	    @SuppressWarnings("rawtypes")
		ArrayList<HashMap> excelDataArray = new ArrayList<HashMap>();
	    FileInputStream inputStream = new FileInputStream(file);
	    Workbook RWorkbook = null;
	    String fileExtensionName = fileName.substring(fileName.indexOf("."));
	    
	    /*if(fileExtensionName.equals(".xlsx")){
	    guru99Workbook = new XSSFWorkbook(inputStream);
	    }
	    //Check condition if the file is xls file
	    else*/ 
	    if(fileExtensionName.equals(".xls")){
	        //If it is xls file then create object of HSSFWorkbook class
	        RWorkbook = new HSSFWorkbook(inputStream);
	    }
	    //Read sheet inside the workbook by its name
	    Sheet guru99Sheet = RWorkbook.getSheet(sheetName);
	    Iterator<Row> iterateRow = guru99Sheet.iterator();
	    while (iterateRow.hasNext())
	    {
	    	HashMap<String, String> exceldatahashmap = new HashMap<String, String>();
	    	Row Rrow = iterateRow.next();
	        if(Rrow.getCell(1).getStringCellValue().equals("Yes")){
	        	String cellvalue="";
	        	//Create a loop to print cell values in a row
	        	
	        	for (int j = 0; j < Rrow.getLastCellNum(); j++) {
	            	String celltype = String.valueOf(Rrow.getCell(j).getCellType());
	            	//if(celltype.equals("STRING"))
	            	if(celltype.equals("1"))
	            	{
	                System.out.print(Rrow.getCell(j).getStringCellValue()+"|| ");
	                String heading = guru99Sheet.getRow(0).getCell(j).getStringCellValue();
	                cellvalue = Rrow.getCell(j).getStringCellValue();
	                exceldatahashmap.put(heading,cellvalue);
	                }else{
	                System.out.print(Rrow.getCell(j).getNumericCellValue()+"|| ");
	                
	                }
	            	
	            		
	        	}
	        	excelDataArray.add(exceldatahashmap);
	        	
	        }else{
	        	//catch exception
	        }
	      }
	    
	    return excelDataArray;  
	        
	    } 

}
