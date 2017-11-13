package ReadExcellData;

import java.io.File;
import java.io.FileInputStream;



import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcell {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
           File scr = new File ("C:\\Users\\Ruslan Omurbekov\\My Drivers\\Selenium\\Excell sheet\\Test data.xlsx");
            
			FileInputStream fis= new FileInputStream(scr);
	
	        XSSFWorkbook wb=new XSSFWorkbook (fis);
	        wb.getSheetAt(0);
	        XSSFSheet sheet1 = wb.getSheetAt(0);
//	        String data0=sheet1.getRow(0).getCell(0).getStringCellValue();
//	        System.out.println("Data from Excell is " + data0);
//	        String data1=sheet1.getRow(0).getCell(1).getStringCellValue();
//	        System.out.println("Data from Excell is " + data1);
//	        String data2=sheet1.getRow(1).getCell(0).getStringCellValue();
//	        System.out.println("Data from Excell is " + data2);
//	        String data3=sheet1.getRow(1).getCell(1).getStringCellValue();
//	        System.out.println("Data from Excell is " + data3);
	       int rowcount= sheet1.getLastRowNum();
	       System.out.println("Total rows is" + rowcount);
	       for (int i=0; i<rowcount; i++){
	    	   String data0=sheet1.getRow(i).getCell(0).getStringCellValue();
	    	   System.out.println("Test Data Excell is"  + data0);
	    	   
	       
	      
	    	   
	    	   
	       }
	       
	        wb.close();
	
	
	
	
	}

}
