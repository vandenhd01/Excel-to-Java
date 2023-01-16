
package utils;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf. usermodel.XSSFWorkbook;

public class ExcelUtils {
	
	
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	
	public ExcelUtils(String excelPath, String sheetName) {
	try {
	workbook = new XSSFWorkbook(excelPath);
	sheet = workbook.getSheet(sheetName);
	}catch(Exception exp) {
	System.out.println(exp.getCause());
	System.out.println(exp.getMessage());
	exp.printStackTrace();
		}
	}
	
	
	
	/*    
	public static void getCellData(int rowNum, int colNum) throws IOException{
	DataFormatter formatter = new DataFormatter();
    //Object value = formatter.formatCellValue(sheet.getRow(rowNum).getCell(colNum));
	//System.out.println(value);
	CellAddress cellAddress = new CellAddress("F2");
	//Row row = sheet.getRow(cellAddress.getRow());
	Row row = sheet.getRow(rowNum);
	Cell cell = row.getCell(colNum);
	
	if (cell.getCellType() == CellType.FORMULA) {
	    switch (cell.getCachedFormulaResultType()) {
	        case BOOLEAN:
	            System.out.println(cell.getBooleanCellValue());
	            break;
	        case NUMERIC:
	            System.out.println(cell.getNumericCellValue());
	            break;
	        case STRING:
	            System.out.println(cell.getRichStringCellValue());
	            
	            break;
	    	}
		}
	}
	*/
	
	
	@SuppressWarnings("unchecked")
	public static RichTextString getCellDataS(int rowNum, int colNum) throws IOException{
	        DataFormatter formatter = new DataFormatter();
	        
	        Row row = sheet.getRow(rowNum);
	        Cell cell = row.getCell(colNum);

	       
	                    RichTextString string=cell.getRichStringCellValue();
	                    
	                    return string;
	               
	}
	
	@SuppressWarnings("unchecked")
	public static Double getCellDataN(int rowNum, int colNum) throws IOException{
	        DataFormatter formatter = new DataFormatter();
	        
	        Row row = sheet.getRow(rowNum);
	        Cell cell = row.getCell(colNum);

	       
	        Double numeric=cell.getNumericCellValue();
            
            return  numeric; 
	               
	}
	
	@SuppressWarnings("unchecked")
	public static <T> T getCellData(int rowNum, int colNum) throws IOException{
	        DataFormatter formatter = new DataFormatter();
	        //Object value = formatter.formatCellValue(sheet.getRow(rowNum).getCell(colNum));
	        //System.out.println(value);
	        CellAddress cellAddress = new CellAddress("F2");
	        //Row row = sheet.getRow(cellAddress.getRow());
	        Row row = sheet.getRow(rowNum);
	        Cell cell = row.getCell(colNum);

	        if (cell.getCellType() == CellType.FORMULA) {
	            switch (cell.getCachedFormulaResultType()) {
	                case BOOLEAN:
	                    Boolean Bool= cell.getBooleanCellValue();
	                    return (T) Bool;
	                    

	                case NUMERIC:
	                    Double numeric=cell.getNumericCellValue();
	                    ///System.out.println("This is a numeric value");
	                    return (T) numeric; 

	                case STRING:
	                    RichTextString string=cell.getRichStringCellValue();
	                    //System.out.println("This is a string value");
	                    return (T) string;
	                default:
	                	return (T) "";

	                }
	        }
	        return null;
	}
        
	
	
public static void getRowCount() {
	try {
		
		int rowCount = sheet.getPhysicalNumberOfRows();
		System.out.println("No of Rows: "+rowCount);
	}catch(Exception exp) {
		System.out.println(exp.getCause());
		System.out.println(exp.getMessage());
		exp.printStackTrace();
			}

		}
}
