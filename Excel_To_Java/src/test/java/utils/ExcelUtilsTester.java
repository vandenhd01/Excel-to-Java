package utils;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;



public class ExcelUtilsTester {
	
	 
	

	public static void main(String[] args) throws IOException {
		String excelPath ="./data/TQQQ_Data.xlsx";
		String sheetName = "Sheet1";
		ExcelUtils excelFile= new ExcelUtils(excelPath,sheetName);
		//excelFile.getRowCount();
		//row then column
		//System.out.println("Ticker="+Stocks.ticker);
		//excelFile.getCellData(1, 5);
		//starts col=5 ends col=9
		List<Stocks> stocks= new ArrayList<Stocks>();
		 //String name=excelFile.getCellData(1, 5);
		//System.out.println("Name is " +name);
		
		for(int row=1; row<101;row++) {
			Stocks s=new Stocks();
			for(int col=0;col<11;col++)
			{
				
			if(col==5) {
				//System.out.println("if statement works");
				
			
			s.setTicker(excelFile.getCellDataS(row, col));
			//System.out.println("Ticker is: "+s.getTicker());
			
			}
			if(col==6) {
				//System.out.println("if statement works");
				
			//System.out.println(excelFile.getCellData(1, col));
			s.setOpen(excelFile.getCellData(row, col));
			//System.out.println("open is: "+s.getOpen());
			
			}
			if(col==7) {
				//System.out.println("if statement works");
				
			
			s.setLow(excelFile.getCellData(row, col));
			//System.out.println("low is: "+s.getLow());
			
			
				}
			if(col==8) {
				//System.out.println("if statement works");
				
			
			s.setHigh(excelFile.getCellData(row, col));
			//System.out.println("high is: "+s.getHigh());
			
			
				}
			if(col==4) {
				//System.out.println("if statement works");
				
			
			s.setPChange(excelFile.getCellDataN(row, col));
			//System.out.println("change is: "+s.getPChange());
			
			
				}
			if(col==10) {
				//System.out.println("if statement works");
				
			
			s.setPrice(excelFile.getCellData(row, col));
			//System.out.println("price is: "+s.getPrice());
			
			
				}
			}
			stocks.add(s);
		}
		System.out.println("");
		for(int cntr=0;cntr<8;cntr++) {
		System.out.println(" "+stocks.get(cntr).getTicker()+" "+stocks.get(cntr).getPrice()+" "+stocks.get(cntr).getPChange()*100);
		
		/*
		Stocks s=new Stocks();
		//System.out.println(excelFile.getCellData(1, 5));
		s.setTicker(excelFile.getCellData(1, 5));
		System.out.println(s.getTicker());
		*/
		
		}
	}

}
