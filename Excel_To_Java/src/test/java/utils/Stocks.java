package utils;

import org.apache.poi.ss.usermodel.RichTextString;

public class Stocks {
	private RichTextString ticker;
	private Object open;
	private Object low;
	private Object high;
	private Double pChange;
	private Object price;
	
//Set methods	
	public void setTicker (RichTextString richTextString){
		
		 ticker= richTextString;
		
		}

	public void setOpen (Object object){
		open= object;
		
		}

	public void setLow (Object object){
	
		low= object;
		
	}

	public void setHigh (Object object){
	
		high= object;
		
	}

	public void setPChange (Double numeric){
	
		pChange= numeric;
		
	}
	public void setPrice(Object object) {
		price=object;
	}

//Return methods
	public Object getTicker (){
		return ticker;
	}


	public Object getOpen (){
		return open;
	}

	public Object getLow (){
		return low;
	}

	public Object getHigh (){
		return high;
	}

	public Double getPChange (){
		return pChange;
	}
	public Object getPrice (){
		return price;
	}


	
	

}
