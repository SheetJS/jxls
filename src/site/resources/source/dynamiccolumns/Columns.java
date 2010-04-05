package net.sf.jxls.sample.dynamicColumns;


public class Columns {
	

	private String name;
	private String rate;
	
	public Columns(String name, String rate){
		this.name= name;
		this.rate= rate;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getRate() {
		return rate;
	}

	public void setRate(String rate) {
		this.rate = rate;
	}
	
}
