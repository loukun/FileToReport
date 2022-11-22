package entity;

import java.util.List;

public class ExcelSheet {

	private String sheetName;
	
	private List<RowObj> rowObj;

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public List<RowObj> getRowObj() {
		return rowObj;
	}

	public void setRowObj(List<RowObj> rowObj) {
		this.rowObj = rowObj;
	}
	
	
}
