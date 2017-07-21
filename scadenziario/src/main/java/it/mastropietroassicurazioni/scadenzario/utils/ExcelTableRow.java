package it.mastropietroassicurazioni.scadenzario.utils;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class ExcelTableRow {

	Map<String, XSSFCell> rowData;
	
	public ExcelTableRow(XSSFRow row, String[] columns, int startColumn) {
		rowData = new HashMap<String, XSSFCell>();
		for(int i=0; i<columns.length; i++){
			XSSFCell cell = row.getCell(startColumn + i);
			rowData.put(columns[i], cell);
		}
	}
	
	public String getValue(String columnName){
		if(rowData.get(columnName) != null){
			rowData.get(columnName).setCellType(CellType.STRING);
			return rowData.get(columnName).getStringCellValue();
		}else{
			return null;
		}
	}
	
	public Date getValueAsDate(String columnName){
		Date dateValue = null;
		try{
			dateValue = rowData.get(columnName).getDateCellValue();
		}catch(Exception e){
			//
		}
		return dateValue;
	}
	
	public Double getValueAsNumber(String columnName){
		Double numericValue = null;
		try{
			numericValue = rowData.get(columnName).getNumericCellValue();
		}catch(Exception e){
			try{
				String unparsedNumericValue = rowData.get(columnName).getStringCellValue();
				unparsedNumericValue = unparsedNumericValue.replaceAll(",", ".");
				numericValue = Double.parseDouble(unparsedNumericValue);
			}catch(Exception e2){
				numericValue = 0d;
			}
		}
		return numericValue;
	}
	
	
}
