package it.mastropietroassicurazioni.scadenzario.utils;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelTableWriter{
	
	private int startRow = 0;
	private int endRow = 0;
	private int currentRow = 0;
	private int startColumn = 0;
	private String[] columnNames;
	private Map<String, Integer> columnIndexes;
	private XSSFSheet sheet;

	public ExcelTableWriter(XSSFWorkbook workbook, String sheetName, String tableName) {
		sheet = workbook.getSheet(sheetName);
		this.columnIndexes = new HashMap<String, Integer>();
		List<XSSFTable> tables = sheet.getTables();
	    for(XSSFTable t : tables) {
	    	if(t.getName().equals(tableName)){
	    		startRow = t.getStartCellReference().getRow();
	    		currentRow = startRow;
	            endRow = t.getEndCellReference().getRow();
	            startColumn = t.getStartCellReference().getCol();
	            int endColumn = t.getEndCellReference().getCol();
	            
	            columnNames = new String[endColumn - startColumn + 1];
	            
	            for (int j = startColumn; j <= endColumn; j++) {
                    XSSFCell cell = sheet.getRow(startRow).getCell(j);
                    if (cell != null) {
                        String columnName = cell.getStringCellValue();
                        columnNames[j - startColumn] = columnName;
                        columnIndexes.put(columnName, j);
                    }
                }   
	    	}
	    }
	            

	}
	
	public void writeExcelRow(ExcelTableRow source){
		currentRow++;
		sheet.shiftRows(currentRow, endRow, 1);
		endRow++;
		XSSFRow row = sheet.getRow(currentRow);
		if(row == null){
			row = sheet.createRow(currentRow);
		}
		for(String columnName: this.columnNames){
			int columnIndex = columnIndexes.get(columnName);
			XSSFCell cell = row.getCell(columnIndex);
			if(cell == null){
				cell = row.createCell(columnIndex);
			}
			cell.setCellValue(source.getValue(columnName));
		}
	}

}
