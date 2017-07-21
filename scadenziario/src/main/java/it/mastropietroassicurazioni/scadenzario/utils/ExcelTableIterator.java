package it.mastropietroassicurazioni.scadenzario.utils;

import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelTableIterator implements Iterator<ExcelTableRow> {
	
	private int startRow = 0;
	private int endRow = 0;
	private int startColumn = 0;
	private int currentRow = 0;
	private String[] columnNames;
	private XSSFSheet sheet;

	public ExcelTableIterator(XSSFWorkbook workbook, String sheetName, String tableName) {
		sheet = workbook.getSheet(sheetName);
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
                    }
                }   
	    	}
	    }
	            

	}
	
	public boolean hasNext() {
		return currentRow < endRow;
	}

	public ExcelTableRow next() {
		currentRow++;
		return new ExcelTableRow(sheet.getRow(currentRow), columnNames, startColumn);
	}

}
