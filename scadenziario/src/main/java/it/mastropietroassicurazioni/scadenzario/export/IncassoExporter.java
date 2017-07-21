package it.mastropietroassicurazioni.scadenzario.export;

import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import it.mastropietroassicurazioni.scadenzario.config.ConfigBean;
import it.mastropietroassicurazioni.scadenzario.utils.ExcelTableIterator;
import it.mastropietroassicurazioni.scadenzario.utils.ExcelTableRow;

public class IncassoExporter extends AbstractExporter {

	private File dataWorkbookFile;
	private String indirizzoFileTemplateIncassi;
	private ConfigBean configBean;
	
	public IncassoExporter(String indirizzoFileDati, String indirizzoFileTemplateIncassi, ConfigBean configBean){
		this.dataWorkbookFile = new File(indirizzoFileDati);
		this.indirizzoFileTemplateIncassi = indirizzoFileTemplateIncassi;
		this.configBean = configBean;
	}
	
	public void export(Date dataIncasso){
		XSSFWorkbook dataWorkbook = null;
		XSSFWorkbook targetWorkbook = null;
		try{
			DateFormat fileNameDateFormat = new SimpleDateFormat("yyyyMMdd");
			File fileTemplateIncassi = new File(indirizzoFileTemplateIncassi);
			Files.copy(Paths.get(indirizzoFileTemplateIncassi), 
					   Paths.get(fileTemplateIncassi.getParentFile().getAbsolutePath() + File.separator + "incassi_tmp.xlsx"), 
					   StandardCopyOption.REPLACE_EXISTING);
			
			dataWorkbook = new XSSFWorkbook(dataWorkbookFile);
			File targetWorkbookFile = new File(fileTemplateIncassi.getParentFile().getAbsolutePath() + File.separator + "incassi_tmp.xlsx");
			targetWorkbook = new XSSFWorkbook(targetWorkbookFile);
			
			XSSFSheet targetSheet = targetWorkbook.getSheetAt(0);
			
			int START_OFFSET = 5;
			int exportedRows = 0;
			
			ExcelTableIterator excelTableIteratorScadenze = new ExcelTableIterator(dataWorkbook, configBean.getSheetPolizze(), configBean.getTabellaPolizze());
			while(excelTableIteratorScadenze.hasNext()){
				ExcelTableRow row = excelTableIteratorScadenze.next();
				if(row.getValueAsDate("Data ultimo pagamento") != null && row.getValueAsDate("Data ultimo pagamento").equals(dataIncasso)){
					XSSFRow newRow = copyRow(targetWorkbook, targetSheet, START_OFFSET + exportedRows, START_OFFSET + exportedRows + 1);
					newRow.setHeight((short)30);
					newRow.getCell(0).setCellValue(row.getValue("Contraente"));
					newRow.getCell(1).setCellValue(row.getValue("Compagnia"));
					newRow.getCell(2).setCellValue(row.getValue("Polizza"));
					if(row.getValueAsDate("Data ultimo pagamento") != null){
						newRow.getCell(3).setCellValue(row.getValueAsDate("Data ultimo pagamento"));
					}else{
						newRow.getCell(3).setCellValue("-");
					}
					if(row.getValueAsDate("Data scadenza") != null){
						newRow.getCell(4).setCellValue(row.getValueAsDate("Data scadenza"));
					}else{
						newRow.getCell(4).setCellValue("-");
					}
					newRow.getCell(5).setCellValue(row.getValue("Targa"));
					if(row.getValueAsDate("Importo rata") != null){
						newRow.getCell(6).setCellValue(row.getValueAsNumber("Importo rata"));
					}else{
						newRow.getCell(6).setCellValue("-");
					}
					if(row.getValueAsDate("Incassato anno") != null){
						newRow.getCell(7).setCellValue(row.getValueAsNumber("Incassato anno"));
					}else{
						newRow.getCell(7).setCellValue("-");
					}
					if(row.getValueAsDate("Importo annuale") != null){
						newRow.getCell(8).setCellValue(row.getValueAsNumber("Importo annuale"));
					}else{
						newRow.getCell(8).setCellValue("-");
					}
					newRow.getCell(9).setCellValue(row.getValue("Ramo"));
					newRow.getCell(10).setCellValue(row.getValue("Tipo pagamento"));
					newRow.getCell(11).setCellValue(row.getValue("OI/ON"));
					newRow.getCell(12).setCellValue(row.getValue("Durata"));
					newRow.getCell(13).setCellValue(row.getValue("I"));
					newRow.getCell(14).setCellValue(row.getValue("II"));
					newRow.getCell(15).setCellValue(row.getValue("III"));
					newRow.getCell(16).setCellValue(row.getValue("IV"));
					newRow.getCell(17).setCellValue(row.getValue("Note polizza"));
					
					exportedRows++;
				}
			}
			removeRow(targetSheet,START_OFFSET);
			
			if(exportedRows > 1){
				targetSheet.getTables().get(0).setCellReferences(new AreaReference(new CellReference(START_OFFSET - 1, 0), 
																				   new CellReference(START_OFFSET + exportedRows - 1, 18)));
			}
			File file = new File(fileTemplateIncassi.getParentFile().getAbsolutePath() + File.separator + "incassi_"+ fileNameDateFormat.format(dataIncasso) +".xlsx");
			FileOutputStream fileOutputStream = new FileOutputStream(file);
			targetWorkbook.write(fileOutputStream);
			fileOutputStream.flush();
			fileOutputStream.close();
			Desktop dt = Desktop.getDesktop();
			dt.open(file);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			try {
				targetWorkbook.close();
				dataWorkbook.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

}
