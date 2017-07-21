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

public class VersamentiExporter extends AbstractExporter {

	private File dataWorkbookFile;
	private String indirizzoFileTemplateVersamenti;
	private ConfigBean configBean;
	
	public VersamentiExporter(String indirizzoFileDati, String indirizzoFileTemplateVersamenti, ConfigBean configBean){
		this.dataWorkbookFile = new File(indirizzoFileDati);
		this.indirizzoFileTemplateVersamenti = indirizzoFileTemplateVersamenti;
		this.configBean = configBean;
	}
	
	public void export(Date dataVersamento){
		XSSFWorkbook dataWorkbook = null;
		XSSFWorkbook targetWorkbook = null;
		try{
			DateFormat fileNameDateFormat = new SimpleDateFormat("yyyyMMdd");
			File fileTemplateVersamenti = new File(indirizzoFileTemplateVersamenti);
			Files.copy(Paths.get(indirizzoFileTemplateVersamenti), 
					   Paths.get(fileTemplateVersamenti.getParentFile().getAbsolutePath() + File.separator + "versamenti_tmp.xlsx"), 
					   StandardCopyOption.REPLACE_EXISTING);
			
			dataWorkbook = new XSSFWorkbook(dataWorkbookFile);
			File targetWorkbookFile = new File(fileTemplateVersamenti.getParentFile().getAbsolutePath() + File.separator + "versamenti_tmp.xlsx");
			targetWorkbook = new XSSFWorkbook(targetWorkbookFile);
			
			XSSFSheet targetSheet = targetWorkbook.getSheetAt(0);
			
			int START_OFFSET = 5;
			int exportedRows = 0;
			
			ExcelTableIterator excelTableIteratorScadenze = new ExcelTableIterator(dataWorkbook, configBean.getSheetVersamenti(), configBean.getTabellaVersamenti());
			while(excelTableIteratorScadenze.hasNext()){
				ExcelTableRow row = excelTableIteratorScadenze.next();				
				if(row.getValueAsDate("Data prevista versamento") != null && row.getValueAsDate("Data prevista versamento").equals(dataVersamento)){
					XSSFRow newRow = copyRow(targetWorkbook, targetSheet, START_OFFSET + exportedRows, START_OFFSET + exportedRows + 1);
					newRow.setHeight((short)30);
					newRow.getCell(0).setCellValue(row.getValue("Contraente"));
					newRow.getCell(1).setCellValue(row.getValue("Compagnia"));
					newRow.getCell(2).setCellValue(row.getValue("Polizza"));
					if(row.getValueAsDate("Importo rata") != null){
						newRow.getCell(3).setCellValue(row.getValueAsNumber("Importo rata"));
					}else{
						newRow.getCell(3).setCellValue("-");
					}
					newRow.getCell(4).setCellValue(row.getValue("Tipo pagamento"));
					if(row.getValueAsDate("Data incasso") != null){
						newRow.getCell(5).setCellValue(row.getValueAsDate("Data incasso"));
					}else{
						newRow.getCell(5).setCellValue("-");
					}
					if(row.getValueAsDate("Data prevista versamento") != null){
						newRow.getCell(6).setCellValue(row.getValueAsDate("Data prevista versamento"));
					}else{
						newRow.getCell(6).setCellValue("-");
					}
					
					newRow.getCell(7).setCellValue(row.getValue("Stato versamento"));
					
					exportedRows++;
				}
			}
			removeRow(targetSheet,START_OFFSET);
			
			if(exportedRows > 1){
				targetSheet.getTables().get(0).setCellReferences(new AreaReference(new CellReference(START_OFFSET - 1, 0), 
																				   new CellReference(START_OFFSET + exportedRows - 1, 26)));
			}
			
			File file = new File(fileTemplateVersamenti.getParentFile().getAbsolutePath() + File.separator + "versamenti_"+ fileNameDateFormat.format(dataVersamento) +".xlsx");
			FileOutputStream fileOutputStream = new FileOutputStream(file);
			targetWorkbook.write(fileOutputStream);
			fileOutputStream.flush();
			fileOutputStream.close();
			Desktop dt = Desktop.getDesktop();
			dt.open(file);
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
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
