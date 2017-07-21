package it.mastropietroassicurazioni.scadenzario.config;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Properties;

public class ConfigBean {

	Properties properties = null;
	private String fileLocation;
	private String incassiTemplateLocation;
	private String versamentiTemplateLocation;
	private String sheetAnagrafica;
	private String tabellaAnagrafica;
	private String sheetPolizze;
	private String tabellaPolizze;
	private String sheetPromemoria;
	private String tabellaPromemoria;
	private String sheetVersamenti;
	private String tabellaVersamenti;
	
	private String smtpServer;
	private int smtpPort;
	private String smtpUsername;
	private String smtpPassword;
	private String smtpMailFrom;
	private String smtpMailCc;
	
	private String smsUsername;
	private String smsPassword;
	private String smsNumberFrom;
	
	private int giorniPromemoriaScadenze;
	private int giorniPromemoriaVersamenti;
	
	public ConfigBean() {
		File propertiesFile = new File("scadenzario.properties");
		FileInputStream propertiesInputStream = null;
		properties = new Properties();
		try {
			if(!propertiesFile.exists()){
				propertiesFile.createNewFile();
			}
			propertiesInputStream = new FileInputStream(propertiesFile);
			properties.load(propertiesInputStream);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally {
			try {
				propertiesInputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		
		this.fileLocation = properties.getProperty("workbook.location");
		this.incassiTemplateLocation = properties.getProperty("template_incassi.location");
		this.versamentiTemplateLocation = properties.getProperty("template_versamenti.location");
		this.sheetAnagrafica = properties.getProperty("sheets.anagrafica");
		this.sheetPolizze = properties.getProperty("sheets.polizze");
		this.sheetPromemoria = properties.getProperty("sheets.promemoria");
		this.sheetVersamenti = properties.getProperty("sheets.versamenti");
		this.tabellaAnagrafica = properties.getProperty("tabelle.anagrafica");
		this.tabellaPolizze = properties.getProperty("tabelle.polizze");
		this.tabellaPromemoria = properties.getProperty("tabelle.promemoria");
		this.tabellaVersamenti = properties.getProperty("tabelle.versamenti");
		this.smtpServer = properties.getProperty("smtp.host");
		this.smtpPort = properties.getProperty("smtp.port") == null? 587: Integer.parseInt(properties.getProperty("smtp.port"));
		this.smtpUsername = properties.getProperty("smtp.username");
		this.smtpPassword = properties.getProperty("smtp.password");
		this.smtpMailFrom = properties.getProperty("smtp.mailfrom");
		this.smtpMailCc = properties.getProperty("smtp.mailcc");
		this.smsUsername = properties.getProperty("sms.username");
		this.smsPassword = properties.getProperty("sms.password");
		this.smsNumberFrom = properties.getProperty("sms.from");
		this.giorniPromemoriaScadenze = properties.getProperty("promemoria.giorni_per_scadenza") == null?  0 : Integer.parseInt(properties.getProperty("promemoria.giorni_per_scadenza"));
		this.giorniPromemoriaVersamenti = properties.getProperty("promemoria.giorni_per_versamenti") == null? 0 : Integer.parseInt(properties.getProperty("promemoria.giorni_per_versamenti"));
	}
	
	public int getGiorniPromemoriaScadenze() {
		return giorniPromemoriaScadenze;
	}
	
	public void setGiorniPromemoriaScadenze(int giorniPromemoriaScadenze) {
		this.giorniPromemoriaScadenze = giorniPromemoriaScadenze;
	}

	public int getGiorniPromemoriaVersamenti() {
		return giorniPromemoriaVersamenti;
	}
	
	public void setGiorniPromemoriaVersamenti(int giorniPromemoriaVersamenti) {
		this.giorniPromemoriaVersamenti = giorniPromemoriaVersamenti;
	}
	
	public String getFileLocation() {
		return fileLocation;
	}
	
	public void setFileLocation(String fileLocation) {
		this.fileLocation = fileLocation;
	}
	
	public String getIncassiTemplateLocation() {
		return incassiTemplateLocation;
	}
	
	public void setIncassiTemplateLocation(String incassiTemplateLocation) {
		this.incassiTemplateLocation = incassiTemplateLocation;
	}
	
	public String getVersamentiTemplateLocation() {
		return versamentiTemplateLocation;
	}
	
	public void setVersamentiTemplateLocation(String versamentiTemplateLocation) {
		this.versamentiTemplateLocation = versamentiTemplateLocation;
	}
	
	public String getSheetAnagrafica() {
		return sheetAnagrafica;
	}

	public void setSheetAnagrafica(String sheetAnagrafica) {
		this.sheetAnagrafica = sheetAnagrafica;
	}

	public String getTabellaAnagrafica() {
		return tabellaAnagrafica;
	}

	public void setTabellaAnagrafica(String tabellaAnagrafica) {
		this.tabellaAnagrafica = tabellaAnagrafica;
	}

	public String getSheetPolizze() {
		return sheetPolizze;
	}

	public void setSheetPolizze(String sheetPolizze) {
		this.sheetPolizze = sheetPolizze;
	}

	public String getTabellaPolizze() {
		return tabellaPolizze;
	}

	public void setTabellaPolizze(String tabellaPolizze) {
		this.tabellaPolizze = tabellaPolizze;
	}

	public String getSheetPromemoria() {
		return sheetPromemoria;
	}

	public void setSheetPromemoria(String sheetPromemoria) {
		this.sheetPromemoria = sheetPromemoria;
	}

	public String getTabellaPromemoria() {
		return tabellaPromemoria;
	}

	public void setTabellaPromemoria(String tabellaPromemoria) {
		this.tabellaPromemoria = tabellaPromemoria;
	}
	
	public String getSheetVersamenti() {
		return sheetVersamenti;
	}
	
	public void setSheetVersamenti(String sheetVersamenti) {
		this.sheetVersamenti = sheetVersamenti;
	}
	
	public String getTabellaVersamenti() {
		return tabellaVersamenti;
	}
	
	public void setTabellaVersamenti(String tabellaVersamenti) {
		this.tabellaVersamenti = tabellaVersamenti;
	}

	public String getSmtpServer() {
		return smtpServer;
	}

	public void setSmtpServer(String smtpServer) {
		this.smtpServer = smtpServer;
	}
	
	public int getSmtpPort() {
		return smtpPort;
	}
	
	public void setSmtpPort(int smtpPort) {
		this.smtpPort = smtpPort;
	}

	public String getSmtpUsername() {
		return smtpUsername;
	}

	public void setSmtpUsername(String smtpUsername) {
		this.smtpUsername = smtpUsername;
	}

	public String getSmtpPassword() {
		return smtpPassword;
	}

	public void setSmtpPassword(String smtpPassword) {
		this.smtpPassword = smtpPassword;
	}

	public String getSmtpMailFrom() {
		return smtpMailFrom;
	}

	public void setSmtpMailFrom(String smtpMailFrom) {
		this.smtpMailFrom = smtpMailFrom;
	}

	public String getSmtpMailCc() {
		return smtpMailCc;
	}

	public void setSmtpMailCc(String smtpMailCc) {
		this.smtpMailCc = smtpMailCc;
	}
	
	public String getSmsUsername() {
		return smsUsername;
	}
	
	public void setSmsUsername(String smsUsername) {
		this.smsUsername = smsUsername;
	}
	
	public String getSmsPassword() {
		return smsPassword;
	}
	
	public void setSmsPassword(String smsPassword) {
		this.smsPassword = smsPassword;
	}
	
	public String getSmsNumberFrom() {
		return smsNumberFrom;
	}
	
	public void setSmsNumberFrom(String smsNumberFrom) {
		this.smsNumberFrom = smsNumberFrom;
	}
	
	public void updateConfig(){
		properties.setProperty("workbook.location", this.fileLocation);
		properties.setProperty("template_incassi.location", this.incassiTemplateLocation);
		properties.setProperty("template_versamenti.location", this.versamentiTemplateLocation);
		properties.setProperty("sheets.anagrafica", this.sheetAnagrafica);
		properties.setProperty("sheets.polizze", this.sheetPolizze);
		properties.setProperty("sheets.promemoria", this.sheetPromemoria);
		properties.setProperty("sheets.versamenti", this.sheetVersamenti);
		properties.setProperty("tabelle.anagrafica", this.tabellaAnagrafica);
		properties.setProperty("tabelle.polizze", this.tabellaPolizze);
		properties.setProperty("tabelle.promemoria", this.tabellaPromemoria);
		properties.setProperty("tabelle.versamenti", this.tabellaVersamenti);
		properties.setProperty("smtp.host", this.smtpServer);
		properties.setProperty("smtp.port", Integer.toString(this.smtpPort));
		properties.setProperty("smtp.username", this.smtpUsername);
		properties.setProperty("smtp.password", this.smtpPassword);
		properties.setProperty("smtp.mailfrom", this.smtpMailFrom);
		properties.setProperty("smtp.mailcc", this.smtpMailCc);
		properties.setProperty("sms.username", this.smsUsername);
		properties.setProperty("sms.password", this.smsPassword);
		properties.setProperty("sms.from", this.smsNumberFrom);
		properties.setProperty("promemoria.giorni_per_scadenza", Integer.toString(this.giorniPromemoriaScadenze));
		properties.setProperty("promemoria.giorni_per_versamenti", Integer.toString(this.giorniPromemoriaVersamenti));
		
		File propertiesFile = new File("scadenzario.properties");
		FileOutputStream propertiesOutputStream = null;
		try {
			propertiesOutputStream = new FileOutputStream(propertiesFile);
			properties.store(propertiesOutputStream, "");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally {
			try {
				propertiesOutputStream.flush();
				propertiesOutputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
	
}
