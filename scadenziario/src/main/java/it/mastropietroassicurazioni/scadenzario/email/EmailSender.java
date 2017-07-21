package it.mastropietroassicurazioni.scadenzario.email;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

import org.apache.commons.lang3.text.StrSubstitutor;

import it.mastropietroassicurazioni.scadenzario.config.ConfigBean;
 
public class EmailSender {
	
//	public static void main(String[] args) {
//		EmailSender emailSender = new EmailSender();
//		ConfigBean configBean = new ConfigBean();
//		configBean.setSmtpServer("smtp.gmail.com");
//		configBean.setSmtpPort(587);
//		configBean.setSmtpUsername("papies@gmail.com");
//		configBean.setSmtpPassword("aqwhamiwyzbzjtta");
//		configBean.setSmtpMailFrom("papies@gmail.com");
//		configBean.setSmtpMailCc("papies@email.it");
//		
//		Map<String, Object> parameters = new HashMap<String, Object>();
//		parameters.put("Contraente","Marco Papetti");
//		parameters.put("Ramo","DANNI");
//		parameters.put("Auto/moto","AUTO");
//		parameters.put("Targa","DT873ZD");
//		parameters.put("Polizza","123456789");
//		parameters.put("Data prossima scadenza","30/07/2017");
//		parameters.put("Importo rata","130,00");
//		
//		try {
//			emailSender.send("polizza", "Mail di prova", "oscarmastropietro@gmail.com", parameters, configBean);
//		} catch (Exception e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//		
//	}
	
    public boolean send(String tipoMail, String subject, String destinatario, Map<String, Object> parameters, ConfigBean configBean) throws Exception {
    	String[] subjectAndText = buildEmailText(tipoMail, parameters);
    	return sendHtmlEmail(configBean.getSmtpServer(), 
    				  		 Integer.toString(configBean.getSmtpPort()), 
    				  		 configBean.getSmtpUsername(), 
    				  		 configBean.getSmtpPassword(), 
    				  		 configBean.getSmtpMailFrom(), 
    				  		 destinatario, 
    				  		 configBean.getSmtpMailCc(),
    				  		 subjectAndText[0], 
    				  		 subjectAndText[1]);
        
    }
    
    private String[] buildEmailText(String tipoMail, Map<String, Object> parameters) throws IOException{
    	 String[] templateStrings = readFile("email_"+ tipoMail +".template");
    	 StrSubstitutor sub = new StrSubstitutor(parameters);
    	 String[] resolvedStrings = new String[templateStrings.length];
    	 for(int i=0; i<templateStrings.length; i++){
    		 resolvedStrings[i] = sub.replace(templateStrings[i]);
    	 }
    	 return resolvedStrings;
    }
     
    public boolean sendHtmlEmail(String host, String port,
            final String userName, final String password, String fromAddress, String toAddress, String ccnAddress,
            String subject, String message) throws AddressException,
            MessagingException {
 
        // sets SMTP server properties
        Properties properties = new Properties();
        properties.put("mail.smtp.host", host);
        properties.put("mail.smtp.port", port);
        properties.put("mail.smtp.auth", "true");
        properties.put("mail.smtp.starttls.enable", "true");
 
        // creates a new session with an authenticator
        Authenticator auth = new Authenticator() {
            public PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(userName, password);
            }
        };
 
        Session session = Session.getInstance(properties, auth);
 
        // creates a new e-mail message
        Message msg = new MimeMessage(session);
 
        msg.setFrom(new InternetAddress(fromAddress));
        InternetAddress[] toAddresses = { new InternetAddress(toAddress) };
        msg.setRecipients(Message.RecipientType.TO, toAddresses);
        if(ccnAddress != null){
            InternetAddress[] ccnAddresses = { new InternetAddress(ccnAddress) };
            msg.setRecipients(Message.RecipientType.BCC, ccnAddresses);
        }
        msg.setSubject(subject);
        msg.setSentDate(new Date());
        // set plain text message
        msg.setContent(message, "text/html; charset=UTF-8");
 
        // sends the e-mail
        Transport.send(msg);
        return true;
    }
    
    private String[] readFile(String file) throws IOException {
        BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(file),"UTF-8"));
        String         line = null;
        String subject = null;
        StringBuilder textStringBuilder = new StringBuilder();
        String ls = System.getProperty("line.separator");

        try {
            while((line = reader.readLine()) != null) {
            	if(subject == null){
            		subject = line;
            	}else{
            		if(textStringBuilder.length() > 0){
                    	textStringBuilder.append(ls);
            		}
                	textStringBuilder.append(line);
            	}
            }
            return new String[]{subject, textStringBuilder.toString()};
        } finally {
            reader.close();
        }
    }
    
}
       