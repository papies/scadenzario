package it.mastropietroassicurazioni.scadenzario.sms;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.text.StrSubstitutor;
import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.HttpVersion;
import org.apache.http.NameValuePair;
import org.apache.http.client.entity.UrlEncodedFormEntity;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.message.BasicNameValuePair;
import org.apache.http.params.BasicHttpParams;
import org.apache.http.params.HttpConnectionParams;
import org.apache.http.params.HttpParams;
import org.apache.http.params.HttpProtocolParamBean;
import org.apache.http.util.EntityUtils;

import it.mastropietroassicurazioni.scadenzario.config.ConfigBean;
 
public class SmsSender {
	
	//private static String USERNAME = "oscarmastropietro";
	//private static String PASSWORD = "oscar2017";
	//private static String SMS_FROM_NUMBER = "393317762704";
	private static String SMS_TYPE = "send_sms_classic";
	private static String SMS_CHECK_CREDIT = "get_credit";
	private static String SMS_FROM_NAME = null;
	
//	public static void main(String[] args) {
//		SmsSender sender = new SmsSender();
//		Map<String, Object> parameters = new HashMap<String, Object>();
//		parameters.put("Contraente", "Marco Papetti");
//		try {
//			sender.send("compleanno", "393317762704", parameters, new ConfigBean());
//		} catch (Exception e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//	}
	
	public int checkSmsLeft(ConfigBean configBean) throws IOException{
		String result = skebbyGatewayCheckCreditSMS(configBean.getSmsUsername(), configBean.getSmsPassword(), "UTF-8");
		int smsLeft = 0;
		String[] components = result.split("&");
		for(String component: components){
			if(component.split("=").length > 1 &&
			   "classic_sms".equals(component.split("=")[0])){
				smsLeft = Integer.parseInt(component.split("=")[1]);
				break;
			}
		}
		return smsLeft;
	}
 
    public boolean send(String tipoSms, String destinatario, Map<String, Object> parameters, ConfigBean configBean) throws IOException {
 
        // SMS CLASSIC dispatch with custom alphanumeric sender
        String result = skebbyGatewaySendSMS(configBean.getSmsUsername(), configBean.getSmsPassword(), new String[]{destinatario}, buildSmsText(tipoSms, parameters), SMS_TYPE, configBean.getSmsNumberFrom(), SMS_FROM_NAME);
         
        // SMS Basic dispatch
        // String result = skebbyGatewaySendSMS(username, password, recipients, "Hi Mike, how are you?", "send_sms_basic", null, null);
          
        // SMS CLASSIC dispatch with custom numeric sender
        // String result = skebbyGatewaySendSMS(username, password, recipients, "Hi Mike, how are you?", "send_sms_classic", "393471234567", null);
          
        // SMS CLASSIC PLUS dispatch (with delivery report) with custom alphanumeric sender
        // String result = skebbyGatewaySendSMS(username, password, recipients, "Hi Mike, how are you?", "send_sms_classic_report", null, "John");
 
        // SMS CLASSIC PLUS dispatch (with delivery report) with custom numeric sender
        // String result = skebbyGatewaySendSMS(username, password, recipients, "Hi Mike, how are you?", "send_sms_classic_report", "393471234567", null);
 
        // ------------------------------------------------------------------
        // Check the complete documentation at http://www.skebby.com/business/index/send-docs/
        // ------------------------------------------------------------------
        // For eventual errors see http:#www.skebby.com/business/index/send-docs/#errorCodesSection
        // WARNING: in case of error DON'T retry the sending, since they are blocking errors
        // ------------------------------------------------------------------   
        //System.out.println("result: "+result);
        return (result.split("&").length > 0 &&
        		result.split("&")[0].split("=").length > 1 &&
        		"success".equals(result.split("&")[0].split("=")[1]));
    }
    
    private String buildSmsText(String tipoSms, Map<String, Object> parameters) throws IOException{
    	 String templateString = readFile("sms_"+ tipoSms +".template");
    	 StrSubstitutor sub = new StrSubstitutor(parameters);
    	 String resolvedString = sub.replace(templateString);
    	 return resolvedString;
    }
     
    protected String skebbyGatewaySendSMS(String username, String password, String [] recipients, String text, String smsType, String senderNumber, String senderString) throws IOException{
        return skebbyGatewaySendSMS(username, password, recipients, text, smsType,  senderNumber,  senderString, "UTF-8");
    }
     
    protected String skebbyGatewaySendSMS(String username, String password, String [] recipients, String text, String smsType, String senderNumber, String senderString, String charset) throws IOException{
         
        if (!charset.equals("UTF-8") && !charset.equals("ISO-8859-1")) {
         
            throw new IllegalArgumentException("Charset not supported.");
        }
         
        String endpoint = "http://gateway.skebby.it/api/send/smseasy/advanced/http.php";
        HttpParams params = new BasicHttpParams();
        HttpConnectionParams.setConnectionTimeout(params, 10*1000);
        DefaultHttpClient httpclient = new DefaultHttpClient(params);
        HttpProtocolParamBean paramsBean = new HttpProtocolParamBean(params);
        paramsBean.setVersion(HttpVersion.HTTP_1_1);
        paramsBean.setContentCharset(charset);
        paramsBean.setHttpElementCharset(charset);
         
        List<NameValuePair> formparams = new ArrayList<NameValuePair>();
        formparams.add(new BasicNameValuePair("method", smsType));
        formparams.add(new BasicNameValuePair("username", username));
        formparams.add(new BasicNameValuePair("password", password));
        if(null != senderNumber)
            formparams.add(new BasicNameValuePair("sender_number", senderNumber));
        if(null != senderString)
            formparams.add(new BasicNameValuePair("sender_string", senderString));
         
        for (String recipient : recipients) {
            formparams.add(new BasicNameValuePair("recipients[]", "39"+ recipient));
        }
        formparams.add(new BasicNameValuePair("text", text));
        formparams.add(new BasicNameValuePair("charset", charset));
 
     
        UrlEncodedFormEntity entity = new UrlEncodedFormEntity(formparams, charset);
        HttpPost post = new HttpPost(endpoint);
        post.setEntity(entity);
         
        HttpResponse response = httpclient.execute(post);
        HttpEntity resultEntity = response.getEntity();
        if(null != resultEntity){
            return EntityUtils.toString(resultEntity);
        }
        return null;
    }
    
    protected String skebbyGatewayCheckCreditSMS(String username, String password, String charset) throws IOException{
        
        if (!charset.equals("UTF-8") && !charset.equals("ISO-8859-1")) {
         
            throw new IllegalArgumentException("Charset not supported.");
        }
         
        String endpoint = "http://gateway.skebby.it/api/send/smseasy/advanced/http.php";
        HttpParams params = new BasicHttpParams();
        HttpConnectionParams.setConnectionTimeout(params, 10*1000);
        DefaultHttpClient httpclient = new DefaultHttpClient(params);
        HttpProtocolParamBean paramsBean = new HttpProtocolParamBean(params);
        paramsBean.setVersion(HttpVersion.HTTP_1_1);
        paramsBean.setContentCharset(charset);
        paramsBean.setHttpElementCharset(charset);
         
        List<NameValuePair> formparams = new ArrayList<NameValuePair>();
        formparams.add(new BasicNameValuePair("method", SMS_CHECK_CREDIT));
        formparams.add(new BasicNameValuePair("username", username));
        formparams.add(new BasicNameValuePair("password", password));
        formparams.add(new BasicNameValuePair("charset", charset));
 
        UrlEncodedFormEntity entity = new UrlEncodedFormEntity(formparams, charset);
        HttpPost post = new HttpPost(endpoint);
        post.setEntity(entity);
         
        HttpResponse response = httpclient.execute(post);
        HttpEntity resultEntity = response.getEntity();
        if(null != resultEntity){
            return EntityUtils.toString(resultEntity);
        }
        return null;
    }
    
    private String readFile(String file) throws IOException {
        BufferedReader reader = new BufferedReader(new FileReader(file));
        String         line = null;
        StringBuilder  stringBuilder = new StringBuilder();
        String         ls = System.getProperty("line.separator");

        try {
            while((line = reader.readLine()) != null) {
            	if(stringBuilder.length() > 0){
            		stringBuilder.append(ls);
            	}
                stringBuilder.append(line);
            }

            return stringBuilder.toString();
        } finally {
            reader.close();
        }
    }
    
}
       