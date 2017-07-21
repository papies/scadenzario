package it.mastropietroassicurazioni.scadenzario.notifications;

import java.util.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.concurrent.ConcurrentMap;

import org.mapdb.DB;
import org.mapdb.DBMaker;

public class NotificationsManager {

	private final static DateFormat DATE_FORMAT = new SimpleDateFormat("yyyyMMdd");
		
	public NotificationsManager() {
		
	}
	
	public NotificationType getNotifiedPolizza(String numeroPolizza, Date dataPagamento){
		DB db = DBMaker.fileDB("notifications.db").make();
		ConcurrentMap polizze = db.hashMap("polizze").createOrOpen();
		String key = numeroPolizza + "_" + DATE_FORMAT.format(dataPagamento);
		NotificationType notificationType = (NotificationType)polizze.get(key);
		db.close();
		return notificationType;
	}
	
	public void putNotifiedPolizza(String numeroPolizza, Date dataPagamento, NotificationType notificationType){
		DB db = DBMaker.fileDB("notifications.db").make();
		ConcurrentMap polizze = db.hashMap("polizze").createOrOpen();
		String key = numeroPolizza + "_" + DATE_FORMAT.format(dataPagamento);
		polizze.put(key, notificationType);
		db.commit();
		db.close();
	}
	
	public NotificationType getNotifiedVersamento(String numeroPolizza, Date dataVersamento){
		DB db = DBMaker.fileDB("notifications.db").make();
		ConcurrentMap versamenti = db.hashMap("versamenti").createOrOpen();
		String key = numeroPolizza + "_" + DATE_FORMAT.format(dataVersamento);
		NotificationType notificationType = (NotificationType)versamenti.get(key);
		db.close();
		return notificationType;
	}
	
	public void putNotifiedVersamento(String numeroPolizza, Date dataVersamento, NotificationType notificationType){
		DB db = DBMaker.fileDB("notifications.db").make();
		ConcurrentMap versamenti = db.hashMap("versamenti").createOrOpen();
		String key = numeroPolizza + "_" + DATE_FORMAT.format(dataVersamento);
		versamenti.put(key, notificationType);
		db.commit();
		db.close();
	}
	
	public NotificationType getNotifiedCompleanno(String nomeCliente, String dataCompleanno){
		DB db = DBMaker.fileDB("notifications.db").make();
		ConcurrentMap compleanni = db.hashMap("compleanni").createOrOpen();
		String key = nomeCliente + "_" + dataCompleanno;
		NotificationType notificationType = (NotificationType)compleanni.get(key);
		db.close();
		return notificationType;
	}
	
	public void putNotifiedCompleanno(String nomeCliente, String dataCompleanno, NotificationType notificationType){
		DB db = DBMaker.fileDB("notifications.db").make();
		ConcurrentMap compleanni = db.hashMap("compleanni").createOrOpen();
		String key = nomeCliente + "_" + dataCompleanno;
		compleanni.put(key, notificationType);
		db.commit();
		db.close();
	}
	
}
