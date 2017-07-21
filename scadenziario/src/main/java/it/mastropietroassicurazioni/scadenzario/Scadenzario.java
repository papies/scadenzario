package it.mastropietroassicurazioni.scadenzario;

import java.awt.BorderLayout;
import java.awt.Component;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.ScrollPaneConstants;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumn;
import javax.swing.table.TableColumnModel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.jgoodies.forms.layout.ColumnSpec;
import com.jgoodies.forms.layout.FormLayout;
import com.jgoodies.forms.layout.FormSpecs;
import com.jgoodies.forms.layout.RowSpec;

import it.mastropietroassicurazioni.scadenzario.config.ConfigBean;
import it.mastropietroassicurazioni.scadenzario.email.EmailSender;
import it.mastropietroassicurazioni.scadenzario.export.IncassoExporter;
import it.mastropietroassicurazioni.scadenzario.export.VersamentiExporter;
import it.mastropietroassicurazioni.scadenzario.notifications.NotificationType;
import it.mastropietroassicurazioni.scadenzario.notifications.NotificationsManager;
import it.mastropietroassicurazioni.scadenzario.sms.SmsSender;
import it.mastropietroassicurazioni.scadenzario.utils.ExcelTableIterator;
import it.mastropietroassicurazioni.scadenzario.utils.ExcelTableRow;

public class Scadenzario {

	private JFrame frmScadenzario;
	
	private JTabbedPane tabbedPane;
	private JPanel panel_scadenze;
	private JPanel panel_versamenti;
	private JPanel panel_compleanni;
	private JPanel panel_promemoria;
	private JPanel panel_configurazione;
	private JLabel lblPosizioneFile;
	private JLabel lblPosizioneTemplateIncassi;
	private JLabel lblNomeTabella;
	
	private JFileChooser sourceFilePicker;
	private JTextField textFieldPosizioneFile;
	private JTextField textFieldPosizioneTemplateIncassi;
	private JTextField textFieldPosizioneTemplateVersamenti;
	private JTable tableScadenze;
	private JTable tableCompleanni;
	private JTable tablePromemoria;
	
	private JButton btnScadenze;
	private JButton btnCompleanni;
	private JTextField textFieldSheetAnagrafica;
	private JLabel lblNomeTabellaAnagrafica;
	private JTextField textFieldTabellaAnagrafica;
	private JLabel lblNomeSheetPolizze;
	private JTextField textFieldSheetPolizze;
	private JLabel lblNomeTabellaPolizze;
	private JTextField textFieldTabellaPolizze;
	private JLabel lblNomeSheetPromemoria;
	private JTextField textFieldSheetPromemoria;
	private JLabel lblNomeTabellaPromemoria;
	private JTextField textFieldTabellaPromemoria;
	private JLabel lblIndirizzoServerSmtp;
	private JTextField textFieldSmtpServer;
	private JLabel lblUsername;
	private JTextField textFieldSmtpUsername;
	private JLabel lblPassword;
	private JTextField textFieldSmtpPassword;
	private JLabel lblMittenteMail;
	private JTextField textFieldSmtpMailFrom;
	private JLabel lblDestinatarioNascostoMail;
	private JTextField textFieldSmtpMailCc;
	private JButton btnSaveConfig;
	
	private ConfigBean configBean = new ConfigBean();
	private JLabel labelPosizioneTemplateVersamenti;
	private JPanel panel;
	private JLabel lblEsportaFoglioIncassi;
	private JTextField textField_dataExportIncassi;
	private JButton btnEsportaIncassi;
	private JLabel lblEsportaFoglioVersamenti;
	private JTextField textField_dataExportVersamenti;
	private JButton btnEsportaVersamenti;
	
	private static final DateFormat DATE_FORMAT = new SimpleDateFormat("dd/MM/yyyy");
	private static final Locale ITALIAN = new Locale("it", "IT", "EURO");
	private static final NumberFormat EURO_FORMAT = NumberFormat.getCurrencyInstance(ITALIAN);
	private static final int MILLISECS_PER_DAY = 1000 * 60 * 60 * 24;
	
	private JLabel lblNomeSheetVersamenti;
	private JLabel lblNomeTabellaVersamenti;
	private JTextField textFieldSheetVersamenti;
	private JTextField textFieldTabellaVersamenti;
	private JTable tableVersamenti;
	private JButton btnVersamenti;
	private JLabel lblNumeroGiorniPromemoria;
	private JLabel lblNumeroGiorniPromemoria_1;
	private JTextField textFieldGiorniPromemoriaScadenze;
	private JTextField textFieldGiorniPromemoriaVersamenti;
	private JLabel lblPortaServerSmtp;
	private JTextField textFieldSmtpPort;
	private JLabel label;
	private JLabel label_1;
	private JTextField textFieldUsernameSms;
	private JTextField textFieldPasswordSms;
	private JLabel lblNumeroProvenienzasms;
	private JTextField textFieldSmsNumberFrom;
	private JLabel lblSmsDisponibili;
	private JLabel lblValoreSmsDisponibili;
	
	

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Scadenzario window = new Scadenzario();
					window.frmScadenzario.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public Scadenzario() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmScadenzario = new JFrame();
		frmScadenzario.setTitle("Scadenzario");
		frmScadenzario.setBounds(100, 100, 830, 683);
		frmScadenzario.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		tabbedPane = new JTabbedPane();
		frmScadenzario.getContentPane().add(tabbedPane, BorderLayout.CENTER);
		
		/** SCADENZE **/
		
		panel_scadenze = new JPanel();
		panel_scadenze.setLayout(new BorderLayout());
		tabbedPane.addTab("Scadenze", null, panel_scadenze, null);
		
		tableScadenze = new JTable(new DefaultTableModel(new String[]{"SMS","Mail","Ignora","Data prossima scadenza","Contraente","Compagnia","Polizza","Stato polizza","Ramo","Auto/moto","Targa","Importo rata","Incassato anno","Tipo pagamento","Data attivazione","Data scadenza","Data ultimo pagamento","E-mail","SMS"},0){
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;
			
			public boolean isCellEditable(int row, int column) {                
				return column < 3;               
			};
			
			@Override
			public Class<?> getColumnClass(int columnIndex) {
				Class<?> clazz = null;
				for(int i=0; i<getRowCount(); i++){
					Object value = getValueAt(0, columnIndex);
					if(value != null){
						clazz = value.getClass();
						break;
					}
				}
				return clazz != null? clazz:String.class;
			}
		});
		tableScadenze.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		JScrollPane scrollPaneScadenze = new JScrollPane(tableScadenze);
		tableScadenze.setDefaultRenderer(Date.class, new DateCellRenderer());
		scrollPaneScadenze.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);
		scrollPaneScadenze.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_ALWAYS);
		panel_scadenze.add(scrollPaneScadenze, BorderLayout.CENTER);	
		btnScadenze = new JButton("Invia notifiche di scadenza");
		btnScadenze.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				NotificationsManager notificationsManager = new NotificationsManager();
				SmsSender smsSender = new SmsSender();
				EmailSender emailSender = new EmailSender();
				for(int rowIndex = 0; rowIndex < tableScadenze.getModel().getRowCount(); rowIndex++){
					Boolean sendSms = (Boolean)tableScadenze.getModel().getValueAt(rowIndex, 0);
					Boolean sendEmail = (Boolean)tableScadenze.getModel().getValueAt(rowIndex, 1);
					Boolean ignore = (Boolean)tableScadenze.getModel().getValueAt(rowIndex, 2);
					String email = (String)tableScadenze.getModel().getValueAt(rowIndex, 17);
					String sms = (String)tableScadenze.getModel().getValueAt(rowIndex, 18);
					
					String numeroPolizza = (String)tableScadenze.getModel().getValueAt(rowIndex, 6);
					Date dataPagamento = (Date)tableScadenze.getModel().getValueAt(rowIndex, 3);
					
					if(sendSms && !"".equals(sms.trim())){
						Map<String, Object> parameters = new HashMap<String, Object>();
						parameters.put("Data prossima scadenza", nullSafeFormat(dataPagamento));
						parameters.put("Contraente", tableScadenze.getModel().getValueAt(rowIndex, 4));
						parameters.put("Compagnia", tableScadenze.getModel().getValueAt(rowIndex, 5));
						parameters.put("Polizza", tableScadenze.getModel().getValueAt(rowIndex, 6));
						parameters.put("Stato polizza", tableScadenze.getModel().getValueAt(rowIndex, 7));
						parameters.put("Ramo", tableScadenze.getModel().getValueAt(rowIndex, 8));
						parameters.put("Auto/moto", tableScadenze.getModel().getValueAt(rowIndex, 9));
						parameters.put("Targa", tableScadenze.getModel().getValueAt(rowIndex, 10));
						parameters.put("Importo rata", tableScadenze.getModel().getValueAt(rowIndex, 11));
						parameters.put("Tipo pagamento", tableScadenze.getModel().getValueAt(rowIndex, 13));
						parameters.put("Data attivazione", nullSafeFormat((Date)tableScadenze.getModel().getValueAt(rowIndex, 14)));
						parameters.put("Data scadenza", nullSafeFormat((Date)tableScadenze.getModel().getValueAt(rowIndex, 15)));
						parameters.put("Data ultimo pagamento", nullSafeFormat((Date)tableScadenze.getModel().getValueAt(rowIndex, 16)));
						
						try{
							if(smsSender.send("polizza", sms, parameters, configBean)){
								notificationsManager.putNotifiedPolizza(numeroPolizza, dataPagamento, NotificationType.SMS);
								try{
									lblValoreSmsDisponibili.setText(Integer.toString(smsSender.checkSmsLeft(configBean)));
								}catch(IOException ioEx){
									//
								}
							}
						}catch(Exception smsEx){
							//
						}
						
					}
					if(sendEmail && !"".equals(email.trim())){
						Map<String, Object> parameters = new HashMap<String, Object>();
						parameters.put("Data prossima scadenza", nullSafeFormat(dataPagamento));
						parameters.put("Contraente", tableScadenze.getModel().getValueAt(rowIndex, 4));
						parameters.put("Compagnia", tableScadenze.getModel().getValueAt(rowIndex, 5));
						parameters.put("Polizza", tableScadenze.getModel().getValueAt(rowIndex, 6));
						parameters.put("Stato polizza", tableScadenze.getModel().getValueAt(rowIndex, 7));
						parameters.put("Ramo", tableScadenze.getModel().getValueAt(rowIndex, 8));
						parameters.put("Auto/moto", tableScadenze.getModel().getValueAt(rowIndex, 9));
						parameters.put("Targa", tableScadenze.getModel().getValueAt(rowIndex, 10));
						parameters.put("Importo rata", tableScadenze.getModel().getValueAt(rowIndex, 11));
						parameters.put("Tipo pagamento", tableScadenze.getModel().getValueAt(rowIndex, 13));
						parameters.put("Data attivazione", nullSafeFormat((Date)tableScadenze.getModel().getValueAt(rowIndex, 14)));
						parameters.put("Data scadenza", nullSafeFormat((Date)tableScadenze.getModel().getValueAt(rowIndex, 15)));
						parameters.put("Data ultimo pagamento", nullSafeFormat((Date)tableScadenze.getModel().getValueAt(rowIndex, 16)));
						
						try{
							emailSender.send("polizza", "Notifica scadenza polizza", email, parameters, configBean);
							notificationsManager.putNotifiedPolizza(numeroPolizza, dataPagamento, NotificationType.EMAIL);
						}catch(Exception smsEx){
							//
						}

					}
					if(ignore){
						notificationsManager.putNotifiedPolizza(numeroPolizza, dataPagamento, NotificationType.IGNORE);
					}
				}
				buildTables();
			}
		});
		panel_scadenze.add(btnScadenze, BorderLayout.SOUTH);
		

		/** VERSAMENTI **/
		
		panel_versamenti = new JPanel();
		tabbedPane.addTab("Versamenti", null, panel_versamenti, null);
		panel_versamenti.setLayout(new BorderLayout());
		
		tableVersamenti = new JTable(new DefaultTableModel(new String[]{"Ignora","Contraente","Compagnia","Polizza","Importo rata","Data incasso","Data prevista versamento"},0){
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;
			
			public boolean isCellEditable(int row, int column) {                
				return column < 1;               
			};
			
			@Override
			public Class<?> getColumnClass(int columnIndex) {
				Class<?> clazz = null;
				for(int i=0; i<getRowCount(); i++){
					Object value = getValueAt(0, columnIndex);
					if(value != null){
						clazz = value.getClass();
						break;
					}
				}
				return clazz != null? clazz:String.class;
			}
		});
		tableScadenze.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		JScrollPane scrollPaneVersamenti = new JScrollPane(tableVersamenti);
		tableVersamenti.setDefaultRenderer(Date.class, new DateCellRenderer());
		scrollPaneVersamenti.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);
		scrollPaneVersamenti.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_ALWAYS);
		panel_versamenti.add(scrollPaneVersamenti, BorderLayout.CENTER);	
		
		btnVersamenti = new JButton("Ignora notifiche di versamento");
		btnVersamenti.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				NotificationsManager notificationsManager = new NotificationsManager();
				for(int rowIndex = 0; rowIndex < tableVersamenti.getModel().getRowCount(); rowIndex++){
					Boolean ignore = (Boolean)tableVersamenti.getModel().getValueAt(rowIndex, 0);
					
					String numeroPolizza = (String)tableVersamenti.getModel().getValueAt(rowIndex, 3);
					Date dataVersamento = (Date)tableVersamenti.getModel().getValueAt(rowIndex, 6);
					
					if(ignore){
						notificationsManager.putNotifiedVersamento(numeroPolizza, dataVersamento, NotificationType.IGNORE);
					}
				}
				buildTables();
			}
		});
		panel_versamenti.add(btnVersamenti, BorderLayout.SOUTH);
		
		/** COMPLEANNI **/
		
		panel_compleanni = new JPanel();
		tabbedPane.addTab("Compleanni", null, panel_compleanni, null);
		panel_compleanni.setLayout(new BorderLayout(0, 0));
		
		tableCompleanni = new JTable(new DefaultTableModel(new String[]{"SMS","Mail","Nome contraente","Email","Telefono preferito"},0){
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {                
                return false;               
			};
			
			@Override
			public Class<?> getColumnClass(int columnIndex) {
				return getValueAt(0, columnIndex).getClass();
			}
		});
		tableCompleanni.setDefaultRenderer(Date.class, new DateCellRenderer());
		JScrollPane scrollPaneCompleanni = new JScrollPane(tableCompleanni);
		scrollPaneCompleanni.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);
		panel_compleanni.add(scrollPaneCompleanni);
		
		btnCompleanni = new JButton("Invia notifiche di compleanno");
		btnCompleanni.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				NotificationsManager notificationsManager = new NotificationsManager();
				for(int rowIndex = 0; rowIndex < tableCompleanni.getModel().getRowCount(); rowIndex++){
					String nomeContraente = (String)tableCompleanni.getModel().getValueAt(rowIndex, 2);
					String dataCompleanno = DATE_FORMAT.format(new Date());
					SmsSender smsSender = new SmsSender();
					EmailSender emailSender = new EmailSender();
					Boolean sendSms = (Boolean)tableCompleanni.getModel().getValueAt(rowIndex, 0);
					Boolean sendEmail = (Boolean)tableCompleanni.getModel().getValueAt(rowIndex, 1);
					String email = (String)tableCompleanni.getModel().getValueAt(rowIndex, 3);
					String sms = (String)tableCompleanni.getModel().getValueAt(rowIndex, 4);
						
					if(sendSms && !"".equals(sms.trim())){
						Map<String, Object> parameters = new HashMap<String, Object>();
						parameters.put("Contraente", nomeContraente);
							
						try{
							if(smsSender.send("compleanno", sms, parameters, configBean)){
								notificationsManager.putNotifiedCompleanno(nomeContraente, dataCompleanno, NotificationType.SMS);
								try{
									lblValoreSmsDisponibili.setText(Integer.toString(smsSender.checkSmsLeft(configBean)));
								}catch(IOException ioEx){
									//
								}
							}
						}catch(Exception smsEx){
							//
						}	
					}
					if(sendEmail && !"".equals(email.trim())){
						Map<String, Object> parameters = new HashMap<String, Object>();
						parameters.put("Contraente", nomeContraente);
							
						try{
							emailSender.send("compleanno", "Buon compleanno da Assicurazioni Mastropietro", email, parameters, configBean);
							notificationsManager.putNotifiedCompleanno(nomeContraente, dataCompleanno, NotificationType.EMAIL);
						}catch(Exception smsEx){
							//
						}	
					}
				}
				buildTables();
			}
		});
		panel_compleanni.add(btnCompleanni, BorderLayout.SOUTH);
		
		panel_promemoria = new JPanel();
		panel_promemoria.setLayout(new BorderLayout());
		tabbedPane.addTab("Promemoria", null, panel_promemoria, null);
		
		tablePromemoria = new JTable(new DefaultTableModel(new String[]{"Contraente","Tipo promemoria","Testo","E-mail","SMS"},0){
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;
			
			public boolean isCellEditable(int row, int column) {                
				return column < 2;               
			};
			
			@Override
			public Class<?> getColumnClass(int columnIndex) {
				Class<?> clazz = null;
				for(int i=0; i<getRowCount(); i++){
					Object value = getValueAt(0, columnIndex);
					if(value != null){
						clazz = value.getClass();
						break;
					}
				}
				return clazz != null? clazz:String.class;
			}
		});
		JScrollPane scrollPanePromemoria = new JScrollPane(tablePromemoria);
		tablePromemoria.setDefaultRenderer(Date.class, new DateCellRenderer());
		tablePromemoria.getColumnModel().getColumn(2).setCellRenderer(new TextAreaRenderer());
		tablePromemoria.getColumnModel().getColumn(2).setPreferredWidth(500);
		scrollPanePromemoria.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);
		panel_promemoria.add(scrollPanePromemoria, BorderLayout.CENTER);
		
		/** COMPLEANNI **/
		
		panel = new JPanel();
		tabbedPane.addTab("Esportazioni", null, panel, null);
		panel.setLayout(new FormLayout(new ColumnSpec[] {
				FormSpecs.RELATED_GAP_COLSPEC,
				FormSpecs.DEFAULT_COLSPEC,
				FormSpecs.RELATED_GAP_COLSPEC,
				FormSpecs.DEFAULT_COLSPEC,
				FormSpecs.RELATED_GAP_COLSPEC,
				ColumnSpec.decode("200px"),},
			new RowSpec[] {
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,}));
		
		lblEsportaFoglioIncassi = new JLabel("Esporta foglio incassi giornaliero");
		panel.add(lblEsportaFoglioIncassi, "2, 2, right, default");
		
		textField_dataExportIncassi = new JTextField();
		textField_dataExportIncassi.setText(DATE_FORMAT.format(new Date()));
		textField_dataExportIncassi.setColumns(10);
		panel.add(textField_dataExportIncassi, "4, 2, fill, default");
		
		btnEsportaIncassi = new JButton("Esporta incassi");
		btnEsportaIncassi.addMouseListener(new MouseAdapter(){
            @Override
            public void mouseClicked(MouseEvent e){
            	IncassoExporter incassoExporter = new IncassoExporter(textFieldPosizioneFile.getText(), 
            														  textFieldPosizioneTemplateIncassi.getText(), 
            														  configBean);
            	try {
					incassoExporter.export(DATE_FORMAT.parse(textField_dataExportIncassi.getText()));
				} catch (ParseException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
            }
        });
		panel.add(btnEsportaIncassi, "6, 2");
		
		lblEsportaFoglioVersamenti = new JLabel("Esporta foglio versamenti giornaliero");
		panel.add(lblEsportaFoglioVersamenti, "2, 4, right, default");
		
		textField_dataExportVersamenti = new JTextField();
		textField_dataExportVersamenti.setColumns(10);
		textField_dataExportVersamenti.setText(DATE_FORMAT.format(new Date()));
		panel.add(textField_dataExportVersamenti, "4, 4, fill, default");
		
		btnEsportaVersamenti = new JButton("Esporta versamenti");
		btnEsportaVersamenti.addMouseListener(new MouseAdapter(){
            @Override
            public void mouseClicked(MouseEvent e){
            	VersamentiExporter versamentiExporter = new VersamentiExporter(textFieldPosizioneFile.getText(), 
            														  	 	   textFieldPosizioneTemplateVersamenti.getText(), 
            														  	 	   configBean);
            	try {
            		versamentiExporter.export(DATE_FORMAT.parse(textField_dataExportVersamenti.getText()));
				} catch (ParseException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
            }
        });
		panel.add(btnEsportaVersamenti, "6, 4");
		
		
		panel_configurazione = new JPanel();
		tabbedPane.addTab("Configurazione", null, panel_configurazione, null);
		panel_configurazione.setLayout(new FormLayout(new ColumnSpec[] {
				FormSpecs.RELATED_GAP_COLSPEC,
				FormSpecs.DEFAULT_COLSPEC,
				FormSpecs.RELATED_GAP_COLSPEC,
				ColumnSpec.decode("default:grow"),
				FormSpecs.RELATED_GAP_COLSPEC,
				ColumnSpec.decode("default:grow"),
				FormSpecs.RELATED_GAP_COLSPEC,
				ColumnSpec.decode("default:grow"),},
			new RowSpec[] {
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,
				FormSpecs.RELATED_GAP_ROWSPEC,
				FormSpecs.DEFAULT_ROWSPEC,}));
		
		lblPosizioneFile = new JLabel("Posizione file");
		panel_configurazione.add(lblPosizioneFile, "2, 2, right, default");
		
		textFieldPosizioneFile = new JTextField();
		panel_configurazione.add(textFieldPosizioneFile, "4, 2, 5, 1, fill, default");
		textFieldPosizioneFile.setColumns(10);
		textFieldPosizioneFile.addMouseListener(new MouseAdapter(){
            @Override
            public void mouseClicked(MouseEvent e){
            	int returnValue = sourceFilePicker.showOpenDialog(frmScadenzario);
            	if (returnValue == JFileChooser.APPROVE_OPTION) {
                    File file = sourceFilePicker.getSelectedFile();
                    textFieldPosizioneFile.setText(file.getAbsolutePath());
                }
            }
        });
		
	   
		textFieldPosizioneFile.setText(configBean.getFileLocation());
		
		lblPosizioneTemplateIncassi = new JLabel("Posizione template incassi");
		panel_configurazione.add(lblPosizioneTemplateIncassi, "2, 4, right, default");
		
		textFieldPosizioneTemplateIncassi = new JTextField();
		textFieldPosizioneTemplateIncassi.setText(configBean.getIncassiTemplateLocation());
		panel_configurazione.add(textFieldPosizioneTemplateIncassi, "4, 4, 5, 1, fill, default");
		textFieldPosizioneTemplateIncassi.setColumns(10);
		
		textFieldPosizioneTemplateIncassi.addMouseListener(new MouseAdapter(){
            @Override
            public void mouseClicked(MouseEvent e){
            	int returnValue = sourceFilePicker.showOpenDialog(frmScadenzario);
            	if (returnValue == JFileChooser.APPROVE_OPTION) {
                    File file = sourceFilePicker.getSelectedFile();
                    textFieldPosizioneTemplateIncassi.setText(file.getAbsolutePath());
                }
            }
        });
		
	   
		
		labelPosizioneTemplateVersamenti = new JLabel("Posizione template versamenti");
		panel_configurazione.add(labelPosizioneTemplateVersamenti, "2, 6, right, default");
		
		textFieldPosizioneTemplateVersamenti = new JTextField();
		textFieldPosizioneTemplateVersamenti.setText(configBean.getVersamentiTemplateLocation());
		textFieldPosizioneTemplateVersamenti.setColumns(10);
		textFieldPosizioneTemplateVersamenti.addMouseListener(new MouseAdapter(){
            @Override
            public void mouseClicked(MouseEvent e){
            	int returnValue = sourceFilePicker.showOpenDialog(frmScadenzario);
            	if (returnValue == JFileChooser.APPROVE_OPTION) {
                    File file = sourceFilePicker.getSelectedFile();
                    textFieldPosizioneTemplateVersamenti.setText(file.getAbsolutePath());
                }
            }
        });
		panel_configurazione.add(textFieldPosizioneTemplateVersamenti, "4, 6, 5, 1, fill, default");
		
		lblNumeroGiorniPromemoria = new JLabel("Numero giorni promemoria scadenze");
		panel_configurazione.add(lblNumeroGiorniPromemoria, "2, 10, right, default");
		
		textFieldGiorniPromemoriaScadenze = new JTextField();
		textFieldGiorniPromemoriaScadenze.setText((String) null);
		textFieldGiorniPromemoriaScadenze.setColumns(10);
		panel_configurazione.add(textFieldGiorniPromemoriaScadenze, "4, 10, 5, 1, fill, default");
		
		lblNumeroGiorniPromemoria_1 = new JLabel("Numero giorni promemoria versamenti");
		panel_configurazione.add(lblNumeroGiorniPromemoria_1, "2, 12, right, default");
		
		textFieldGiorniPromemoriaVersamenti = new JTextField();
		textFieldGiorniPromemoriaVersamenti.setText((String) null);
		textFieldGiorniPromemoriaVersamenti.setColumns(10);
		panel_configurazione.add(textFieldGiorniPromemoriaVersamenti, "4, 12, 5, 1, fill, default");
		
		lblNomeTabella = new JLabel("Nome sheet anagrafica");
		panel_configurazione.add(lblNomeTabella, "2, 16, right, default");
		
		textFieldSheetAnagrafica = new JTextField();
		panel_configurazione.add(textFieldSheetAnagrafica, "4, 16, 5, 1, fill, default");
		textFieldSheetAnagrafica.setColumns(10);
		
		lblNomeTabellaAnagrafica = new JLabel("Nome tabella anagrafica");
		panel_configurazione.add(lblNomeTabellaAnagrafica, "2, 18, right, default");
		
		textFieldTabellaAnagrafica = new JTextField();
		textFieldTabellaAnagrafica.setColumns(10);
		panel_configurazione.add(textFieldTabellaAnagrafica, "4, 18, 5, 1, fill, default");
		
		lblNomeSheetPolizze = new JLabel("Nome sheet polizze");
		panel_configurazione.add(lblNomeSheetPolizze, "2, 20, right, default");
		
		textFieldSheetPolizze = new JTextField();
		textFieldSheetPolizze.setColumns(10);
		panel_configurazione.add(textFieldSheetPolizze, "4, 20, 5, 1, fill, default");
		
		lblNomeTabellaPolizze = new JLabel("Nome tabella polizze");
		panel_configurazione.add(lblNomeTabellaPolizze, "2, 22, right, default");
		
		textFieldTabellaPolizze = new JTextField();
		textFieldTabellaPolizze.setColumns(10);
		panel_configurazione.add(textFieldTabellaPolizze, "4, 22, 5, 1, fill, default");
		
		lblNomeSheetPromemoria = new JLabel("Nome sheet promemoria");
		panel_configurazione.add(lblNomeSheetPromemoria, "2, 24, right, default");
		
		textFieldSheetPromemoria = new JTextField();
		textFieldSheetPromemoria.setColumns(10);
		panel_configurazione.add(textFieldSheetPromemoria, "4, 24, 5, 1, fill, default");
		
		lblNomeTabellaPromemoria = new JLabel("Nome tabella promemoria");
		panel_configurazione.add(lblNomeTabellaPromemoria, "2, 26, right, default");
		
		textFieldTabellaPromemoria = new JTextField();
		textFieldTabellaPromemoria.setColumns(10);
		panel_configurazione.add(textFieldTabellaPromemoria, "4, 26, 5, 1, fill, default");
		
		lblNomeSheetVersamenti = new JLabel("Nome sheet versamenti");
		panel_configurazione.add(lblNomeSheetVersamenti, "2, 28, right, default");
		
		textFieldSheetVersamenti = new JTextField();
		textFieldSheetVersamenti.setText((String) null);
		textFieldSheetVersamenti.setColumns(10);
		panel_configurazione.add(textFieldSheetVersamenti, "4, 28, 5, 1, fill, default");
		
		lblNomeTabellaVersamenti = new JLabel("Nome tabella versamenti");
		panel_configurazione.add(lblNomeTabellaVersamenti, "2, 30, right, default");
		
		textFieldTabellaVersamenti = new JTextField();
		textFieldTabellaVersamenti.setText((String) null);
		textFieldTabellaVersamenti.setColumns(10);
		panel_configurazione.add(textFieldTabellaVersamenti, "4, 30, 5, 1, fill, default");
		
		lblIndirizzoServerSmtp = new JLabel("Indirizzo server SMTP");
		panel_configurazione.add(lblIndirizzoServerSmtp, "2, 34, right, default");
		
		textFieldSmtpServer = new JTextField();
		textFieldSmtpServer.setColumns(10);
		panel_configurazione.add(textFieldSmtpServer, "4, 34, fill, default");
		
		label = new JLabel("Username SMS");
		panel_configurazione.add(label, "6, 34, right, default");
		
		textFieldUsernameSms = new JTextField();
		textFieldUsernameSms.setText(configBean.getSmsUsername());
		textFieldUsernameSms.setColumns(10);
		panel_configurazione.add(textFieldUsernameSms, "8, 34, fill, default");
		
		lblPortaServerSmtp = new JLabel("Porta server SMTP");
		panel_configurazione.add(lblPortaServerSmtp, "2, 36, right, default");
		
		textFieldSmtpPort = new JTextField();
		textFieldSmtpPort.setText((String) null);
		textFieldSmtpPort.setColumns(10);
		panel_configurazione.add(textFieldSmtpPort, "4, 36, fill, default");
		
		label_1 = new JLabel("Password SMS");
		panel_configurazione.add(label_1, "6, 36, right, default");
		
		textFieldPasswordSms = new JTextField();
		textFieldPasswordSms.setText(configBean.getSmsPassword());
		textFieldPasswordSms.setColumns(10);
		panel_configurazione.add(textFieldPasswordSms, "8, 36, fill, default");
		
		lblUsername = new JLabel("Username SMTP");
		panel_configurazione.add(lblUsername, "2, 38, right, default");
		
		textFieldSmtpUsername = new JTextField();
		textFieldSmtpUsername.setColumns(10);
		panel_configurazione.add(textFieldSmtpUsername, "4, 38, fill, default");
		
		lblNumeroProvenienzasms = new JLabel("Numero provenienza SMS");
		panel_configurazione.add(lblNumeroProvenienzasms, "6, 38, right, default");
		
		textFieldSmsNumberFrom = new JTextField();
		textFieldSmsNumberFrom.setText(configBean.getSmsNumberFrom());
		textFieldSmsNumberFrom.setColumns(10);
		panel_configurazione.add(textFieldSmsNumberFrom, "8, 38, fill, default");
		
		lblPassword = new JLabel("Password SMTP");
		panel_configurazione.add(lblPassword, "2, 40, right, default");
		
		textFieldSmtpPassword = new JTextField();
		textFieldSmtpPassword.setColumns(10);
		panel_configurazione.add(textFieldSmtpPassword, "4, 40, fill, default");
		
		lblSmsDisponibili = new JLabel("SMS disponibili");
		panel_configurazione.add(lblSmsDisponibili, "6, 40, right, default");
		
		lblValoreSmsDisponibili = new JLabel("-");
		lblValoreSmsDisponibili.setFont(new Font("Tahoma", Font.BOLD, 11));
		panel_configurazione.add(lblValoreSmsDisponibili, "8, 40");
		SmsSender smsSender = new SmsSender();
		try{
			lblValoreSmsDisponibili.setText(Integer.toString(smsSender.checkSmsLeft(configBean)));
		}catch(IOException ioEx){
			//
		}
		
		lblMittenteMail = new JLabel("Mittente mail");
		panel_configurazione.add(lblMittenteMail, "2, 42, right, default");
		
		textFieldSmtpMailFrom = new JTextField();
		textFieldSmtpMailFrom.setColumns(10);
		panel_configurazione.add(textFieldSmtpMailFrom, "4, 42, fill, default");
		
		lblDestinatarioNascostoMail = new JLabel("Destinatario nascosto mail");
		panel_configurazione.add(lblDestinatarioNascostoMail, "2, 44, right, default");
		
		textFieldSmtpMailCc = new JTextField();
		textFieldSmtpMailCc.setColumns(10);
		panel_configurazione.add(textFieldSmtpMailCc, "4, 44, fill, default");
		
		sourceFilePicker = new JFileChooser();
	    FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel spreadsheets", "xlsm", "xslx", "xls");
	    sourceFilePicker.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
	    sourceFilePicker.setFileFilter(filter);
	    
	    textFieldSheetAnagrafica.setText(configBean.getSheetAnagrafica());
	    textFieldSheetPolizze.setText(configBean.getSheetPolizze());
	    textFieldSheetPromemoria.setText(configBean.getSheetPromemoria());
	    textFieldSheetVersamenti.setText(configBean.getSheetVersamenti());
	    textFieldTabellaAnagrafica.setText(configBean.getTabellaAnagrafica());
	    textFieldTabellaPolizze.setText(configBean.getTabellaPolizze());
	    textFieldTabellaPromemoria.setText(configBean.getTabellaPromemoria());
	    textFieldTabellaVersamenti.setText(configBean.getTabellaVersamenti());
	    textFieldSmtpServer.setText(configBean.getSmtpServer());
	    textFieldSmtpPort.setText(Integer.toString(configBean.getSmtpPort()));
	    textFieldSmtpUsername.setText(configBean.getSmtpUsername());
	    textFieldSmtpPassword.setText(configBean.getSmtpPassword());
	    textFieldSmtpMailFrom.setText(configBean.getSmtpMailFrom());
	    textFieldSmtpMailCc.setText(configBean.getSmtpMailCc());
	    textFieldGiorniPromemoriaScadenze.setText(Integer.toString(configBean.getGiorniPromemoriaScadenze()));
	    textFieldGiorniPromemoriaVersamenti.setText(Integer.toString(configBean.getGiorniPromemoriaVersamenti()));
	    
	    btnSaveConfig = new JButton("Salva configurazioni");
	    panel_configurazione.add(btnSaveConfig, "1, 48, 8, 1");
	    btnSaveConfig.addMouseListener(new MouseAdapter() {
	    	@Override
	    	public void mouseClicked(MouseEvent e) {
	    		configBean.setFileLocation(textFieldPosizioneFile.getText());
	    		configBean.setIncassiTemplateLocation(textFieldPosizioneTemplateIncassi.getText());
	    		configBean.setVersamentiTemplateLocation(textFieldPosizioneTemplateVersamenti.getText());
	    	    configBean.setSheetAnagrafica(textFieldSheetAnagrafica.getText());
	    	    configBean.setSheetPolizze(textFieldSheetPolizze.getText());
	    	    configBean.setSheetPromemoria(textFieldSheetPromemoria.getText());
	    	    configBean.setSheetVersamenti(textFieldSheetVersamenti.getText());
	    	    configBean.setTabellaAnagrafica(textFieldTabellaAnagrafica.getText());
	    	    configBean.setTabellaPolizze(textFieldTabellaPolizze.getText());
	    	    configBean.setTabellaPromemoria(textFieldTabellaPromemoria.getText());
	    	    configBean.setTabellaVersamenti(textFieldTabellaVersamenti.getText());
	    	    configBean.setSmtpServer(textFieldSmtpServer.getText());
	    	    configBean.setSmtpPort(Integer.parseInt(textFieldSmtpPort.getText()));
	    	    configBean.setSmtpUsername(textFieldSmtpUsername.getText());
	    	    configBean.setSmtpPassword(textFieldSmtpPassword.getText());
	    	    configBean.setSmtpMailFrom(textFieldSmtpMailFrom.getText());
	    	    configBean.setSmtpMailCc(textFieldSmtpMailCc.getText());
	    	    configBean.setSmsUsername(textFieldUsernameSms.getText());
	    	    configBean.setSmsPassword(textFieldPasswordSms.getText());
	    	    configBean.setSmsNumberFrom(textFieldSmsNumberFrom.getText());
	    	    configBean.setGiorniPromemoriaScadenze(Integer.parseInt(textFieldGiorniPromemoriaScadenze.getText()));
	    	    configBean.setGiorniPromemoriaVersamenti(Integer.parseInt(textFieldGiorniPromemoriaVersamenti.getText()));
	    	    configBean.updateConfig();
	    	}
	    });
	
	    buildTables();
	}
	
	private void buildTables(){
		DefaultTableModel modelScadenze = (DefaultTableModel) tableScadenze.getModel();
		DefaultTableModel modelVersamenti = (DefaultTableModel) tableVersamenti.getModel();
		DefaultTableModel modelPromemoria = (DefaultTableModel) tablePromemoria.getModel();
		DefaultTableModel modelCompleanni = (DefaultTableModel) tableCompleanni.getModel();
		
		for(DefaultTableModel dm: new DefaultTableModel[]{modelScadenze, modelVersamenti, modelPromemoria, modelCompleanni}){
			int rowCount = dm.getRowCount();
			//Remove rows one by one from the end of the table
			for (int i = rowCount - 1; i >= 0; i--) {
			    dm.removeRow(i);
			}
		}
		
		NotificationsManager notificationsManager = new NotificationsManager();
		XSSFWorkbook workbook = null;
		try {	
			Files.copy(Paths.get(configBean.getFileLocation()), 
					   Paths.get("tmp_file_polizze.xlsm"), 
					   StandardCopyOption.REPLACE_EXISTING);
			
			workbook = new XSSFWorkbook(new File("tmp_file_polizze.xlsm"));		
			
			ExcelTableIterator excelTableIteratorScadenze = new ExcelTableIterator(workbook, configBean.getSheetPolizze(), configBean.getTabellaPolizze());
			while(excelTableIteratorScadenze.hasNext()){
				ExcelTableRow row = excelTableIteratorScadenze.next();
				Date nextPaymentDay = row.getValueAsDate("Data prossima scadenza");
				String numeroPolizza = row.getValue("Polizza");
				String statoPolizza = row.getValue("Stato polizza");
				Date today = new Date();
				
				long tempoScadenza = -1;
				
				if("ATTIVA".equalsIgnoreCase(statoPolizza) && nextPaymentDay != null){
				
					tempoScadenza = (nextPaymentDay.getTime() - today.getTime());
				
					if(tempoScadenza <= (configBean.getGiorniPromemoriaScadenze() * MILLISECS_PER_DAY) && notificationsManager.getNotifiedPolizza(numeroPolizza, nextPaymentDay) == null) {
						modelScadenze.addRow(new Object[]{
													"SI".equals(row.getValue("Consenso SMS")) && row.getValue("Telefono preferito") != null && !"".equals(row.getValue("Telefono preferito").trim()),
													"SI".equals(row.getValue("Consenso MAIL")) && row.getValue("E-mail") != null && !"".equals(row.getValue("E-mail").trim()),
													false,
													row.getValueAsDate("Data prossima scadenza"),
													row.getValue("Contraente"),
													row.getValue("Compagnia"),
													row.getValue("Polizza"),
													row.getValue("Stato polizza"),
													row.getValue("Ramo"),
													row.getValue("Auto/moto"),
													row.getValue("Targa"),
													EURO_FORMAT.format(row.getValueAsNumber("Importo rata")),
													EURO_FORMAT.format(row.getValueAsNumber("Incassato anno")),
													row.getValue("Tipo pagamento"),
													row.getValueAsDate("Data attivazione"),
													row.getValueAsDate("Data scadenza"),
													row.getValueAsDate("Data ultimo pagamento"),
        								  		    row.getValue("Consenso MAIL").equals("SI")?row.getValue("E-mail"):"",
        								  		    row.getValue("Consenso SMS").equals("SI")?row.getValue("Telefono preferito"):""});
					}
				}
			}
			
			ExcelTableIterator excelTableIteratorVersamenti = new ExcelTableIterator(workbook, configBean.getSheetVersamenti(), configBean.getTabellaVersamenti());
			while(excelTableIteratorVersamenti.hasNext()){
				ExcelTableRow row = excelTableIteratorVersamenti.next();
				Date dataVersamento = row.getValueAsDate("Data prevista versamento");
				Date today = new Date();
				String numeroPolizza = row.getValue("Polizza");
				
				long tempoScadenza = -1;
				
				if(dataVersamento != null){
				
					tempoScadenza = (dataVersamento.getTime() - today.getTime());
				
					if("IN ATTESA".equalsIgnoreCase(row.getValue("Stato versamento")) &&
					   tempoScadenza <= (configBean.getGiorniPromemoriaVersamenti() * MILLISECS_PER_DAY) && 
					   notificationsManager.getNotifiedVersamento(numeroPolizza, dataVersamento) == null) {
						modelVersamenti.addRow(new Object[]{
													false,
													row.getValue("Contraente"),
													row.getValue("Compagnia"),
													row.getValue("Polizza"),
													EURO_FORMAT.format(row.getValueAsNumber("Importo rata")),
													row.getValueAsDate("Data incasso"),
													row.getValueAsDate("Data prevista versamento")});
					}
				}
			}
			
			ExcelTableIterator excelTableIteratorPromemoria = new ExcelTableIterator(workbook, configBean.getSheetPromemoria(), configBean.getTabellaPromemoria());
			while(excelTableIteratorPromemoria.hasNext()){
				ExcelTableRow row = excelTableIteratorPromemoria.next();
				Date reminderDay = row.getValueAsDate("Data promemoria");
				Date today = new Date();
				
				boolean promemoriaOggi = reminderDay != null && 
									     reminderDay.getDate() == today.getDate() && 
									     reminderDay.getMonth() == today.getMonth() && 
									     reminderDay.getYear() == today.getYear();
					
				if(promemoriaOggi) {
					modelPromemoria.addRow(new Object[]{
													row.getValue("Contraente"),
													row.getValue("Tipo promemoria"),
													row.getValue("Testo")});
				}
			}
			
			
			ExcelTableIterator excelTableIteratorCompleanni = new ExcelTableIterator(workbook, configBean.getSheetAnagrafica(), configBean.getTabellaAnagrafica());
			while(excelTableIteratorCompleanni.hasNext()){
				ExcelTableRow row = excelTableIteratorCompleanni.next();
				Date birthDate = row.getValueAsDate("DATA NASCITA");
				if(birthDate != null && birthDate.getDate() == new Date().getDate() && birthDate.getMonth() == new Date().getMonth()) {
					modelCompleanni.addRow(new Object[]{
														"SI".equals(row.getValue("Consenso SMS")) && row.getValue("Telefono preferito") != null && !"".equals(row.getValue("Telefono preferito").trim()),
														"SI".equals(row.getValue("Consenso MAIL")) && row.getValue("MAIL") != null && !"".equals(row.getValue("MAIL").trim()),
														row.getValue("Nome/Ragione sociale"),
            								  		    row.getValue("Consenso MAIL").equals("SI")?row.getValue("MAIL"):"",
            								  		    row.getValue("Consenso SMS").equals("SI")?row.getValue("Telefono preferito"):""});
            		
				}
			}
		} catch (InvalidFormatException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}finally {
			try {
				workbook.close();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
		}
	}
	
	class DateCellRenderer extends DefaultTableCellRenderer {

	    SimpleDateFormat f = new SimpleDateFormat("dd/MM/yyyy");

	    public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
	        if( value instanceof Date) {
	            value = f.format(value);
	        }
	        return super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
	    }
	};
	
	class TextAreaRenderer extends JTextArea implements TableCellRenderer { 
	    private final DefaultTableCellRenderer renderer = new DefaultTableCellRenderer(); 
	 
	    // Column heights are placed in this Map 
	    private final Map<JTable, Map<Object, Map<Object, Integer>>> tablecellSizes = new HashMap<JTable, Map<Object, Map<Object, Integer>>>(); 
	 
	    /** 
	     * Creates a text area renderer. 
	     */ 
	    public TextAreaRenderer() { 
	        setLineWrap(true); 
	        setWrapStyleWord(true); 
	    } 
	 
	    /** 
	     * Returns the component used for drawing the cell.  This method is 
	     * used to configure the renderer appropriately before drawing. 
	     * 
	     * @param table      - JTable object 
	     * @param value      - the value of the cell to be rendered. 
	     * @param isSelected - isSelected   true if the cell is to be rendered with the selection highlighted; 
	     *                   otherwise false. 
	     * @param hasFocus   - if true, render cell appropriately. 
	     * @param row        - The row index of the cell being drawn. 
	     * @param column     - The column index of the cell being drawn. 
	     * @return - Returns the component used for drawing the cell. 
	     */ 
	    public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, 
	                                                   boolean hasFocus, int row, int column) { 
	        // set the Font, Color, etc. 
	        renderer.getTableCellRendererComponent(table, value, 
	                isSelected, hasFocus, row, column); 
	        setForeground(renderer.getForeground()); 
	        setBackground(renderer.getBackground()); 
	        setBorder(renderer.getBorder()); 
	        setFont(renderer.getFont()); 
	        setText(renderer.getText()); 
	 
	        TableColumnModel columnModel = table.getColumnModel(); 
	        setSize(columnModel.getColumn(column).getWidth(), 0); 
	        int height_wanted = (int) getPreferredSize().getHeight(); 
	        addSize(table, row, column, height_wanted); 
	        height_wanted = findTotalMaximumRowSize(table, row); 
	        if (height_wanted != table.getRowHeight(row)) { 
	            table.setRowHeight(row, height_wanted); 
	        } 
	        return this; 
	    } 
	 
	    /** 
	     * @param table  - JTable object 
	     * @param row    - The row index of the cell being drawn. 
	     * @param column - The column index of the cell being drawn. 
	     * @param height - Row cell height as int value 
	     *               This method will add size to cell based on row and column number 
	     */ 
	    private void addSize(JTable table, int row, int column, int height) { 
	        Map<Object, Map<Object, Integer>> rowsMap = tablecellSizes.get(table); 
	        if (rowsMap == null) { 
	            tablecellSizes.put(table, rowsMap = new HashMap<Object, Map<Object, Integer>>()); 
	        } 
	        Map<Object, Integer> rowheightsMap = rowsMap.get(row); 
	        if (rowheightsMap == null) { 
	            rowsMap.put(row, rowheightsMap = new HashMap<Object, Integer>()); 
	        } 
	        rowheightsMap.put(column, height); 
	    } 
	 
	    /** 
	     * Look through all columns and get the renderer.  If it is 
	     * also a TextAreaRenderer, we look at the maximum height in 
	     * its hash table for this row. 
	     * 
	     * @param table -JTable object 
	     * @param row   - The row index of the cell being drawn. 
	     * @return row maximum height as integer value 
	     */ 
	    private int findTotalMaximumRowSize(JTable table, int row) { 
	        int maximum_height = 0; 
	        Enumeration<TableColumn> columns = table.getColumnModel().getColumns(); 
	        while (columns.hasMoreElements()) { 
	            TableColumn tc = columns.nextElement(); 
	            TableCellRenderer cellRenderer = tc.getCellRenderer(); 
	            if (cellRenderer instanceof TextAreaRenderer) { 
	                TextAreaRenderer tar = (TextAreaRenderer) cellRenderer; 
	                maximum_height = Math.max(maximum_height, 
	                        tar.findMaximumRowSize(table, row)); 
	            } 
	        } 
	        return maximum_height; 
	    } 
	 
	    /** 
	     * This will find the maximum row size 
	     * 
	     * @param table - JTable object 
	     * @param row   - The row index of the cell being drawn. 
	     * @return row maximum height as integer value 
	     */ 
	    private int findMaximumRowSize(JTable table, int row) { 
	        Map<Object, Map<Object, Integer>> rows = tablecellSizes.get(table); 
	        if (rows == null) return 0; 
	        Map<Object, Integer> rowheights = rows.get(row); 
	        if (rowheights == null) return 0; 
	        int maximum_height = 0; 
	        for (Map.Entry<Object, Integer> entry : rowheights.entrySet()) { 
	            int cellHeight = entry.getValue(); 
	            maximum_height = Math.max(maximum_height, cellHeight); 
	        } 
	        return maximum_height; 
	    } 
	}
	
	private String nullSafeFormat(Date date){
		return date == null? "":DATE_FORMAT.format(date);
	}

	
}
