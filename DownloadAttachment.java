import java.io.BufferedReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.net.URL;
import java.nio.charset.Charset;
import java.rmi.RemoteException;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;

import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.Vector;
import java.text.*;
import java.util.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Properties;
import java.util.regex.Pattern;

import javax.swing.filechooser.FileSystemView;
import javax.swing.text.Document;
import javax.xml.namespace.QName;
import javax.xml.rpc.ServiceException;
import org.apache.axis.message.MessageElement;
import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Level;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.log4j.xml.DOMConfigurator;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;


import org.w3c.dom.Element;
import org.apache.commons.codec.binary.Base64;

import com.sforce.soap.partner.DeleteResult;
import com.sforce.soap.partner.DescribeGlobalResult;
import com.sforce.soap.partner.DescribeLayout;
import com.sforce.soap.partner.DescribeLayoutItem;
import com.sforce.soap.partner.DescribeLayoutResult;
import com.sforce.soap.partner.DescribeLayoutRow;
import com.sforce.soap.partner.DescribeLayoutSection;
import com.sforce.soap.partner.DescribeSObjectResult;
import com.sforce.soap.partner.DescribeSoftphoneLayoutCallType;
import com.sforce.soap.partner.DescribeSoftphoneLayoutInfoField;
import com.sforce.soap.partner.DescribeSoftphoneLayoutItem;
import com.sforce.soap.partner.DescribeSoftphoneLayoutResult;
import com.sforce.soap.partner.DescribeSoftphoneLayoutSection;
import com.sforce.soap.partner.DescribeTab;
import com.sforce.soap.partner.DescribeTabSetResult;
import com.sforce.soap.partner.EmailPriority;
import com.sforce.soap.partner.EmptyRecycleBinResult;
import com.sforce.soap.partner.Field;
import com.sforce.soap.partner.FieldType;
import com.sforce.soap.partner.GetDeletedResult;
import com.sforce.soap.partner.GetUpdatedResult;
import com.sforce.soap.partner.GetUserInfoResult;
import com.sforce.soap.partner.LeadConvert;
import com.sforce.soap.partner.LeadConvertResult;
import com.sforce.soap.partner.LoginResult;
import com.sforce.soap.partner.MergeRequest;
import com.sforce.soap.partner.MergeResult;
import com.sforce.soap.partner.PicklistEntry;
import com.sforce.soap.partner.ProcessRequest;
import com.sforce.soap.partner.ProcessResult;
import com.sforce.soap.partner.ProcessSubmitRequest;
import com.sforce.soap.partner.ProcessWorkitemRequest;
import com.sforce.soap.partner.QueryOptions;
import com.sforce.soap.partner.QueryResult;
import com.sforce.soap.partner.ResetPasswordResult;
import com.sforce.soap.partner.SaveResult;
import com.sforce.soap.partner.SearchRecord;
import com.sforce.soap.partner.SearchResult;
import com.sforce.soap.partner.SendEmailResult;
import com.sforce.soap.partner.SessionHeader;
import com.sforce.soap.partner.SetPasswordResult;
import com.sforce.soap.partner.SforceServiceLocator;
import com.sforce.soap.partner.SingleEmailMessage;
import com.sforce.soap.partner.SoapBindingStub;
import com.sforce.soap.partner.UndeleteResult;
import com.sforce.soap.partner.UpsertResult;
import com.sforce.soap.partner.fault.ApiFault;
import com.sforce.soap.partner.fault.InvalidQueryLocatorFault;
import com.sforce.soap.partner.fault.LoginFault;
import com.sforce.soap.partner.fault.UnexpectedErrorFault;
import com.sforce.soap.partner.sobject.SObject;

import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.graphics.Point;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Text;
import org.eclipse.swt.SWT;
//import org.eclipse.swt.graphics.Rectangle;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.custom.CLabel;


import java.awt.BorderLayout;
import java.awt.Image;
import java.awt.Insets;
import java.awt.Toolkit;

import javax.swing.JPanel;
import javax.swing.JFrame;
import java.awt.Dimension;
import javax.swing.JLabel;
import java.awt.Rectangle;

import javax.swing.Action;
import javax.swing.ImageIcon;
import javax.swing.JTextField;
import javax.swing.JPasswordField;
import javax.swing.JButton;
import javax.swing.JTextArea;
import javax.swing.SwingWorker;

import java.awt.GridBagLayout;
import java.awt.GridBagConstraints;
import java.beans.PropertyChangeEvent;

import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JComboBox;

import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.awt.ComponentOrientation;
import org.apache.tools.zip.ZipEntry;
import org.apache.tools.zip.ZipOutputStream;


public class DownloadAttachment extends JFrame {

	private static final long serialVersionUID = 1L;
	private JPanel jContentPane = null;
	private JLabel lblUserName = null;
	private JLabel lblPassword = null;
	private JTextField txtUserName = null;
	private JPasswordField txtPassword = null;
	private JButton btnLogin = null;
	private JButton btnCopy = null;
	private JTextArea tarMessage = null;

	private String csvRow;  //  @jve:decl-index=0:
	private boolean logon;
	private final Log log;
	private SoapBindingStub binding;
	private Text txaMessage = null;
	private JPanel pnlLogin = null;
	private JPanel pnlDownload = null;
	private JLabel lblDays = null;
	private JButton btnDownload = null;
	private JButton btnClose = null;
	private JProgressBar prgbDownload = null;

	private FileOutputStream fos = null;  //  @jve:decl-index=0:
	private OutputStreamWriter writer = null;
	private String[] titles;
	private JProgressBar prgbAllSteps = null;
	private JScrollPane spnMessage = null;
	private JComboBox cbbDays = null;
	private JLabel lblAll = null;
	private JLabel lblStep = null;
	final static private SimpleDateFormat timeFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
	final static private SimpleDateFormat logTimeFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");  //  @jve:decl-index=0:

	private boolean blSuccess = true;
	private String proxyHost = null;
	private String proxyPort = null;
	private String csvFile = null;
	private String attachFile = null;
	private String duration = null;
	private String saveFolder = null;
	private String Headquarters = null;
	private static final String propertyFile = "./conf/sfdc.properties";  //  @jve:decl-index=0:
	private final String Soap_Address_Product = "https://www.salesforce.com/services/Soap/u/26.0";
	private final String Soap_Address_Sandbox = "https://test.salesforce.com/services/Soap/u/26.0";
	private String hp_path = "\\HP用物件データ";
	private String work_path = "\\work";
	private String log_path = "\\log";
	private String csvFileName = "";  //  @jve:decl-index=0:
	private int durationDays = 0;

	/**
	 * This is the default constructor
	 */
	public DownloadAttachment() {
		super();
		PropertyConfigurator.configure("./conf/log4j.properties");
		log = LogFactory.getLog(getClass());
		getProperties();
		initialize();
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		titles = new String[]{
				"店舗コード","シリアルコード","物件コード","物件登録日","物件更新日","物件公開日","物件種別","物件種類","物件種目",
				"郵便番号","所在地(都道府県）","所在地（区市町村)","所在地（大字・通称）","所在地（字名・丁目)","所在地（地番以下)",
				"物件名称","小学区","中学区","セールスポイント","備考","おすすめ区分","おすすめ区分２","店舗整理コード","沿線１","駅１",
				"沿線なし指定","徒歩１","バス停１","バス１","停歩１","車１","距離１","沿線２","駅２","徒歩２","バス停２","バス２",
				"停歩２","車２","距離２","価格","消費税区分","坪単価","地代","管理費","管理費有無","修繕積立費","修繕積立費有無",
				"修繕積立基金","路線価","都市計画","用途地域","建ぺい率","容積率","国土法の届出","地目","土地面積基準","土地面積・㎡",
				"土地面積・坪","接道種別１","接道方向１","接道幅員１","接道接面幅１","接道種別２","接道方向２","接道幅員２","接道接面幅２",
				"道路位置指定","不適合接道","土地現況","建築条件","土地権利","土地持分","ガス","水道","汚水","雑排水","私道有無","私道持分",
				"私道面積基準","私道面積","セットバック要否","セットバック部分","セットバック面積","敷延有無","敷延面積","借地条件","借地期間",
				"借地開始日","区画整理","計画道路","開発許可番号","地域地区","特記事項","建物構造","建築階数","建築階数（地下）","建築面積",
				"建築延床面積・㎡","建築延床面積・坪","専有面積基準","専有面積・㎡","専有面積・坪","バルコニー面積","ルーフバルコニー面積","総戸数",
				"販売戸数","部屋地下フラグ","部屋の階","部屋番号","建物向き","築年月","築年 和暦（年号）","築年 和暦（年）","建物現況","建築確認番号",
				"管理形態","管理会社","施工会社","間取","間取詳細１（階）","間取詳細１（タイプ）","間取詳細１（広さ）","間取詳細２（階）","間取詳細２（タイプ）",
				"間取詳細２（広さ）","間取詳細３（階）","間取詳細３（タイプ）","間取詳細３（広さ）","間取詳細４（階）","間取詳細４（タイプ）","間取詳細４（広さ）",
				"間取詳細５（階）","間取詳細５（タイプ）","間取詳細５（広さ）","間取詳細６（階）","間取詳細６（タイプ）","間取詳細６（広さ）","間取詳細７（階）",
				"間取詳細７（タイプ）","間取詳細７（広さ）","間取詳細８（階）","間取詳細８（タイプ）","間取詳細８（広さ）","駐車場状況","駐車場料金",
				"駐車場タイプ","駐車可能台数","駐輪場","インターネット設備","設備その他","設備詳細フォーマットタイプ","設備詳細","当社の取引態様",
				"当社の担当者コード","当社の担当者名","当社の手数料","手数料マーク","物件確認書送信日","物件シート有無","商談中フラグ","引渡条件",
				"引渡時期","引渡日（年）","引渡日（月）","引渡日（日）","特集コード","特集名","特集エリアコード","特集エリア","ＵＲＬ種別１","ＵＲＬ１",
				"ＵＲＬ種別２","ＵＲＬ２","公的融資情報","業者向けコメント","売主１","売主１住所","売主１電話番号","売主１携帯番号","売主２","売主２住所",
				"売主２電話番号","売主２携帯番号","提供会社コード","提供会社名","情報提供会社電話番号","提供会社FAX","提供会社取引態様",
				"提供会社手数料","提供会社取引状況","取引状況確認日","広告掲載可否","インターネット掲載の可否","掲載確認日","提供会社物件コード",
				"提供会社担当者","別業者","情報媒体コード","情報媒体名","画像１（間取）","画像２（外観）","画像３","画像４","画像５","画像６","画像７",
				"画像８","パノラマ公開","パノラマURL","シート用文章","オープンハウス区分","オープン開始日","オープン開始時間（時）","オープン開始時間（分）",
				"オープン終了日","オープン終了時間（時）","オープン終了時間（分）","オープンハウス受付場所","売却済物件","売却価格","売却業者",
				"売却年月日","売却先顧客番号","売り止め物件","インターネット公開","インターネット公開ＩＤ","公開先フォーマットタイプ","公開サイト",
				"物件公開（店舗）","物件公開（マッチングメール）","物件公開（タッチパネル）","客付可否","当社が支払う手数料区分","当社が支払う手数料率",
				"当社が支払う手数料金","交通起点・最寄施設","保証金","保証金（単位）","権利金","権利金（単位）","最適用途","管理方式","楽器相談",
				"事務所","二人入居","性別限定","単身者","法人","学生","高齢者","公庫利用","手付金保証","階毎の延床面積（1F）・㎡",
				"階毎の延床面積（1F）・坪","階毎の延床面積（2F）・㎡","階毎の延床面積（2F）・坪","階毎の延床面積（その他）","物件シート用備考",
				"インターネット初回公開日","緯度(WGS)","経度(WGS)","緯度(Tokyo97)","経度(Tokyo97)","地図公開","地勢","接道状況","角地",
				"築年月不詳","画像コメント１","画像コメント２","画像コメント３","画像コメント４","画像コメント５","画像コメント６","画像コメント７",
				"画像コメント８","エリア指定","建築許可理由","構造その他","駐車場の場所","レインズ登録済","社内向け備考","アーカイブ","画像９",
				"画像１０","画像１１","画像１２","画像１３","画像１４","画像１５","画像１６","画像１７","画像１８","画像１９","画像２０","画像２１",
				"画像２２","画像２３","画像２４","画像２５","画像コメント９","画像コメント１０","画像コメント１１","画像コメント１２","画像コメント１３",
				"画像コメント１４","画像コメント１５","画像コメント１６","画像コメント１７","画像コメント１８","画像コメント１９","画像コメント２０","画像コメント２１",
				"画像コメント２２","画像コメント２３","画像コメント２４","画像コメント２５","画像タイトル１","画像タイトル２","画像タイトル３","画像タイトル４",
				"画像タイトル５","画像タイトル６","画像タイトル７","画像タイトル８","画像タイトル９","画像タイトル１０","画像タイトル１１","画像タイトル１２",
				"画像タイトル１３","画像タイトル１４","画像タイトル１５","画像タイトル１６","画像タイトル１７","画像タイトル１８","画像タイトル１９",
				"画像タイトル２０","画像タイトル２１","画像タイトル２２","画像タイトル２３","画像タイトル２４","画像タイトル２５","周辺環境画像名称３",
				"周辺環境画像名称４","周辺環境画像名称５","周辺環境画像名称６","周辺環境画像名称７","周辺環境画像名称８","周辺環境画像名称９",
				"周辺環境画像名称１０","周辺環境画像名称１１","周辺環境画像名称１２","周辺環境画像名称１３","周辺環境画像名称１４",
				"周辺環境画像名称１５","周辺環境画像名称１６","周辺環境画像名称１７","周辺環境画像名称１８","周辺環境画像名称１９",
				"周辺環境画像名称２０","周辺環境画像名称２１","周辺環境画像名称２２","周辺環境画像名称２３","周辺環境画像名称２４",
				"周辺環境画像名称２５","周辺環境画像距離３","周辺環境画像距離４","周辺環境画像距離５","周辺環境画像距離６","周辺環境画像距離７",
				"周辺環境画像距離８","周辺環境画像距離９","周辺環境画像距離１０","周辺環境画像距離１１","周辺環境画像距離１２",
				"周辺環境画像距離１３","周辺環境画像距離１４","周辺環境画像距離１５","周辺環境画像距離１６","周辺環境画像距離１７",
				"周辺環境画像距離１８","周辺環境画像距離１９","周辺環境画像距離２０","周辺環境画像距離２１","周辺環境画像距離２２",
				"周辺環境画像距離２３","周辺環境画像距離２４","周辺環境画像距離２５","動画公開","動画URL","21Search公開","汎用フラグ",
				"予定利回り（表面）","予定利回り（実質）","予定年間収入","PDF","用途地域2","都市計画2","その他費用"
		};
	}


	public static void main(String[] args){
		DownloadAttachment da = new DownloadAttachment();
		da.setVisible(true);
		da.show();
	}
	/**
	 * This method initializes this
	 *
	 * @return void
	 */
	private void initialize() {
		try{
			this.setSize(720, 568);
			Toolkit toolkit = Toolkit.getDefaultToolkit();
			Dimension screenSize = toolkit.getScreenSize();
			double screenWith = screenSize.getWidth();
			double screenHeight = screenSize.getHeight();
			int x = (int) ((screenWith - getWidth()) / 2);
			int y = (int) ((screenHeight - getHeight()) / 2);
			this.setLocation(x, y);
			this.setContentPane(getJContentPane());
			this.setTitle("HP用物件データダウンロード");
			Image img = toolkit.getImage(getClass().getResource("/img/las.png"));
			Image bigImage = img.getScaledInstance(20,20,Image.SCALE_AREA_AVERAGING);
			this.setIconImage(bigImage);
			pnlDownload.setVisible(false);

			initialDirectory();
		}catch(Exception ex){
			blSuccess = false;
			log(ex.toString());
		}
	}

	private boolean initialDirectory(){
		boolean blRtn = false;
		try{
			//FileSystemView fsv = FileSystemView.getFileSystemView();
			//java.io.File f = fsv.getHomeDirectory();
			String currDir = System.getProperty("user.dir");
			hp_path = currDir + hp_path;
			work_path = currDir + work_path;
			log_path = currDir + log_path;
			java.io.File d = new java.io.File(hp_path);
			if(d.isDirectory() == false){
				d.mkdirs();
			}
			d = new java.io.File(work_path);
			if(d.isDirectory() == false){
				d.mkdirs();
			}
			d = new java.io.File(log_path);
			if(d.isDirectory() == false){
				d.mkdirs();
			}
			blRtn = true;
		}catch(Exception ex){
			blSuccess = false;
			log(ex.toString());
		}
		return blRtn;
	}
	/**
	 * This method initializes jContentPane
	 *
	 * @return javax.swing.JPanel
	 */
	private JPanel getJContentPane() {
		if (jContentPane == null) {
			lblPassword = new JLabel();
			lblPassword.setText("パスワード");
			lblPassword.setBounds(new Rectangle(74, 70, 93, 24));
			lblUserName = new JLabel();
			lblUserName.setText("ユーザID");
			lblUserName.setBounds(new Rectangle(74, 20, 91, 23));
			jContentPane = new JPanel();
			jContentPane.setLayout(null);
			jContentPane.add(getPnlLogin(), null);
			jContentPane.add(getPnlDownload(), null);
			jContentPane.add(getSpnMessage(), null);
			jContentPane.add(getBtnCopy(), null);
			jContentPane.add(getBtnClose(), null);
		}
		return jContentPane;
	}

	/**
	 * This method initializes txtUserName
	 *
	 * @return javax.swing.JTextField
	 */
	private JTextField getTxtUserName() {
		if (txtUserName == null) {
			txtUserName = new JTextField();
			txtUserName.setBounds(new Rectangle(211, 14, 255, 30));
		}
		return txtUserName;
	}

	/**
	 * This method initializes txtPassword
	 *
	 * @return javax.swing.JPasswordField
	 */
	private JPasswordField getTxtPassword() {
		if (txtPassword == null) {
			txtPassword = new JPasswordField();
			txtPassword.setBounds(new Rectangle(212, 68, 255, 31));
		}
		return txtPassword;
	}

	/**
	 * This method initializes btnLogin
	 *
	 * @return javax.swing.JButton
	 */
	private JButton getBtnLogin() {
		if (btnLogin == null) {
			btnLogin = new JButton();
			btnLogin.setText("ログイン");
			btnLogin.setBounds(new Rectangle(212, 139, 110, 34));
			btnLogin.addActionListener(new java.awt.event.ActionListener() {
				public void actionPerformed(java.awt.event.ActionEvent e) {
					if(login()){
						pnlLogin.setVisible(false);
						pnlDownload.setVisible(true);
					}
				}
			});
		}
		return btnLogin;
	}

	/**
	 * This method initializes btnLogin
	 *
	 * @return javax.swing.JButton
	 */
	private JButton getBtnCopy() {
		if (btnCopy == null) {
			btnCopy = new JButton();
			btnCopy.setIcon(new ImageIcon(getClass().getResource("/img/copy.png")));
			btnCopy.setToolTipText("メッセージコピー");
			btnCopy.setBounds(new Rectangle(610, 293, 30, 31));
			btnCopy.addActionListener(new java.awt.event.ActionListener() {
				public void actionPerformed(java.awt.event.ActionEvent e) {
					 String strSelect = tarMessage.getSelectedText();
					 if(strSelect == null || strSelect.length() ==0){
						 strSelect = tarMessage.getText();
					 }
					 StringSelection data = new StringSelection(strSelect);
					 Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
					 clipboard.setContents(data, data);
				}
			});
		}
		return btnCopy;
	}
	/**
	 * This method initializes tarMessage
	 *
	 * @return javax.swing.JTextArea
	 */
	private JTextArea getTarMessage() {
		if (tarMessage == null) {
			tarMessage = new JTextArea();
			tarMessage.setMargin(new Insets(5, 5, 5, 5));
			tarMessage.setEditable(false);
		}
		return tarMessage;
	}

	private void log(String strMsg){
		String strTimestamp = logTimeFormat.format(new Date());
		if(strMsg != null && strMsg.trim().length() > 0){
			strMsg = " " + strTimestamp + "| " + strMsg + "\n";
			tarMessage.append(strMsg);
			tarMessage.setCaretPosition(tarMessage.getDocument().getLength());
			log.info(strMsg);

		}
	}


	private boolean login(){
		boolean blRtn = false;
		try{
			if(proxyHost != null && proxyHost.trim().length() > 0 &&
					proxyPort != null && proxyPort.trim().length() > 0){
				System.setProperty("http.proxyHost", proxyHost);
				System.setProperty("http.proxyPort", proxyPort);
			}

			String strUserName = txtUserName.getText();
			char[] chrPswd = txtPassword.getPassword();
			if(strUserName == null || strUserName.trim().length() < 1){
				log("ユーザIDを入力してください");
				return blRtn;
			}

			if(chrPswd == null || chrPswd.length < 1){
				log("パスワードを入力してください");
				return blRtn;
			}

			String strPassword = new String(txtPassword.getPassword());
			SforceServiceLocator ssl = new SforceServiceLocator();

			if(strUserName.endsWith(".test")){
				//SandBox
				ssl.setSoapEndpointAddress(Soap_Address_Sandbox);
			}else{
				//本番
				ssl.setSoapEndpointAddress(Soap_Address_Product);
			}
/*
			if(strUserName.endsWith(".com") || strUserName.endsWith(".jp") ||
					strUserName.endsWith(".co.jp")){
				ssl.setSoapEndpointAddress(Soap_Address_Product);
			}else{
				ssl.setSoapEndpointAddress(Soap_Address_Sandbox);
			}
*/
			binding = (SoapBindingStub) ssl.getSoap();
			binding.setTimeout(60000);
			LoginResult loginResult = binding.login(strUserName, strPassword);
			binding._setProperty(SoapBindingStub.ENDPOINT_ADDRESS_PROPERTY, loginResult.getServerUrl());

			SessionHeader sh = new SessionHeader();
			sh.setSessionId(loginResult.getSessionId());
			binding.setHeader(new SforceServiceLocator().getServiceName().getNamespaceURI(), "SessionHeader", sh);
			//log("ログインしました");
			blRtn = true;
		}catch(Exception ex){
			log(ex.toString());
		}
		return blRtn;
	}

	private void clearRow(){
		csvRow = "";
	}

	private void addColVal(String val){
		if(val == null){
			val = "";
		}
		csvRow +=  "\"" + val + "\",";
	}

	private void openCsvFile(){
		try{
			if(fos == null){
				java.io.File d = new java.io.File(hp_path);
				if(d.isDirectory() == false){
					d.mkdirs();
				}
				csvFileName = csvFile + "_" + getTimestamp()+ ".csv";
				String filePath = hp_path + "\\" + csvFileName;
				fos = new FileOutputStream(filePath);
				writer = new OutputStreamWriter(fos,"Windows-31J");
			}
		}catch (Exception e){
			blSuccess = false;
			log(e.toString());
		}
	}

	private void closeCsvFile(){
		if(fos != null){
			try{
				writer.flush();
				writer.close();
				fos.flush();
				fos.close();
				writer = null;
				fos = null;
			}catch (IOException e){
				blSuccess = false;
				log(e.toString());
			}
		}
	}

	private  void createFile(String fileName, byte[] body){
		try{
			String path = work_path;
			java.io.File d = new java.io.File(path);
			if(d.isDirectory() == false){
				d.mkdirs();
			}
			String filePath = path + "\\" + fileName;
			FileOutputStream fos = new FileOutputStream(filePath);
			fos.write(body);
			fos.close();
			//log(fileName + "をダウンロードしました");
		}catch (IOException e){
			blSuccess = false;
			log(e.toString());
		}
	}

	/**
	 * This method initializes pnlLogin
	 *
	 * @return javax.swing.JPanel
	 */
	private JPanel getPnlLogin() {
		if (pnlLogin == null) {
			pnlLogin = new JPanel();
			pnlLogin.setLayout(null);
			pnlLogin.setBounds(new Rectangle(82, 61, 558, 175));
			pnlLogin.add(getTxtUserName(), null);
			pnlLogin.add(lblUserName, null);
			pnlLogin.add(getTxtPassword(), null);
			pnlLogin.add(lblPassword, null);
			pnlLogin.add(getBtnLogin(),null);
		}
		return pnlLogin;
	}

	/**
	 * This method initializes pnlDownload
	 *
	 * @return javax.swing.JPanel
	 */
	private JPanel getPnlDownload() {
		if (pnlDownload == null) {
			lblStep = new JLabel();
			lblStep.setBounds(new Rectangle(9, 180, 103, 27));
			lblStep.setDisplayedMnemonic(KeyEvent.VK_UNDEFINED);
			lblStep.setText("CSVファイル作成");
			lblAll = new JLabel();
			lblAll.setBounds(new Rectangle(9, 143, 75, 25));
			lblAll.setText("全体");
			lblDays = new JLabel();
			lblDays.setBounds(new Rectangle(120, 28, 128, 26));
			lblDays.setText("公開限定期間");
			pnlDownload = new JPanel();
			pnlDownload.setLayout(null);
			pnlDownload.setEnabled(true);
			pnlDownload.setBounds(new Rectangle(88, 58, 547, 226));
			pnlDownload.add(lblDays, null);
			pnlDownload.add(getBtnDownload(), null);
			pnlDownload.add(getPrgbDownload(), null);
			pnlDownload.add(getCbbDays(), null);
			pnlDownload.add(getPrgbAllSteps(), null);
			pnlDownload.add(lblAll, null);
			pnlDownload.add(lblStep, null);
		}
		return pnlDownload;
	}

	/**
	 * This method initializes btnDownload
	 *
	 * @return javax.swing.JButton
	 */
	private JButton getBtnDownload() {
		if (btnDownload == null) {
			btnDownload = new JButton();
			btnDownload.setBounds(new Rectangle(223, 78, 115, 33));
			btnDownload.setText("ダウンロード");
			btnDownload.addActionListener(new java.awt.event.ActionListener() {
				public void actionPerformed(java.awt.event.ActionEvent e) {
					try{
						blSuccess = true;
						tarMessage.setText("");
						File d = new java.io.File(work_path);
						if(d.isDirectory() == false){
							d.mkdirs();
						}else{
							deleteDirectory(work_path);
							//d.mkdirs();
						}

						Object objDays = cbbDays.getSelectedItem();
						if(objDays ==null){
							String msg = "公開限定期間を選択してください";
							tarMessage.append(msg);
							tarMessage.setCaretPosition(tarMessage.getDocument().getLength());
							log.info(msg);
							return;
						}

					    btnDownload.setEnabled(false);
					    prgbAllSteps.setMinimum(0);
					    Task task = new Task();
					    task.addPropertyChangeListener(new java.beans.PropertyChangeListener() {
							public void propertyChange(PropertyChangeEvent evt)  {
								if ("progress" == evt.getPropertyName()) {
								      int progress = (Integer) evt.getNewValue();
								      prgbDownload.setIndeterminate(false);
								      prgbDownload.setValue(progress);
								      /* String msg = String.format(" 実行： %d%%", progress);
								      msg = " " + logTimeFormat.format(new Date()) + "| " + msg + "\n";
								      tarMessage.append(msg);
								      tarMessage.setCaretPosition(tarMessage.getDocument().getLength());
								      log.info(msg); */
							    }
							}
						});
					    task.execute();
					}catch(Exception ex){
						blSuccess = false;
						log(ex.toString());
					}
				}
			});
		}
		return btnDownload;
	}

	private JButton getBtnClose(){
		if(btnClose == null){
			btnClose = new JButton();
			btnClose.setText("終了");
			btnClose.setBounds(new Rectangle(621, 7, 87, 33));
			btnClose.addActionListener(new java.awt.event.ActionListener() {
				public void actionPerformed(java.awt.event.ActionEvent e) {
					 System.exit(0);
				}
			});
		}
		return btnClose;
	}

	/**
	 * This method initializes prgbDownload
	 *
	 * @return javax.swing.JProgressBar
	 */
	private JProgressBar getPrgbDownload() {
		if (prgbDownload == null) {
			prgbDownload = new JProgressBar(0, 100);
			prgbDownload.setIndeterminate(false);
			prgbDownload.setBounds(new Rectangle(121, 181, 425, 27));
			prgbDownload.setComponentOrientation(ComponentOrientation.LEFT_TO_RIGHT);
			prgbDownload.setStringPainted(true);
		}
		return prgbDownload;
	}

	/**
	 * This method initializes prgbAllSteps
	 *
	 * @return javax.swing.JProgressBar
	 */
	private JProgressBar getPrgbAllSteps() {
		if (prgbAllSteps == null) {
			prgbAllSteps = new JProgressBar(0,100);
			prgbAllSteps.setBounds(new Rectangle(121, 141, 424, 26));
			prgbAllSteps.setStringPainted(true);
		}
		return prgbAllSteps;
	}

	/**
	 * This method initializes spnMessage
	 *
	 * @return javax.swing.JScrollPane
	 */
	private JScrollPane getSpnMessage() {
		if (spnMessage == null) {
			spnMessage = new JScrollPane(getTarMessage());
			spnMessage.setBounds(new Rectangle(83, 325, 558, 185));
			spnMessage.setViewportView(getTarMessage());
		}
		return spnMessage;
	}

	private String getIntVal(String strVal){
		String strRtn = "";
		if(strVal != null && strVal.trim().length() > 0){
			if(strVal.indexOf(".") > 0){
				strRtn = strVal.substring(0, strVal.indexOf("."));
			}else{
				strRtn = strVal;
			}
		}
		return strRtn;
	}

	private String getPrevVal(String pVal,int fraction){
		String strVal = "";
		if(pVal != null && pVal.trim().length() > 0){
			try{
				Double dVal = Double.valueOf(pVal);
				String sVal = String.valueOf(dVal);
				if(sVal != null && sVal.trim().length() > 0){
					if(sVal.indexOf(".") > 0){
						if(fraction == 0){
							strVal = sVal.substring(0,sVal.indexOf("."));
						}else if(fraction == 1){
							strVal = sVal.substring(0,sVal.indexOf(".")+ 2);
						}
					}else{
						if(fraction == 0){
							strVal = sVal;
						}else if(fraction == 1){
							strVal = sVal + ".0";
						}
					}
				}
			}catch(Exception ex){
				ex.printStackTrace();
			}
		}
		return strVal;
	}

	private String getTimestamp(){
		String strRtn = "";
		strRtn = timeFormat.format(new Date());
		return strRtn;
	}

	private void deleteDirectory(String filepath) throws IOException{
		try{
			File f = new File(filepath);
			if(f.exists() && f.isDirectory()){
				if(f.listFiles().length==0){
					//f.delete();
			    }else{
			    	 File delFile[]=f.listFiles();
				     int i =f.listFiles().length;
				     for(int j=0;j<i;j++){
				         if(delFile[j].isDirectory()){
				        	 deleteDirectory(delFile[j].getAbsolutePath());
				         }
				         delFile[j].delete();
				     }
				     //f.delete();
			    }
			}
		}catch(Exception e){
			blSuccess = false;
			log(e.toString());
		}
	}

	private void getProperties(){
		Properties prop = new Properties();
		FileInputStream inputStream = null;
		try {
			inputStream = new FileInputStream(propertyFile);
			prop.load(inputStream);
			inputStream.close();
		} catch (IOException e) {
			log(e.toString());
			return;
		}
		proxyHost = prop.getProperty("proxyHost");
		proxyPort = prop.getProperty("proxyPort");
		csvFile  = prop.getProperty("csvFile");
		attachFile = prop.getProperty("attachFile");
		duration = prop.getProperty("duration");
		saveFolder = prop.getProperty("saveFolder");
		Headquarters = prop.getProperty("Headquarters");
		//log("saveFolder111="+saveFolder);
		try{
		saveFolder = new String(saveFolder.getBytes("ISO-8859-1"),"UTF-8");
		System.out.println(saveFolder);
		//log("saveFolder222="+saveFolder);
		}catch(Exception ex){
			log(ex.toString());
		}
		if(duration == null || duration.trim().length() == 0){
			durationDays = 7;
		}else{
			try{
				durationDays = Integer.valueOf(duration);
			}catch(Exception ex){
				log(ex.toString());
			}
		}
	}

	/**
	 * This method initializes cbbDays
	 *
	 * @return javax.swing.JComboBox
	 */
	private JComboBox getCbbDays() {
		if (cbbDays == null) {
			cbbDays = new JComboBox();
			cbbDays.setBounds(new Rectangle(340, 29, 98, 26));
			cbbDays.addItem("１週間");
			cbbDays.addItem("２週間");
			cbbDays.addItem("３週間");
			cbbDays.addItem("４週間");
		}
		return cbbDays;
	}

	class Task extends SwingWorker<Void, Void> {
	    @Override
	    public Void doInBackground() {
	      Random random = new Random();
	      int progress = 0;
	      setProgress(0);

		  log("CSVファイル作成開始");
	      lblStep.setText("CSVファイル作成");

	      String strParentIds = "";
	      HashMap<String,String> checkMap = new HashMap<String,String>();
	      HashMap<String,HashMap> convMap = new HashMap<String, HashMap>();
	      HashMap tempMap = new HashMap();
	      int recordCount = 0;
	      int recordIndex = 0;
	      int fileCount = 0;
	      int downFileCount = 0;
	      boolean blImg = false;
	      boolean blPdf = false;
	      Object objDays = (Object)cbbDays.getSelectedItem();
	      int intDays = 0;
	      if(objDays != null){
	    	  String strDays = (String)objDays;
	    	  if(strDays.equals("１週間")){
	    		  intDays = 7;
	    	  }else if(strDays.equals("２週間")){
	    		  intDays = 14;
	    	  }else if(strDays.equals("３週間")){
	    		  intDays = 21;
	    	  }else if(strDays.equals("４週間")){
	    		  intDays = 28;
	    	  }
	      }
	      try {
	    	  QueryResult queryResult = binding.query("select Id,owner_company__c,Name, "
						+ " LAS_bukken__r.Prefecture__c,LAS_bukken__r.address_city__c,LAS_bukken__r.address_city_rest__c, "
						+ " LAS_bukken__r.address_chome__c,LAS_bukken__r.address_chiban__c,LAS_bukken__r.Name, "
						+ " LAS_bukken__r.traffic1_line__c,LAS_bukken__r.traffic1_station__c,LAS_bukken__r.traffic1_walking__c, "
						+ " LAS_bukken__r.traffic1_bus_stop__c,LAS_bukken__r.traffic1_bus_minutes__c,LAS_bukken__r.traffic1_bus_stop_minutes__c, "
						+ " LAS_bukken__r.traffic2_line__c,LAS_bukken__r.traffic2_station__c,LAS_bukken__r.traffic2_walking__c, "
						+ " LAS_bukken__r.traffic2_bus_stop__c,LAS_bukken__r.traffic2_bus_minutes__c,LAS_bukken__r.traffic2_bus_stop_minutes__c, "
						+ " SellPrice__c,land_rent__c,ManagementCost__c,RepairReserve__c,LAS_bukken__r.UseDistrict__c, "
						+ " LAS_bukken__r.LandCertificate__c,LAS_bukken__r.Structure__c,LAS_bukken__r.Kaidate__c, "
						+ " LAS_bukken__r.PersonalArea__c,LAS_bukken__r.PersonalAreaTsubo__c,LAS_bukken__r.balconySpace__c, "
						+ " LAS_bukken__r.RoomAmount__c,LAS_bukken__r.FloorLocation__c,LAS_bukken__r.RoomNo__c,LAS_bukken__r.bulding_direction__c, "
						+ " LAS_bukken__r.BuiltYear__c,LAS_bukken__r.BuiltMonth__c,PresentState__c,ManagerPresence__c,LAS_bukken__r.RoomType__c, "
						+ " LAS_bukken__r.Parking_state__c,DeliverState__c,DeliverDate__c,misoku_up__c,sell_offer_date__c,status__c,ManagementType__c, "
						+ " LAS_bukken__r.Latitude__c,LAS_bukken__r.Longitude__c,LAS_bukken__r.other_structure__c, "
						+ " GrossRate__c,NetRate__c,annual_rent__c,CompanyNote__c,rental_period__c,miscCost__c, "
						+ " (select Id,ParentId,Name,BodyLength from Attachments where name like 'HP\\_%' order by name) "
						+ " from LAS_Project__c "
						+ " where misoku_up__c != null "
						+ " and sell_offer_date__c = null "
						+ " and status__c in ('販売中（空室）', '販売中（賃貸中）') "
						+ " and Headquarters__c = '" + Headquarters + "'");
						//+ " and (misoku_up__c = LAST_N_DAYS:" + intDays + " or status__c in ('販売中（空室）', '販売中（賃貸中）'))");
				if(queryResult != null  && queryResult.getSize() > 0){
					recordCount = queryResult.getSize();
					openCsvFile();
					QueryResult convResult = binding.query("select ItemNo__c,BeforeValue__c,AfterValue__c "
														+ " from LASConvertHP__c order by ItemNo__c");
					if(convResult != null && convResult.getSize() > 0){
						SObject[] convRecords = convResult.getRecords();
						for(int i=0; i<convRecords.length; i++){
							MessageElement[] elements = convRecords[i].get_any();
							String itemNo = elements[0].getValue();
							if(itemNo.indexOf(".") > 0){
								itemNo = itemNo.substring(0,itemNo.indexOf("0")-1);
							}
							String bVal = elements[1].getValue();
							String aVal = elements[2].getValue();
							if(!convMap.containsKey(itemNo)){
								HashMap<String,String> valMap = new HashMap<String,String>();
								valMap.put(bVal, aVal);
								convMap.put(itemNo, valMap);
							}else{
								HashMap<String,String> valMap = (HashMap<String,String>)convMap.get(itemNo);
								valMap.put(bVal,aVal);
							}
						}
						clearRow();
						for(Integer i=0;i<titles.length;i++){
							addColVal(titles[i]);
						}
						if(csvRow != null && csvRow.length() > 0 && csvRow.endsWith(",")){
							csvRow = csvRow.substring(0,csvRow.length()-1);
						}
						writer.write(csvRow);
						writer.write("\r\n");
					}
				}

				if(recordCount > 0){
					boolean queryMore;
					do {
						queryMore = false;
						if (queryResult != null && queryResult.getSize() != 0) {
							SObject[] records = queryResult.getRecords();
							SObject soParent;

							for (int i = 0; i < records.length; i++) {
								MessageElement[] elements = records[i].get_any();
								String sProjId =  elements[0].getValue();
								String strVal = "";
								String strTemp = "";
								Boolean isOtherStructure = false;
								Boolean hasNoLine = false;
								clearRow();

								//owner_company__c
								strVal = elements[1].getValue();
								tempMap = convMap.get("1");
								if(tempMap != null && tempMap.containsKey(strVal)){
									strVal = (String)tempMap.get(strVal);
								}
								/*if(strVal.equals("LAS") || strVal.equals("MRE")){
									strVal = "10-";
								}else if(strVal.equals("LAS北海道")){
									strVal = "20-";
								}else if(strVal.equals("LAS中部")){
									strVal = "30-";
								}else if(strVal.equals("LAS近畿")){
									strVal = "40-";
								}else if(strVal.equals("LAS京都")){
									strVal = "50-";
								}else if(strVal.equals("LAS九州")){
									strVal = "60-";
								}else if(strVal.equals("LAS沖縄")){
									strVal = "70-";
								}*/

								strTemp = strVal;
								addColVal(strVal);
								strVal = elements[2].getValue()==null?"":elements[2].getValue().substring(4);
								addColVal(strVal);
								addColVal(strTemp + strVal);

								Date dt = new Date();
								DateFormat formatter = new SimpleDateFormat("yyyy/MM/dd hh:mm:ss");
								strVal = formatter.format(dt);
								addColVal(strVal);
								addColVal(strVal);
								addColVal(strVal);

								addColVal("居住用");
								addColVal("売マンション");
								addColVal("中古マンション");
								addColVal("");

								// LAS_bukken__r.Prefecture__c,LAS_bukken__r.address_city__c,LAS_bukken__r.address_city_rest__c,
								// LAS_bukken__r.address_chome__c,LAS_bukken__r.address_chiban__c, LAS_bukken__r.Name,
								// LAS_bukken__r.RoomNo__c
								soParent = (SObject)elements[3].getObjectValue();
								MessageElement[] subElements = soParent.get_any();
								addColVal(subElements[0].getValue());
								addColVal(subElements[1].getValue());
								addColVal(subElements[2].getValue());
								addColVal(subElements[3].getValue());
								addColVal(subElements[4].getValue());

								strTemp = subElements[27].getValue();
								if(strTemp != null && strTemp.trim().length() > 0){
									addColVal(subElements[5].getValue() + "_" + strTemp +  "号室");
								}else{
									addColVal(subElements[5].getValue());
								}

								for(int m=0;m<7;m++){
									addColVal("");
								}

								// LAS_bukken__r.traffic1_line__c,LAS_bukken__r.traffic1_station__c,LAS_bukken__r.traffic1_walking__c,
								// LAS_bukken__r.traffic1_bus_stop__c,LAS_bukken__r.traffic1_bus_minutes__c,LAS_bukken__r.traffic1_bus_stop_minutes__c,
								if(subElements[6].getValue() == null || subElements[6].getValue().trim().length() ==0){
									hasNoLine = true;
								}
								if(hasNoLine){
									addColVal("");
									addColVal("");
									addColVal("沿線なし");
									addColVal("");
									addColVal("");
									addColVal("");
									addColVal("");
									addColVal(getIntVal(subElements[11].getValue()));
								}else{
									addColVal(subElements[6].getValue());
									addColVal(subElements[7].getValue());
									addColVal("");
									addColVal(getIntVal(subElements[8].getValue()));
									addColVal(subElements[9].getValue());
									addColVal(getIntVal(subElements[10].getValue()));
									addColVal(getIntVal(subElements[11].getValue()));
									addColVal("");
								}
								addColVal("");

								// LAS_bukken__r.traffic2_line__c,LAS_bukken__r.traffic2_station__c,LAS_bukken__r.traffic2_walking__c,
								// LAS_bukken__r.traffic2_bus_stop__c,LAS_bukken__r.traffic2_bus_minutes__c,LAS_bukken__r.traffic2_bus_stop_minutes__c,
								if(hasNoLine){
									for(int m=0;m<8;m++){
										addColVal("");
									}
								}else{
									addColVal(subElements[12].getValue());
									addColVal(subElements[13].getValue());
									addColVal(getIntVal(subElements[14].getValue()));
									addColVal(subElements[15].getValue());
									addColVal(getIntVal(subElements[16].getValue()));
									addColVal(getIntVal(subElements[17].getValue()));
									addColVal("");
									addColVal("");
								}
								// SellPrice__c
								strVal = elements[4].getValue();
								if(strVal != null){
									Double price = Double.valueOf(strVal);
									price = Math.floor(price / 10000);
									strVal = ""+price;
								}
								addColVal(strVal);

								addColVal("税込");
								addColVal("");

								// land_rent__c,ManagementCost__c
								addColVal(getIntVal(elements[5].getValue()));
								strVal = getIntVal(elements[6].getValue());
								addColVal(strVal);
								if(strVal == null || strVal.trim().length() < 1 || strVal.equals("0")){
									strVal = "無";
								}else{
									strVal = "";
								}
								addColVal(strVal);

								// RepairReserve__c
								strVal = getIntVal(elements[7].getValue());
								addColVal(strVal);
								if(strVal == null || strVal.trim().length() < 1 || strVal.equals("0")){
									strVal = "無";
								}else{
									strVal = "";
								}
								addColVal(strVal);

								addColVal("");
								addColVal("");
								addColVal("");

								// LAS_bukken__r.UseDistrict__c
								strVal = subElements[18].getValue();
								if(strVal != null && strVal.trim().length() > 0){
									String[] items = strVal.split(";");
									if(items != null && items.length > 0){
										strVal = items[0];
										tempMap = convMap.get("52");
										if(tempMap != null && tempMap.containsKey(strVal)){
											strVal = (String)tempMap.get(strVal);
										}
										/*if(strVal.equals("第1種低層住居専用地域")){
											strVal = "第一種低層住居専用地域";
										}else if(strVal.equals("第2種低層住居専用地域")){
											strVal = "第二種低層住居専用地域";
										}else if(strVal.equals("第1種中高層住居専用地域")){
											strVal = "第一種中高層住居専用地域";
										}else if(strVal.equals("第2種中高層住居専用地域")){
											strVal = "第二種中高層住居専用地域";
										}else if(strVal.equals("第1種住居地域")){
											strVal = "第一種住居地域";
										}else if(strVal.equals("第2種住居地域")){
											strVal = "第二種住居地域";
										}*/
									}else{
										strVal = "";
									}
								}else{
									strVal = "";
								}
								addColVal(strVal);

								for(int m=0;m<19;m++){
									addColVal("");
								}

								// LAS_bukken__r.LandCertificate__c
								strVal = subElements[19].getValue();
								if(strVal != null){
									tempMap = convMap.get("72");
									if(tempMap != null && tempMap.containsKey(strVal)){
										strVal = (String)tempMap.get(strVal);
									}
									/*
									if(strVal.equals("所有権（敷地権）")){
										strVal = "所有権の共有";
									}else if(strVal.equals("地上権")){
										strVal = "（旧法）地上権";
									}else if(strVal.equals("賃借権")){
										strVal = "普通借地権（賃借権）";
									}else if(strVal.equals("所有権･賃借権")){
										strVal = "所有権一部賃借権";
									}else if(strVal.equals("所有権･地上権")){
										strVal = "所有権一部地上権";
									}*/
								}
								addColVal(strVal);

								for(int m=0;m<15;m++){
									addColVal("");
								}
								strVal = getIntVal(elements[20].getValue());
								addColVal(strVal);
								for(int m=0;m<6;m++){
									addColVal("");
								}
								// LAS_bukken__r.Structure__c
								strVal = subElements[20].getValue();
								System.out.println("strVal="+strVal);
								if(strVal != null){
									tempMap = convMap.get("95");
									if(tempMap != null && tempMap.containsKey(strVal)){
										strVal = (String)tempMap.get(strVal);
									}
									/*if(strVal.equals("RC造")){
										strVal = "鉄筋コンクリート造";
									}else if(strVal.equals("SRC/RC造")){
										strVal = "鉄骨鉄筋コンクリート造";
									}else if(strVal.equals("SRC造")){
										strVal = "鉄骨鉄筋コンクリート造";
									}else if(strVal.equals("S造")){
										strVal = "鉄骨造";
									}*/
								}
								if(strVal.equals("その他")){
									isOtherStructure = true;
								}
								addColVal(strVal);

								//LAS_bukken__r.Kaidate__c
								strVal = getIntVal(subElements[21].getValue());
								Integer intFloor = null;
								try{
									intFloor = Integer.parseInt(strVal);
								}catch(Exception exc){
								}
								if(intFloor == null){
									addColVal("");
								}else{
									strVal = String.valueOf(intFloor.intValue());
									addColVal(strVal);
								}

								for(int m=0;m<5;m++){
									addColVal("");
								}

								//LAS_bukken__r.PersonalArea__c,LAS_bukken__r.PersonalAreaTsubo__c,LAS_bukken__r.balconySpace__c,LAS_bukken__r.RoomAmount__c,
								// LAS_bukken__r.FloorLocation__c,LAS_bukken__r.RoomNo__c,LAS_bukken__r.bulding_direction__c
								addColVal(subElements[22].getValue());
								addColVal(subElements[23].getValue());
								addColVal(subElements[24].getValue());
								addColVal("");
								addColVal(getIntVal(subElements[25].getValue()));
								addColVal("");
								addColVal("");
								addColVal(subElements[26].getValue());
								addColVal(subElements[27].getValue());
								addColVal(subElements[28].getValue());

								//LAS_bukken__r.BuiltYear__c,LAS_bukken__r.BuiltMonth__c
								strTemp = subElements[29].getValue();
								strVal = subElements[30].getValue();
								if(strVal != null && strVal.indexOf("月")>0){
									strVal = strVal.substring(0,strVal.indexOf("月"));
									if(strVal != null && strVal.trim().length() < 2){
										strVal = "0"+ strVal + "月";
									}else{
										strVal = strVal + "月";
									}
								}
								addColVal(strTemp + strVal);
								addColVal("");
								addColVal("");

								//PresentState__c
								strVal = elements[8].getValue();
								if(strVal != null){
									tempMap = convMap.get("115");
									if(tempMap != null && tempMap.containsKey(strVal)){
										strVal = (String)tempMap.get(strVal);
									}
									/*
									 (strVal.equals("申込中") || strVal.equals("退却予定"))){
									strVal = "空室";
									*/
								}
								addColVal(strVal);
								addColVal("");
								//ManagerPresence__c
								addColVal(elements[9].getValue());
								addColVal("");
								addColVal("");

								//LAS_bukken__r.RoomType__c
								addColVal(subElements[31].getValue());

								for(int m=0;m<24;m++){
									addColVal("");
								}

								// LAS_bukken__r.Parking_state__c
								strVal = subElements[32].getValue();
								if(strVal != null){
									tempMap = convMap.get("145");
									if(tempMap != null && tempMap.containsKey(strVal)){
										strVal = (String)tempMap.get(strVal);
									}
									/*
									if(strVal.equals("駐車場無し")){
										strVal = "無";
									}else if(strVal.equals("駐車場有り")){
										strVal = "空無";
									}else if(strVal.equals("駐車場有り：空きなし：有料")){
										strVal = "空無";
									}else if(strVal.equals("駐車場有り：空きなし：無料")){
										strVal = "空無";
									}else if(strVal.equals("駐車場有り：空きあり：有料")){
										strVal = "有り有料";
									}else if(strVal.equals("駐車場有り：空きあり：無料")){
										strVal = "有り無料";
									}*/
								}
								addColVal(strVal);

								for(int m=0;m<8;m++){
									addColVal("");
								}

								// owner_company__c
								strVal = elements[1].getValue();
								if(strVal != null){
									tempMap = convMap.get("154");
									if(tempMap != null && tempMap.containsKey(strVal)){
										strVal = (String)tempMap.get(strVal);
									}
								}
								/*if(strVal.equals("MRE") || strVal.equals("LAS")){
									strVal = "売主";
								}else{
									strVal = "代理";
								}*/

								addColVal(strVal);

								for(int m=0;m<7;m++){
									addColVal("");
								}
								// DeliverState__c,DeliverDate__c
								addColVal(elements[10].getValue());
								strVal = elements[11].getValue();
								if(strVal != null){
									tempMap = convMap.get("164");
									if(tempMap != null && tempMap.containsKey(strVal)){
										strVal = (String)tempMap.get(strVal);
									}
								}
								/*if(strVal != null && strVal.equals("即時")){
									strVal = "即引き渡し可";
								}*/
								addColVal(strVal);

								for(int m=0;m<13;m++){
									addColVal("");
								}

								// owner_company__c
								strVal = elements[1].getValue();
								if(strVal != null){
									tempMap = convMap.get("177");
									if(tempMap != null && tempMap.containsKey(strVal)){
										strVal = (String)tempMap.get(strVal);
									}
								}
								/*if(strVal != null && strVal.equals("MRE")){
									strVal = "LAS";
								}*/
								addColVal(strVal);
								addColVal("");

								// owner_company__c
								strVal = elements[1].getValue();
								if(strVal != null){
									tempMap = convMap.get("179");
									if(tempMap != null && tempMap.containsKey(strVal)){
										strVal = (String)tempMap.get(strVal);
									}
								}
								/*if(strVal.equals("LAS")){
									strVal = "03-6703-0250";
								}else if(strVal.equals("LAS北海道")){
									strVal = "011-281-1052";
								}else if(strVal.equals("LAS中部")){
									strVal = "052-533-0680";
								}else if(strVal.equals("LAS近畿")){
									strVal = "06-6530-0250";
								}else if(strVal.equals("LAS京都")){
									strVal = "075-253-6283";
								}else if(strVal.equals("LAS九州")){
									strVal = "092-303-6380";
								}else if(strVal.equals("LAS沖縄")){
									strVal = "098-862-1781";
								}else if(strVal.equals("MRE")){
									strVal = "03-6703-0250";
								}*/
								addColVal(strVal);

								for(int m=0;m<21;m++){
									addColVal("");
								}

								String[] imgFileNames = new String[26];
								// 添付ファイルを取得
								QueryResult attResults = (QueryResult)elements[22].getObjectValue();
								ArrayList<String> lstNameCheck = new ArrayList<String>();
								String strDuplicate = "";
								boolean blHpFile = false;
								boolean bDone = false;
								int idxAtta = 2;
								ArrayList<String> checkList = new ArrayList<String>();

								if(attResults != null && attResults.getSize() > 0){
									while (!bDone) {
						        		SObject[] atts = attResults.getRecords();
						        		for (int m = 0; m < atts.length; m++) {
							            	MessageElement[] att = atts[m].get_any();
							            	String fileName = att[2].getValue();
							            	try{
							            		if(fileName != null && fileName.length() > 0){
							            			String[] sIdxs = fileName.split("_");
							            			if(sIdxs != null && sIdxs[0].equals("HP")){
							            				if(!lstNameCheck.contains(fileName)){
							            					lstNameCheck.add(fileName);
							            				}else{
							            					if(strDuplicate.length() > 0 ){
							            						if(strDuplicate.indexOf(fileName) < 0){
							            							strDuplicate += "、" + fileName;
							            						}
							            					}else{
							            						strDuplicate += fileName;
							            					}
							            				}
							            				if(sIdxs.length > 3){
									            		    String sIndex = sIdxs[2];
									            		    if(sIndex.equalsIgnoreCase("01")){
									            		    	if(!checkList.contains(sIndex)){
									            					checkList.add(sIndex);
									            				}
									            		    	imgFileNames[0] = fileName;
									            		    }else if(sIndex.equalsIgnoreCase("02")){
									            		    	if(!checkList.contains(sIndex)){
									            					checkList.add(sIndex);
									            				}
									            		    	imgFileNames[1] = fileName;
									            		    }else if(sIndex.equalsIgnoreCase("z99")){
									            		    	if(!checkList.contains(sIndex)){
									            					checkList.add(sIndex);
									            				}
									            		    	imgFileNames[25] = fileName;
									            		    }else{
									            		    	int intIndex = idxAtta++;
									            		    	if(intIndex > 24){
									            		    		intIndex = intIndex + 1;
									            		    	}
									            		    	imgFileNames[intIndex] = fileName;
									            		    }
							            				}else{
							            					int intIndex = idxAtta++;
								            		    	if(intIndex > 24){
								            		    		intIndex = intIndex + 1;
								            		    	}
								            		    	imgFileNames[intIndex] = fileName;

							            				}
								            		    blHpFile = true;
								            		    fileCount += 1;

								            		    //---仕様変更によって、以下のコードをコメントする。
								            		    /*String sIndex = sIdxs[2];
								            		    if(sIndex.startsWith("0") || sIndex.startsWith("z")){
								            		    	sIndex = sIndex.substring(1);
								            		    }
								            			Integer idxFile = Integer.valueOf(sIndex);
								            			if(idxFile == 1 || idxFile == 2 || idxFile == 99){
								            				if(!checkList.contains(sIdxs[2])){
								            					checkList.add(sIdxs[2]);
								            				}
								            			}

								            			if(idxFile != null && idxFile > -1){
									            			if(idxFile > 0 && idxFile < 26){
									            				imgFileNames[idxFile] = fileName;
									            			}else if(idxFile == 99){
									            				imgFileNames[25] = fileName;
									            			}
								            			}
								            			if(sIdxs[0].equals("HP")){
								            				blHpFile = true;
								            			}
								            			*/
								            		}
							            		}
							            	}catch(Exception e){
							            		blSuccess = false;
							            		log(e.toString());
							            	}
						        		}

						        		if (attResults.isDone()) {
							                bDone = true;
							            } else {
							            	attResults = binding.queryMore(attResults.getQueryLocator());
							            }
						        	}
									if(strDuplicate.length() > 0){
										strDuplicate = "同じファイルが登録されています(" + strDuplicate + ")";
									}
									String strAlert = "";
									if(checkList != null){
										if(!checkList.contains("01")){
											if(strAlert.length() > 0){
												strAlert = strAlert + "、間取り画像ファイル";
											}else{
												strAlert = "間取り画像ファイル";
											}
										}
										if(!checkList.contains("02")){
											if(strAlert.length() > 0){
												strAlert = strAlert + "、外観画像ファイル";
											}else{
												strAlert = "外観画像ファイル";
											}
										}
										if(!checkList.contains("z99")){
											if(strAlert.length() > 0){
												strAlert = strAlert + "、マイソクファイル";
											}else{
												strAlert = "マイソクファイル";
											}
										}
										if(strAlert.length() > 0){
											strAlert = strAlert + "が登録されていません";
										}

									}
									String strMessage = "";
									if(strDuplicate.length() > 0){
										strMessage = strDuplicate;
									}
									if(strAlert.length() > 0){
										if(strMessage.length() > 0){
											strMessage += "、且つ " + strAlert;
										}else{
											strMessage = strAlert;
										}
									}
									if(strMessage.length() > 0){
										checkMap.put(elements[2].getValue(), strMessage);
									}
									if(blHpFile){
										strParentIds += "'" + sProjId + "',";
									}
						        }else{
						        	checkMap.put(elements[2].getValue(), "間取り画像ファイル、外観画像ファイル、マイソクファイルが登録されていません");
						        }
						        for(Integer n = 0; n < 26; n++){
						        	if(imgFileNames[n] != null && imgFileNames[n].trim().length() < 1){
						        		imgFileNames[n] = "";
						        	}
						        }

						        for(Integer n=0;n<8;n++){
						        	addColVal(imgFileNames[n]);
						        }

						        for(int m=0;m<17;m++){
									addColVal("");
								}

						        //インターネット公開 misoku_up__c,sell_offer_date__c,status__c
						        strTemp = elements[12].getValue();

					        	DateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");
					        	Date dtMisoku = fmt.parse(elements[12].getValue());

					        	Date dtNow = new Date();
					        	String sNow = fmt.format(dtNow);
					        	Date dtTemp = fmt.parse(sNow);

					        	long DAY_IN_MS = 1000 * 60 * 60 * 24;
					        	Date dtLastSevenDay = new Date(dtTemp.getTime() - (intDays * DAY_IN_MS));

					        	if(dtMisoku.getTime() > dtLastSevenDay.getTime()){
					        		addColVal("非公開");
					        	}else {
					        		addColVal("公開");
					        	}

						        for(int m=0;m<10;m++){
									addColVal("");
								}
								if(hasNoLine){
									addColVal(subElements[9].getValue());
								}else{
									addColVal("");
								}
								for(int m=0;m<5;m++){
									addColVal("");
								}
						        //ManagementType__c
						        addColVal(elements[15].getValue());

						        for(int m=0;m<17;m++){
									addColVal("");
								}

						        //LAS_bukken__r.Latitude__c,LAS_bukken__r.Longitude__c
						        addColVal(subElements[33].getValue());
						        addColVal(subElements[34].getValue());

						        for(int m=0;m<7;m++){
									addColVal("");
								}

						        if(imgFileNames[0] != null && imgFileNames[0].trim().length() > 0){
						        	strVal = "間取";
						        }else{
						        	strVal = "";
						        }
						        addColVal(strVal);

						        if(imgFileNames[1] != null && imgFileNames[1].trim().length() > 0){
						        	strVal = "建物外観";
						        }else{
						        	strVal = "";
						        }
						        addColVal(strVal);

						        for(int m=0;m<8;m++){
						        	addColVal("");
						        }
						        if(isOtherStructure){
						        	 addColVal(subElements[35].getValue());
						        }else{
						        	addColVal("");
						        }
						        addColVal("");
						        addColVal("");

						        //CompanyNote__c
						        //strVal = elements[19].getValue();
						        //addColVal(strVal);
						        addColVal("");
						        addColVal("");

						        for(Integer n = 8; n <25; n++){
						        	addColVal(imgFileNames[n]);
						        }
						        for(Integer n = 0; n <92; n++){
						        	addColVal("");
						        }

						        // GrossRate__c,NetRate__c,annual_rent__c
						        addColVal(elements[16].getValue());
						        addColVal(elements[17].getValue());
						        strVal = elements[18].getValue();
								if(strVal != null){
									Double price = Double.valueOf(strVal);
									strVal = String.valueOf(price / 10000);
									strVal = getPrevVal(strVal,1);
								}
								addColVal(strVal);
								//PDF
								addColVal(imgFileNames[25]);
								// LAS_bukken__r.UseDistrict__c
								strVal = subElements[18].getValue();
								if(strVal != null && strVal.trim().length() > 0){
									String[] items = strVal.split(";");
									if(items != null && items.length > 1){
										strVal = items[1];
										tempMap = convMap.get("52");
										if(tempMap != null && tempMap.containsKey(strVal)){
											strVal = (String)tempMap.get(strVal);
										}
										/*
										if(strVal.equals("第1種低層住居専用地域")){
											strVal = "第一種低層住居専用地域";
										}else if(strVal.equals("第2種低層住居専用地域")){
											strVal = "第二種低層住居専用地域";
										}else if(strVal.equals("第1種中高層住居専用地域")){
											strVal = "第一種中高層住居専用地域";
										}else if(strVal.equals("第2種中高層住居専用地域")){
											strVal = "第二種中高層住居専用地域";
										}else if(strVal.equals("第1種住居地域")){
											strVal = "第一種住居地域";
										}else if(strVal.equals("第2種住居地域")){
											strVal = "第二種住居地域";
										}*/
									}else{
										strVal = "";
									}
								}else{
									strVal = "";
								}
								addColVal(strVal);

								addColVal("");
						        addColVal(getIntVal(elements[21].getValue()));

								if(csvRow != null && csvRow.length() > 0 && csvRow.endsWith(",")){
									csvRow = csvRow.substring(0,csvRow.length()-1);
								}
								writer.write(csvRow);
								writer.write("\r\n");

								recordIndex += 1;
								progress = (int)Math.floor((double)recordIndex / (double)recordCount * 100);
								setProgress(progress);
								prgbAllSteps.setValue((int)Math.floor((double)recordIndex / (double)recordCount * 15));
							}
						}
						if (!queryResult.isDone()) {
							queryMore = true;
							queryResult = binding.queryMore(queryResult.getQueryLocator());
						}
					} while (queryMore);

					if(checkMap != null && checkMap.size() > 0){
						Iterator iter = checkMap.entrySet().iterator();
						while (iter.hasNext()) {
						    Map.Entry entry = (Map.Entry) iter.next();
						    String key = (String)entry.getKey();
						    String val = (String)entry.getValue();
						    String strMsg = " " + key + "案件は" + val;
						    log(strMsg);
						}
						//---仕様変更によって、以下4行コードをコメントされる。
						//blSuccess = false;
						//strParentIds = "";
						//closeCsvFile();
						//return null;
					}
					closeCsvFile();
				}
				log("CSVファイル作成完了");

				//---ファイルをダウンロード
				log("添付ファイルダウンロード開始");
				setProgress(0);
			    recordIndex = 0;
			    lblStep.setText("ダウンロード");
			    ArrayList<String> lstActualFiles = new ArrayList<String>();

				if(strParentIds != null && strParentIds.endsWith(",")){
					strParentIds = strParentIds.substring(0,strParentIds.lastIndexOf(","));
				}
				if(strParentIds.length() > 14){
					QueryResult results = binding.query("Select Id, Name,ParentId, body, ContentType "
							             + " From Attachment Where ParentId in("+ strParentIds +")"
		            					 + " and Name like 'HP\\_%' and LastModifiedDate = LAST_N_DAYS:" + durationDays
		            					 + " order by name ");

					if(results != null && results.getSize() > 0){
						recordCount = results.getSize();
						boolean done = false;
						while (!done) {
							SObject[] attaches = results.getRecords();
				            for (int m = 0; m < attaches.length; m++) {
				            	MessageElement[] att = attaches[m].get_any();
				                String name = att[1].getValue();
				                String parentId = att[2].getValue();
				                String body = att[3].getValue();
				                String[] arrNames = name.split("_");

				                if(arrNames != null && arrNames[0].equals("HP")){
				                	createFile(name.toString(),Base64.decodeBase64(body));
			                		downFileCount += 1;
			                		if(!lstActualFiles.contains(name)){
			                			lstActualFiles.add(name);
			                		}
				                }

				                recordIndex += 1;
				                progress = (int)Math.floor((double)downFileCount / (double)fileCount * 100);
								setProgress(progress);
								prgbAllSteps.setValue(15+(int)Math.floor((double)downFileCount / (double)fileCount * 70));
				            }
				            if (results.isDone()) {
				                done = true;
				            } else {
				            	results = binding.queryMore(results.getQueryLocator());
				            }
				        }
					}else{
						setProgress(100);
						prgbAllSteps.setValue(85);
					}
				}
				log("添付ファイルダウンロード完了");

				//----ファイルを圧縮する
				log("ダウンロードファイル圧縮開始");
				setProgress(0);
			    lblStep.setText("ファイル圧縮");
			    int fileIndex = 0;

			    if(lstActualFiles.size() > 0){
			    	downFileCount = lstActualFiles.size();
			    }

				List<String> fileList = new ArrayList<String>();
				FileSystemView fsv = FileSystemView.getFileSystemView();
				java.io.File f = fsv.getHomeDirectory();
				byte[] buffer = new byte[1024];

				String path = work_path;
				java.io.File d = new java.io.File(path);
				String strDate = getTimestamp();
				String zipFileName = attachFile + "_" + strDate  + ".zip";
				String zipFilePath = hp_path + "\\" + zipFileName;
				if(d.isDirectory()){
					FileOutputStream fos = new FileOutputStream(zipFilePath);
		    		ZipOutputStream zos = new ZipOutputStream(fos);
		    		zos.setEncoding("MS932");
		    		File[] lstFiles = d.listFiles();

		    		for(File file : lstFiles){
		    			if(file.isFile()){
			        		ZipEntry ze= new ZipEntry(file.getName());
			            	zos.putNextEntry(ze);
			            	FileInputStream in = new FileInputStream(path + File.separator + file.getName());
			            	int len;
			            	while ((len = in.read(buffer)) > 0) {
			            		zos.write(buffer, 0, len);
			            	}
			            	in.close();
			            	fileIndex += 1;

			                progress = (int)Math.floor((double)fileIndex / (double)downFileCount * 100);
							setProgress(progress);
							prgbAllSteps.setValue(85+(int)Math.floor((double)fileIndex / (double)downFileCount * 7));
		    			}
		        	}
		    		zos.closeEntry();
		        	zos.close();
				}
				if(downFileCount == 0){
	    			setProgress(100);
					prgbAllSteps.setValue(92);
				}

				log("ダウンロードファイル圧縮完了");

				//----ファイルを移行する
				log("生成ファイル移動開始");
				strDate = strDate.substring(0,strDate.indexOf("_"));
				String savePath = saveFolder + "\\" + strDate;
				File fold = new File(savePath);
				if(fold.isDirectory() == false){
					fold.mkdirs();
				}
				File csvFile = new File(hp_path + "\\" + csvFileName);
				File saveCsvFile = new File(savePath + "\\" + csvFileName);
				File zipFile = new File(zipFilePath);
				File saveZipFile = new File (savePath + "\\" + zipFileName);
				if(csvFile.renameTo(saveCsvFile) == false ||
						zipFile.renameTo(saveZipFile) == false){
					log("生成ファイル移動失敗");
					return null;
				}else{
					log("生成ファイル移動完了");
				}
				prgbAllSteps.setValue(98);


				//----ファイルを削除する
				log("ダウンロードファイル削除開始");
				deleteDirectory(work_path);
	    		prgbAllSteps.setValue(100);
	    		log("ダウンロードファイル削除完了");

	    		blSuccess = true;
	      } catch (Exception ex) {
	    	  blSuccess = false;
	    	  log(ex.toString());
	      }

	      return null;
	    }

	    /*
	     * Executed in event dispatch thread
	     */
	    public void done() {
	      Toolkit.getDefaultToolkit().beep();
	      btnDownload.setEnabled(true);
	      if(blSuccess){
	    	  log("全ての処理完了");
	      }else{
	    	  log("問題を解決してから、もう一度ダウンロードしてください");
	      }
		  tarMessage.append(" \n");
		  tarMessage.setCaretPosition(tarMessage.getDocument().getLength());
		  log.info(" \n");
	    }
	  }


	public void p(String strMsg){
		System.out.println(strMsg);
	}

	public static void sp(String strMsg){
		System.out.println(strMsg);
	}
}  //  @jve:decl-index=0:visual-constraint="50,7"
