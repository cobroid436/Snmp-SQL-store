package SNMPExcelToSQL;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jdesktop.swingx.JXDatePicker;
import org.snmp4j.CommunityTarget;
import org.snmp4j.PDU;
import org.snmp4j.Snmp;
import org.snmp4j.TransportMapping;
import org.snmp4j.event.ResponseEvent;
import org.snmp4j.mp.SnmpConstants;
import org.snmp4j.smi.*;
import org.snmp4j.transport.DefaultUdpTransportMapping;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.*;
import java.text.ParseException;
import java.util.Iterator;
import java.util.Vector;

public class HelloSNMPExcelToSQL extends JPanel {

    private static String fileSourceXLSX="";
    private static JXDatePicker picker = new JXDatePicker();
	private static JXDatePicker picker2 = new JXDatePicker();
	private static JFrame frame = new JFrame("Parser");
	private static JButton button2 = new JButton("Start");
	private static JButton button = new JButton("Filter");
	private static JTextField text = new JTextField(17);
    private static String connectionUrl = "";
    private static String oidPages  = "1.3.6.1.2.1.43.10.2.1.4.1.1";
    private static String oidSN  = ".1.3.6.1.2.1.43.5.1.1.17.1";
    //	private static String oidSN2  = ".1.3.6.1.4.1.11.2.3.9.4.2.1.1.3.3.0";
    private static String oidMacX  = ".1.3.6.1.2.1.2.2.1.6.1";
    private static String oidMacHp  = ".1.3.6.1.2.1.2.2.1.6.2";
    private static String oidPlaceHp  = ".1.3.6.1.2.1.1.5.0";
    private static String oidPlaceKyo  = ".1.3.6.1.4.1.1347.40.10.1.1.5.1";

    //OID=.1.3.6.1.2.1.43.5.1.1.16.1, Type=OctetString, Value=TASKalfa 3051ci
    //OID=.1.3.6.1.2.1.25.3.2.1.3.1, Type=OctetString, Value=TASKalfa 3051ci
    //OID=.1.3.6.1.2.1.43.5.1.1.16.1, Type=OctetString, Value=FS-4300DN
    //OID=.1.3.6.1.2.1.25.3.2.1.3.1, Type=OctetString, Value=FS-4300DN
    //OID=.1.3.6.1.2.1.25.3.2.1.3.1, Type=OctetString, Value=HP Officejet Pro X476dw MFP
    //OID=.1.3.6.1.2.1.25.3.2.1.3.1, Type=OctetString, Value=HP LaserJet P3010 Series
    //OID=.1.3.6.1.2.1.43.5.1.1.16.1, Type=OctetString, Value=HP LaserJet Pro MFP M521dn
    //OID=.1.3.6.1.2.1.25.3.2.1.3.1, Type=OctetString, Value=HP LaserJet Pro MFP M521dn
    //OID=.1.3.6.1.2.1.25.3.2.1.3.1, Type=OctetString, Value=Xerox Phaser 3300MFP
    private static String oidModel  = ".1.3.6.1.2.1.25.3.2.1.3.1";

    private static String oidHpBlack = ".1.3.6.1.2.1.43.11.1.1.6.1.1";
    private static String oidHpC = ".1.3.6.1.2.1.43.11.1.1.6.1.2";
    private static String oidHpM = ".1.3.6.1.2.1.43.11.1.1.6.1.3";
    private static String oidHpE = ".1.3.6.1.2.1.43.11.1.1.6.1.4";
    private static String oidHpBlackRem = ".1.3.6.1.2.1.43.11.1.1.9.1.1";
    private static String oidHpCRem = ".1.3.6.1.2.1.43.11.1.1.9.1.2";
    private static String oidHpMRem = ".1.3.6.1.2.1.43.11.1.1.9.1.3";
    private static String oidHpERem = ".1.3.6.1.2.1.43.11.1.1.9.1.4";

	private static int snmpVersion  = SnmpConstants.version1;
	private static String community  = "public";
    public static JProgressBar progressBar;

    public static class Task extends SwingWorker<Void, Void> {
        @Override
        public Void doInBackground() throws IOException, ParseException {
            System.out.println("Work doInBackground");
            java.util.Date date = new java.util.Date();
            Timestamp sqlDate = new java.sql.Timestamp(date.getTime());
            Connection con = null;
            Statement stmt = null;
            ResultSet rs = null;

            try {

                InputStream ExcelFileToRead = new FileInputStream(fileSourceXLSX);
                System.out.printf("ExcelFileToRead %s\n", ExcelFileToRead.available());
                XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
                System.out.printf("wb %s\n", wb.toString());
                XSSFSheet sheet = wb.getSheetAt(0);
                System.out.printf("sheet %s\n", sheet.toString());
                XSSFRow row;
                XSSFCell cell;
                Iterator rows = sheet.rowIterator();
                int lstR = sheet.getLastRowNum();
                System.out.printf("getLastRowNum %d\n", lstR);

                Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
                con = DriverManager.getConnection(connectionUrl);
                System.out.printf("Connection %s\n", con);

                int i = 0;
                String sip = "";
                while (rows.hasNext()) {
                    row = (XSSFRow) rows.next();
                    Iterator cells = row.cellIterator();
                    int pages=0;
                    int rem=0;
                    while (cells.hasNext()) {
                        cell = (XSSFCell) cells.next();
                        System.out.print(cell.getStringCellValue() + " " + i + " \n");
                        if (cell.getColumnIndex() == 0) {
                            i++;
                            if (cell.getStringCellValue() == "" || cell.getStringCellValue().length() < 10) continue;

                            System.out.print(cell.getStringCellValue() + " " + i + " \n");
                            sip = cell.getStringCellValue();
                            String strPages = snmpGet(sip, oidPages);
                            System.out.println("strPages: " + strPages);
                            if (strPages == "") continue;

                            try {
                                pages = Integer.parseInt(strPages);
                             }catch (NumberFormatException e) {
                                System.err.println("Неверный формат строки!");
                            }

                            String strRem = snmpGet(sip, oidHpBlackRem);
                            System.out.println("strPlace: " + strRem);
                            if (strRem == "") strRem = "000000";
                            try {
                                rem = Integer.parseInt(strRem);
                            }catch (NumberFormatException e) {
                                System.err.println("Неверный формат строки!");
                            }

                            String strPlace = snmpGet(sip, oidPlaceHp);
                            if (strPlace.length() < 6) strPlace = snmpGet(sip, oidPlaceKyo);
                            System.out.println("strPlace: " + strPlace);

                            String strSN = snmpGet(sip, oidSN);
                            if (strSN.equals("")) strSN = "!error";
                            System.out.println("strSN: " + strSN);

                            String strModel = snmpGet(sip, oidModel);
                            if (strModel.equals("")) strModel = "!error";
                            System.out.println("strModel: " + strModel);

                            String strMac = snmpGet(sip, oidMacHp);
                            System.out.println("strMac: " + strMac);
                            if (strMac.length() < 6 || strMac.equals("00:00:00:00:00:00"))
                                strMac = snmpGet(sip, oidMacX);
                            System.out.println("strMacX: " + strMac);

                            String insertTableSQL = "INSERT INTO parseFromXLSX"
                                    + "(IP, SN, MAC, HOST, Пробег, Дата, Остаток, MODEL) VALUES"
                                    + "(?,?,?,?,?,?,?,?)";
                            PreparedStatement preparedStatement = con.prepareStatement(insertTableSQL);
                            System.out.println("strMac3: " + strMac);
                            preparedStatement.setString(1, sip);
                            preparedStatement.setString(2, strSN);
                            preparedStatement.setString(3, strMac);
                            preparedStatement.setString(4, strPlace);
                            preparedStatement.setInt(5, pages);
                            preparedStatement.setTimestamp(6, sqlDate);
                            preparedStatement.setInt(7, rem);
                            preparedStatement.setString(8, strModel);
                            preparedStatement.executeUpdate();


                            double k = (double) i / lstR * 100d;
                            progressBar.setValue((int) k);
                        }
                    }

                }

            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                if (rs != null) try {
                    rs.close();
                } catch (Exception e) {
                }
                if (stmt != null) try {
                    stmt.close();
                } catch (Exception e) {
                }
                if (con != null) try {
                    con.close();
                } catch (Exception e) {
                }
            }
            progressBar.setValue(100);

//        text.show();       
//        button.show();
//        picker2.show();
//        picker.show();
//        frame.setSize(new Dimension(350,190));
//    	System.exit(1);
            return null;
        }
    
}
 
    public HelloSNMPExcelToSQL() {
    	super(new BorderLayout());

        progressBar = new JProgressBar(0, 100);
        progressBar.setValue(0);
        progressBar.setStringPainted(true);
        progressBar.setPreferredSize(new Dimension(250,30));

        JPanel panel = new JPanel();
        panel.add(progressBar);
        panel.add(button2);
        add(panel, BorderLayout.PAGE_START);
        setBorder(BorderFactory.createEmptyBorder(20, 20, 30, 20));
        this.setHandler();
        
    }

    public void setHandler() {

        button2.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent arg0) {
                Task privet = new Task();
                privet.execute();
                progressBar.setValue(0);
                button2.hide();
                frame.revalidate();
            }
        });
    }

    private static void createAndShowGUI() {
        
    	frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setLocation(400, 300);
        frame.setPreferredSize(new Dimension(380,190));
        JComponent newContentPane = new HelloSNMPExcelToSQL();
        
        newContentPane.setOpaque(true);
        frame.setContentPane(newContentPane);
        frame.pack();
        frame.setVisible(true);      
            
    }
	
	public static String snmpGet(String strIP, String oidValue) throws IOException
		{	
		  	String rez = "";
		  	TransportMapping transport = new DefaultUdpTransportMapping();
		    transport.listen();

		    CommunityTarget comtarget = new CommunityTarget();
		    comtarget.setCommunity(new OctetString(community));
		    comtarget.setVersion(snmpVersion);
		    comtarget.setAddress(new UdpAddress(strIP + "/161"));
		    comtarget.setRetries(2);
		    comtarget.setTimeout(300);

		    PDU pdu = new PDU();
		    pdu.add(new VariableBinding(new OID(oidValue)));
		    pdu.setType(PDU.GET);
		    pdu.setRequestID(new Integer32(1));

		    Snmp snmp = new Snmp(transport);
		    ResponseEvent response = snmp.get(pdu, comtarget);

		    if (response != null)
		    {
		      PDU responsePDU = response.getResponse();
		      if (responsePDU != null)
		      {
		        int errorStatus = responsePDU.getErrorStatus();
//		        int errorIndex = responsePDU.getErrorIndex();
//		        String errorStatusText = responsePDU.getErrorStatusText();

		        if (errorStatus == PDU.noError)
		        {
		          PDU pduresponse=response.getResponse();
		          String str=pduresponse.getVariableBindings().firstElement().toString();
		          if(str.contains("="))
		          {
		          int len = str.indexOf("=");
		          str=str.substring(len+1, str.length()).trim();
		          }
//		          System.out.println("Snmp Get Response = " + str);
		          rez = str;
		        }
//		        else
		        //  {
		        //  System.out.println("Error: Request Failed");
		        // System.out.println("Error Status = " + errorStatus);
		        // System.out.println("Error Index = " + errorIndex);
		        // System.out.println("Error Status Text = " + errorStatusText);		          
		        //}
		      }
		      //else
		      //{
		      // System.out.println("Error: Response PDU is null");
		      //}
		    }
		    else
		    {
		     System.out.printf("Error: Agent Timeout... [%s]",strIP);
		    }
		    snmp.close();
		    
		    return rez;
		}
	  	  

    public static long compareTwoTimeStamps(java.sql.Timestamp currentTime, java.sql.Timestamp oldTime)
    {
        long milliseconds1 = oldTime.getTime();
      long milliseconds2 = currentTime.getTime();

      long diff = milliseconds2 - milliseconds1;
//      long diffSeconds = diff / 1000;
      long diffMinutes = diff / (60 * 1000);
//      long diffHours = diff / (60 * 60 * 1000);
//      long diffDays = diff / (24 * 60 * 60 * 1000);

        return diffMinutes;
    }


    static void print ( String[] _str){
        //Выводим список аргументов командной строки в консоль
        for (int i =0; i<_str.length;i++){
            System.out.println("[" + i + "]: " +_str[i]);
        }
    }


    public static void main(String[] args) {

        Vector<Integer> VecArgsNum = null;
        System.out.println("Всего аргументов в командной строке: "+args.length);
        print (args);
        VecArgsNum = new Vector<Integer>(0);
        {// блок проверки исходных данных
            System.out.println("Требуется два аргумента путь таблице, sql сервер с .lsr.ru порт по умолчанию 1433");
            System.out.println("Требуется databaseName=parserResult;user=snmp;password=145236");
            System.out.println("Скрирт для создания таблицы прилагается parsingDoneGeneratedScriptTable.sql");
            if (args.length != 2) {
                //Ошибка неправильное количество аргументов
                System.out.println("Ошибка неправильное количество аргументов: " + args.length);

                return;
            }
            //Преобразование массива строк во входные переменные
            System.out.println("Преобразование массива строк в переменные окружения");
            fileSourceXLSX=args[0];
            connectionUrl="jdbc:sqlserver://"+args[1]+":1433;" +
                    "databaseName=parserResult;user=snmp;password=145236";
            System.out.printf("fileSourceXLSX=[%s]\n",fileSourceXLSX);
            System.out.printf("connectionUrl=[%s]\n",connectionUrl);
        }

		
		javax.swing.SwingUtilities.invokeLater(new Runnable() {
	        public void run() {
	            createAndShowGUI();
	        }
	    });

   }

}