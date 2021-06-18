package zer0p.tools;

import org.apache.commons.cli.*;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.Message;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.InputStreamReader;
import java.sql.*;
import java.util.Date;
import java.util.Properties;

public class MainApp {

    static Logger log = LoggerFactory.getLogger("data_excel_mail");

    public static void main(String[] args) {
        try {

            String path = "";

            try {
                path = MainApp.class.getProtectionDomain().getCodeSource().getLocation().getPath();
                path = path.substring(0,path.lastIndexOf("."));
                path = path.substring(0,path.lastIndexOf("/"));
                path = java.net.URLDecoder.decode(path, "UTF-8"); // 转换处理中文及空格
            }catch (Exception ex){
                path = MainApp.class.getClassLoader().getResource(".").getPath();
            }

            log.info("path:" + path);

            File fproperties = new File(path + File.separator + "data_excel_mail.properties");

            Properties cmd = new Properties();
            cmd.load(new InputStreamReader(new FileInputStream(fproperties),"utf-8"));

            String mDriver = cmd.getProperty("driver");
            String mConn = cmd.getProperty("conn");
            String mUser = cmd.getProperty("u");
            String mPsw = cmd.getProperty("p");
            String mSql = cmd.getProperty("sql");
            String mFile = cmd.getProperty("file");
            String mMail = cmd.getProperty("mail");
            String mMailserver = cmd.getProperty("mailserver");
            String msender = cmd.getProperty("sender");
            String msenderpsw = cmd.getProperty("senderpsw");

            log.info(mDriver);
            log.info(mConn);
            log.info(mSql);
            log.info(mFile);
            log.info(mMail);

            //database
            Class.forName(mDriver);
            Connection conn = DriverManager.getConnection(mConn, mUser, mPsw);
            PreparedStatement ps = conn.prepareStatement(mSql);
            ResultSet data = ps.executeQuery();
            ResultSetMetaData rsmd = data.getMetaData();


            //excel
            HSSFWorkbook workbook = new HSSFWorkbook();
            Integer rowIndex = 0;

            HSSFSheet sheet = workbook.createSheet();

            HSSFRow rowHead = sheet.createRow(rowIndex);
            for (int i = 1; i <= rsmd.getColumnCount(); i++) {
                rowHead.createCell(i - 1).setCellValue(rsmd.getColumnName(i));
            }
            rowIndex++;

            while (data.next()) {
                HSSFRow row = sheet.createRow(rowIndex);

                for (int i = 1; i <= rsmd.getColumnCount(); i++) {
                    row.createCell(i - 1).setCellValue(data.getString(i));
                }

                rowIndex = rowIndex + 1;

                if (rowIndex >= 65525) {
                    sheet = workbook.createSheet();
                    rowIndex = 0;
                }
            }

            log.info("rows :" + rowIndex);

            //write
            String attfilePath = path + File.separator + mFile + ".xlsx";
            workbook.write(new File(attfilePath));

            //mail
            if (mMailserver != null && mMailserver.isEmpty() == false) {
                Properties props = new Properties();
                props.setProperty("mail.smtp.auth", "true");
                props.setProperty("mail.transport.protocol", "smtp");
                props.setProperty("mail.smtp.host", mMailserver);

                Session session = Session.getInstance(props);
                session.setDebug(true);

                MimeMessage msg = new MimeMessage(session);
                msg.setFrom(new InternetAddress(msender));
                msg.setRecipient(MimeMessage.RecipientType.TO, new InternetAddress(mMail));
                msg.setSubject("data_excel_mail " + mFile, "UTF-8");
                msg.setSentDate(new Date());

                MimeBodyPart attachment = new MimeBodyPart();
                DataHandler dh2 = new DataHandler(new FileDataSource(attfilePath));
                attachment.setDataHandler(dh2);
                attachment.setFileName(MimeUtility.encodeText(dh2.getName()));

                MimeMultipart mm = new MimeMultipart();
                mm.addBodyPart(attachment);

                msg.setContent(mm);

                Transport transport = session.getTransport();
                transport.connect(msender,msenderpsw);
                transport.sendMessage(msg, msg.getAllRecipients());
                //transport.sendMessage(msg, new Address[]{new InternetAddress("xxx@qq.com")});
                transport.close();
            }
        } catch (Exception ex) {
            log.error(ex.getMessage());
            ex.printStackTrace();
        }
    }
}
