package body;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;



public class CLCShippedReport {
    public static void main(String[] args) {
        Connection connect = null;
        Statement statement = null;
        ResultSet rs = null;

        try {
            String SQL = "select * from Graph.countries";

            File connectfile = new File("src/body/Config.properties");
            Properties connectini = new Properties();
            connectini.load(new FileInputStream(connectfile));
            
            // Update these properties for MySQL connection
            String mysqlUrl = connectini.getProperty("spring.datasource.url");
            String mysqlUsername = connectini.getProperty("spring.datasource.username");
            String mysqlPassword = connectini.getProperty("spring.datasource.password");

            Class.forName("com.mysql.cj.jdbc.Driver");
            connect = DriverManager.getConnection(mysqlUrl, mysqlUsername, mysqlPassword);

            statement = connect.createStatement();
            rs = statement.executeQuery(SQL);
            try (SXSSFWorkbook wb = new SXSSFWorkbook(5000)) {
				SXSSFSheet sheet = wb.createSheet();
				int rownum = 0;
				SXSSFRow rows = sheet.createRow(rownum);
				ResultSetMetaData md = rs.getMetaData();
				int columnCount = md.getColumnCount();
				int i;
				for (i = 1; i <= columnCount; i++)
				    rows.createCell(i - 1).setCellValue(md.getColumnName(i));
				while (rs.next()) {
				    rownum++;
				    rows = sheet.createRow(rownum);
				    for (i = 1; i <= columnCount; i++)
				        rows.createCell(i - 1).setCellValue(rs.getString(i));
				}

				FileOutputStream xlsStream = new FileOutputStream("CLCData.xlsx");
				wb.write(xlsStream);
				wb.close();
				xlsStream.close();
				
				System.out.println("excel completed");
			}
            sendEmailWithAttachment("CLCData.xlsx", connectini);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (rs != null) {
                    rs.close();
                }
                if (statement != null) {
                    statement.close();
                }
                if (connect != null) {
                    connect.close();
                }
            } catch (Exception e) {
            	e.printStackTrace();
            }
        }
    }

    private static void sendEmailWithAttachment(String attachmentFilename, Properties configProps) {
        final String username = configProps.getProperty("mailFrom");
        final String password = configProps.getProperty("mailPassword");
        String host = configProps.getProperty("mail.smtp.host");

        Properties props = new Properties();
        props.put("mail.smtp.host", host);
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.port", configProps.getProperty("mail.smtp.port"));
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.ssl.trust", "smtp.gmail.com");  // Replace with your SMTP server host if not using Gmail


        Session session = Session.getInstance(props,
            new Authenticator() {
                protected PasswordAuthentication getPasswordAuthentication() {
                    return new PasswordAuthentication(username, password);
                }
            });

        try {
       
        String subjectTemplate = configProps.getProperty("mail.subject");
        String currentDay=getCurrentDay();
        String subject = subjectTemplate.replace("{currentDay}", currentDay);
       
       
            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress(username));
           
            String[] toRecipients = configProps.getProperty("Mail_TO").split(",");
            for (String toRecipient : toRecipients) {
                message.addRecipients(Message.RecipientType.TO, InternetAddress.parse(toRecipient.trim()));
            }
           
           
           message.addRecipients(Message.RecipientType.CC, InternetAddress.parse(configProps.getProperty("Mail_CC")));
           
           message.setSubject(subject);
            MimeBodyPart messageBodyPart = new MimeBodyPart();
            messageBodyPart.setContent(configProps.getProperty("mail.messagebody"), "text/html");
            MimeBodyPart attachmentPart = new MimeBodyPart();
            DataSource source = new FileDataSource(attachmentFilename);
            attachmentPart.setDataHandler(new DataHandler(source));
            attachmentPart.setFileName(configProps.getProperty("mail.attachment.filename"));
            Multipart multipart = new MimeMultipart();
            multipart.addBodyPart(messageBodyPart);
            multipart.addBodyPart(attachmentPart);
            message.setContent(multipart);
           

            Transport.send(message);
            System.out.println("sent mail");

         

        } catch (MessagingException e) {
            throw new RuntimeException(e);
        }
    }


 private static String getCurrentDay() {
       SimpleDateFormat df = new SimpleDateFormat("dd-MM-yyyy");
       Date today = Calendar.getInstance().getTime();
       return df.format(today);
   }
 
}
