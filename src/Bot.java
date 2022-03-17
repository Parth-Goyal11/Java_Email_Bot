import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Properties;
import javax.mail.*;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Bot {
    Session session = null;
    MimeMessage email = null;

    FileInputStream file;





    public static void main(String[] args) throws MessagingException, IOException {
        Bot mailer = new Bot();
        mailer.setServerConfiguration();
        mailer.writeEmail();
        mailer.sendEmailAndAuthenticate();
        System.out.println("Email Sent Successfully");


    }
    public void setServerConfiguration(){
        Properties properties = new Properties();
        properties.put("mail.smtp.auth", "true");
        properties.put("mail.smtp.starttls.enable", "true");
        properties.put("mail.smtp.host", "smtp.gmail.com");
        properties.put("mail.smtp.port", "587");


        session = Session.getDefaultInstance(properties, null);
    }

    public MimeMessage writeEmail() throws MessagingException, IOException {
        ArrayList<String> recipients = returnEmails();
        email = new MimeMessage(session);

        for(int i = 0; i<recipients.size(); i++){
            email.addRecipient(Message.RecipientType.TO, new InternetAddress(recipients.get(i)));
        }

        
        return email;
    }

    public void sendEmailAndAuthenticate() throws MessagingException, IOException {
        String sender = "epicparth@gmail.com";
        String password = "****************";
        String server = "smpt.gmail.com";
        ArrayList<String> names = returnNames();
        ArrayList<String> emails = returnEmails();
        Address[] array = email.getAllRecipients();
        for(int j = 0; j<names.size(); j++){
            Address[] currReceiver = {array[j]};

            email.setSubject("Hello " + names.get(j));
            email.setText("Dear " + names.get(j) +","+ "\n This is being emailed with the JavaMail API ");
            Transport.send(email, currReceiver, sender, password);
        }

    }

    public static ArrayList<String> returnNames() throws IOException {
        FileInputStream file;
        ArrayList<String> emails = new ArrayList<>();
        ArrayList<String> names = new ArrayList<>();

        try {

            file = new FileInputStream("TestFile.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {

                Row row = rowIterator.next();

                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case NUMERIC:
                            break;
                        case STRING:
                            if (cell.getStringCellValue().contains("@")) {
                                emails.add(cell.getStringCellValue());
                            } else {
                                names.add(cell.getStringCellValue());
                            }
                            break;
                    }

                }

            }
            file.close();
            return names;

        } catch (IOException e) {
            e.printStackTrace();
        }


        return names;
    }

    public static ArrayList<String> returnEmails() throws IOException {
        FileInputStream file;
        ArrayList<String> emails = new ArrayList<>();
        ArrayList<String> names = new ArrayList<>();

        try {

            file = new FileInputStream("TestFile.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {

                Row row = rowIterator.next();

                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case NUMERIC:
                            break;
                        case STRING:
                            if (cell.getStringCellValue().contains("@")) {
                                emails.add(cell.getStringCellValue());
                            } else {
                                names.add(cell.getStringCellValue());
                            }
                            break;
                    }

                }

            }
            file.close();
            return emails;

        } catch (IOException e) {
            e.printStackTrace();
        }


        return emails;
    }

}
