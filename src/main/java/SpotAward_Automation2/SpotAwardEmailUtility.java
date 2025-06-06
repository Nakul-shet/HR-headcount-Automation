package SpotAward_Automation2;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.*;
import java.util.Arrays;
import java.util.Properties;

public class SpotAwardEmailUtility {

    public static void sendEmail(String [] recipients, String sender, String subject, String body) {
        final String password = SpotAwardConfig.SENDER_PASSWORD;

        Properties properties = System.getProperties();
        properties.setProperty("mail.smtp.host", "ismtp.allegisgroup.com");
        properties.put("mail.smtp.auth", "true");

        Session session = Session.getInstance(properties, new Authenticator() {
            @Override
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(sender, password);
            }
        });

        try {

            MimeMessage message = new MimeMessage(session);
            message.setFrom(new InternetAddress(sender));
            for (String recipient : recipients) {
                message.addRecipient(Message.RecipientType.TO, new InternetAddress(recipient));
            }

            message.setSubject(subject);

            MimeBodyPart textPart = new MimeBodyPart();
            textPart.setContent(body, "text/html; charset=utf-8"); // Ensure the email body is sent as HTML content

//            MimeBodyPart attachmentPart = new MimeBodyPart();
//            DataSource source = new FileDataSource(attachmentPath);
//            attachmentPart.setDataHandler(new DataHandler(source));
//            attachmentPart.setFileName(source.getName());

            Multipart multipart = new MimeMultipart();
            multipart.addBodyPart(textPart);
            //multipart.addBodyPart(attachmentPart);

            message.setContent(multipart);

            Transport.send(message);
            System.out.println("Email sent successfully to " + Arrays.toString(recipients));
        } catch (MessagingException e) {
            e.printStackTrace();
        }
    }
}
