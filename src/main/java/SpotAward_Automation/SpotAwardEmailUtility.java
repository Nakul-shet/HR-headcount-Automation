package SpotAward_Automation;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.*;
import java.util.Arrays;
import java.util.Properties;

public class SpotAwardEmailUtility {

    public static void sendEmail(String [] recipients, String sender, String subject, String body, String attachmentPath) {
        final String password = "N@!@#$kuLk09";

        Properties properties = System.getProperties();
        properties.setProperty("mail.smtp.host", "ismtp.allegisgroup.com");
        properties.put("mail.smtp.auth", "true");

        // Create a session with authentication
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

            // Create the email body
            MimeBodyPart textPart = new MimeBodyPart();
            textPart.setContent(body, "text/html; charset=utf-8"); // Ensure the email body is sent as HTML content

            // Create the attachment part
            MimeBodyPart attachmentPart = new MimeBodyPart();
            DataSource source = new FileDataSource(attachmentPath);
            attachmentPart.setDataHandler(new DataHandler(source));
            attachmentPart.setFileName(source.getName());

            // Combine the parts into a multipart
            Multipart multipart = new MimeMultipart();
            multipart.addBodyPart(textPart);
            multipart.addBodyPart(attachmentPart);

            // Set the content of the message
            message.setContent(multipart);

            // Send the email
            Transport.send(message);
            System.out.println("Email sent successfully to " + Arrays.toString(recipients));
        } catch (MessagingException e) {
            e.printStackTrace();
        }
    }
}
