package HR_Automation_Utilities;

import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Arrays;
import java.util.Objects;
import java.util.Properties;

public class SpotAwardEmailSenderUtility {

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
            textPart.setContent(body, "text/html; charset=utf-8");

            Multipart multipart = new MimeMultipart();
            multipart.addBodyPart(textPart);

            message.setContent(multipart);

            String messageId;
            boolean messageIdAvailable;

            switch(SpotAwardConfig.runEnvironment){
                case "local":
                    messageIdAvailable = false;
                    messageId = null;

                    try (BufferedReader reader = new BufferedReader(new FileReader(System.getProperty("user.dir") + "/src/main/resources/message_id.txt"))) {
                        messageId = reader.readLine();
                        if (messageId != null && !messageId.isEmpty()) {
                            messageIdAvailable = true;

                            message.setHeader("In-Reply-To", messageId);
                            message.setHeader("References", messageId);
                        }
                        if(Objects.equals(messageId, "Expired")){
                            messageIdAvailable = true;
                        }
                    } catch (IOException e) {
                        e.printStackTrace();
                    }

                    message.saveChanges();

                    Transport.send(message);

                    if(!messageIdAvailable){
                        messageId = message.getMessageID();
                        try (FileWriter writer = new FileWriter(System.getProperty("user.dir")+"/src/main/resources/message_id.txt")) {
                            writer.write(messageId);
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                    break;
                case "remote":
                    messageId = null;
                    messageIdAvailable = false;

                    try (BufferedReader reader = new BufferedReader(new FileReader("/var/jenkins_home/shared/message_id.txt"))) {
                        messageId = reader.readLine();
                        if (messageId != null && !messageId.isEmpty()) {
                            messageIdAvailable = true;

                            message.setHeader("In-Reply-To", messageId);
                            message.setHeader("References", messageId);
                        }
                        if(Objects.equals(messageId, "Expired")){
                            messageIdAvailable = true;
                        }
                    } catch (IOException e) {
                        e.printStackTrace();
                    }

                    message.saveChanges();
                    Transport.send(message);

                    if (!messageIdAvailable) {
                        messageId = message.getMessageID();
                        try (FileWriter writer = new FileWriter("/var/jenkins_home/shared/message_id.txt")) {
                            writer.write(messageId);
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                    break;
            }

            System.out.println("Email sent successfully to " + Arrays.toString(recipients));
        } catch (MessagingException e) {
            e.printStackTrace();
        }
    }
}
