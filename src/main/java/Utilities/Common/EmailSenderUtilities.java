package Utilities.Common;

import Utilities.Configuration.AppConfig;
import Utilities.Configuration.MasterConfig;
import Utilities.Configuration.SpotAwardConfig;

import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Objects;
import java.util.Properties;

public class EmailSenderUtilities {

    public static void sendEmail(String[] recipients, String[] ccRecipients, String sender, String subject, String body, String emailType) {
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

            if(subject.contains("Congratulations")){
                for (String recipient : recipients) {
                    message.addRecipient(Message.RecipientType.BCC, new InternetAddress(recipient));
                }
            }else{
                for (String recipient : recipients) {
                    message.addRecipient(Message.RecipientType.TO, new InternetAddress(recipient));
                }
            }

            if (ccRecipients != null) {
                for (String ccRecipient : ccRecipients) {
                    message.addRecipient(Message.RecipientType.CC, new InternetAddress(ccRecipient));
                }
            }

            message.setSubject(subject);

            MimeBodyPart textPart = new MimeBodyPart();
            textPart.setContent(body, "text/html; charset=utf-8");

            Multipart multipart = new MimeMultipart();
            multipart.addBodyPart(textPart);

            if(subject.contains("Congratulations")){
                MimeBodyPart attachmentPart = new MimeBodyPart();
                attachmentPart.attachFile(System.getProperty("user.dir") + "//DataFiles//Pluxee card activation processes.eml"); // or .pdf, .txt, etc.
                attachmentPart.setFileName("Pluxee Card Activation Processes.eml"); // optional, controls download name
                multipart.addBodyPart(attachmentPart);
            }

            message.setContent(multipart);

            String messageId;
            boolean messageIdAvailable;
            boolean expiredId;

            switch (SpotAwardConfig.runEnvironment) {
                case "local":
                    messageIdAvailable = false;
                    expiredId = false;

                    String []emailTypes = emailType.contains(" ")?emailType.trim().split(" "): new String[]{emailType};

                    if(!(emailTypes.length >1)){

                        System.out.println(Arrays.toString(emailTypes));

                        String path = "/src/main/resources/" + emailTypes[0] + "_message_id.txt";

                        try (BufferedReader reader = new BufferedReader(new FileReader(System.getProperty("user.dir") + path))) {
                            messageId = reader.readLine();

                            if (messageId != null && !messageId.isEmpty()) {
                                if(messageId.equals("Expired")){
                                    expiredId = true;
                                }
                                else{
                                    messageIdAvailable = true;
                                    message.setHeader("In-Reply-To", messageId);
                                    message.setHeader("References", messageId);
                                }
                            }
                        } catch (IOException e) {
                            e.printStackTrace();
                        }

                        message.saveChanges();
                        Transport.send(message);

                        if (!messageIdAvailable && !expiredId) {
                            messageId = message.getMessageID();
                            try (FileWriter writer = new FileWriter(System.getProperty("user.dir") + path)) {
                                writer.write(messageId);
                            } catch (IOException e) {
                                e.printStackTrace();
                            }
                        }
                    }
                    else{
                        System.out.println(Arrays.toString(emailTypes));
                        message.saveChanges();
                        Transport.send(message);
                    }

                    break;
                case "jenkins":
                    messageId = null;
                    messageIdAvailable = false;

                    try (BufferedReader reader = new BufferedReader(new FileReader("/var/jenkins_home/shared/spot_practice_message_id.txt"))) {
                        messageId = reader.readLine();
                        if (messageId != null && !messageId.isEmpty()) {
                            messageIdAvailable = true;

                            message.setHeader("In-Reply-To", messageId);
                            message.setHeader("References", messageId);
                        }
                        if (Objects.equals(messageId, "Expired")) {
                            messageIdAvailable = true;
                        }
                    } catch (IOException e) {
                        e.printStackTrace();
                    }

                    message.saveChanges();
                    Transport.send(message);

                    if (!messageIdAvailable) {
                        messageId = message.getMessageID();
                        try (FileWriter writer = new FileWriter("/var/jenkins_home/shared/spot_practice_message_id.txt")) {
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
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    static AppConfig config = MasterConfig.getDataBasedOnActiveConfig(MasterConfig.activeEnvironment);

    public static String[] getToEmailBasedOnRunType(String runType, String orgTo) {
        String[] recepients = new String[0];
        if (runType.equalsIgnoreCase("test")) {
            recepients = config.getRecipients();
        } else if (config.getLocalRunFor().equalsIgnoreCase("prod")) {
            recepients = Objects.requireNonNull(ExcelUtilities.getToEmailAddresses(orgTo));
        }
        return recepients;
    }

    public static String[] getCCEmailBasedOnRunType(String runType, String orgCC) {
        String[] cc = new String[0];
        if (runType.equalsIgnoreCase("test")) {
            cc = null;
        } else if (config.getLocalRunFor().equalsIgnoreCase("prod")) {
            cc = Objects.requireNonNull(ExcelUtilities.getToEmailAddresses(orgCC));
        }
        return cc;
    }
}
