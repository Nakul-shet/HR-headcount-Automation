package Trigger2_Reminder;

import HR_Automation_Utilities.SpotAwardConfig;
import HR_Automation_Utilities.SpotAwardEmailSenderUtility;
import HR_Automation_Utilities.SpotAwardEmailBodyBuilderService;

import java.io.ByteArrayOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintStream;

public class SpotAwardReminder2 {

    public static void main(String[] args) {
        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(new ByteArrayOutputStream()));
        String emailBody = "";
        try {
            emailBody = SpotAwardEmailBodyBuilderService.buildReminderEmailBody2();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.setOut(originalOut);
        }
        String sender = SpotAwardConfig.SENDER_ID;
        String subject = "Last Call: Don’t Ghost the Greats! Nominate Now! - Spot Awards " + java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH) + " " + java.time.LocalDate.now().getYear();
        SpotAwardEmailSenderUtility.sendEmail(
                SpotAwardConfig.RECIPIENTS,
                sender,
                subject,
                emailBody
        );

        //clearMessageIdFileLocal();
        clearMessageIdFileJenkins();
    }

    public static void clearMessageIdFileLocal() {
        String filePath = System.getProperty("user.dir") + "/src/main/resources/message_id.txt";
        try (FileWriter writer = new FileWriter(filePath, false)) {
            writer.write("");
            System.out.println("message_id.txt has been cleared.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void clearMessageIdFileJenkins() {
        String filePath = "/var/jenkins_home/shared/message_id.txt";
        try (FileWriter writer = new FileWriter(filePath, false)) {
            writer.write("");
            System.out.println("message_id.txt has been cleared.");
        } catch (IOException e) {
            System.err.println("Failed to clear message_id.txt");
            e.printStackTrace();
        }
    }


}
