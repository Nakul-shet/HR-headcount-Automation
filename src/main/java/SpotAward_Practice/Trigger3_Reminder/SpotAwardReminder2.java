package SpotAward_Practice.Trigger3_Reminder;

import HR_Automation_Utilities.SpotAwardConfig;
import HR_Automation_Utilities.SpotAwardEmailSenderUtility;
import HR_Automation_Utilities.SpotAwardEmailBodyBuilderService;

import java.io.ByteArrayOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintStream;

import static HR_Automation_Utilities.SpotAwardEmailSenderUtility.getCCEmailBasedOnRunType;
import static HR_Automation_Utilities.SpotAwardEmailSenderUtility.getToEmailBasedOnRunType;

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
                getToEmailBasedOnRunType(SpotAwardConfig.localRunFor , "prac-to"),
                getCCEmailBasedOnRunType(SpotAwardConfig.localRunFor , "prac-cc"),
                sender,
                subject,
                emailBody,
                "spot_practice"
        );

        switch(SpotAwardConfig.runEnvironment){
            case "local":
                clearMessageIdFileLocal();
                break;
            case "remote":
                clearMessageIdFileJenkins();
                break;
        }

    }

    public static void clearMessageIdFileLocal() {
        String filePath = System.getProperty("user.dir") + "/src/main/resources/spot_practice_message_id.txt";
        try (FileWriter writer = new FileWriter(filePath, false)) {
            writer.write("Expired");
            System.out.println("spot_practice_message_id.txt has been cleared.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void clearMessageIdFileJenkins() {
        String filePath = "/var/jenkins_home/shared/spot_practice_message_id.txt";
        try (FileWriter writer = new FileWriter(filePath, false)) {
            writer.write("Expired");
            System.out.println("spot_practice_message_id.txt has been cleared.");
        } catch (IOException e) {
            System.err.println("Failed to clear spot_practice_message_id.txt");
            e.printStackTrace();
        }
    }

}
