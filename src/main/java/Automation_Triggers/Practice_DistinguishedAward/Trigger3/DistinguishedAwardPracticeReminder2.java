package Automation_Triggers.Practice_DistinguishedAward.Trigger3;

import Utilities.Common.EmailBodyUtilities;
import Utilities.Common.EmailSenderUtilities;
import Utilities.Configuration.SpotAwardConfig;
import Utilities.Service.DistinguishedAwardPracticeEmailBodyBuilderService;

import java.io.ByteArrayOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintStream;

import static Utilities.Common.EmailSenderUtilities.getCCEmailBasedOnRunType;
import static Utilities.Common.EmailSenderUtilities.getToEmailBasedOnRunType;

public class DistinguishedAwardPracticeReminder2 {

    public static void main(String[] args) {
        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(new ByteArrayOutputStream()));
        String emailBody = "";
        try {
            emailBody = DistinguishedAwardPracticeEmailBodyBuilderService.buildReminderEmailBody2();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.setOut(originalOut);
        }
        String sender = SpotAwardConfig.SENDER_ID;
        String subject = "Last Call: Don’t Ghost the Greats! Nominate Now! - Nominations for the Reward and Recognition (Distinguished Award) - Q" + EmailBodyUtilities.getCurrentQuarter() + " " + java.time.LocalDate.now().getYear();
        EmailSenderUtilities.sendEmail(
                getToEmailBasedOnRunType(SpotAwardConfig.localRunFor , "prac-to"),
                getCCEmailBasedOnRunType(SpotAwardConfig.localRunFor , "prac-cc"),
                sender,
                subject,
                emailBody,
                "distinguished_practice"
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
        String filePath = System.getProperty("user.dir") + "/src/main/resources/distinguished_practice_message_id.txt";
        try (FileWriter writer = new FileWriter(filePath, false)) {
            writer.write("Expired");
            System.out.println("distinguished_practice_message_id.txt has been cleared.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void clearMessageIdFileJenkins() {
        String filePath = "/var/jenkins_home/shared/distinguished_practice_message_id.txt";
        try (FileWriter writer = new FileWriter(filePath, false)) {
            writer.write("Expired");
            System.out.println("distinguished_practice_message_id.txt has been cleared.");
        } catch (IOException e) {
            System.err.println("Failed to clear distinguished_practice_message_id.txt");
            e.printStackTrace();
        }
    }
}
