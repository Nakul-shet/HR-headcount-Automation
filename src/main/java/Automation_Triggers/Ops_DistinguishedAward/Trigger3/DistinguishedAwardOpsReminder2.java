package Automation_Triggers.Ops_DistinguishedAward.Trigger3;

import Utilities.Common.EmailBodyUtilities;
import Utilities.Configuration.AppConfig;
import Utilities.Configuration.MasterConfig;
import Utilities.Service.DistinguishedAwardOperationsEmailBodyBuilderService;
import Utilities.Service.DistinguishedAwardPracticeEmailBodyBuilderService;
import Utilities.Configuration.SpotAwardConfig;
import Utilities.Common.EmailSenderUtilities;

import java.io.ByteArrayOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintStream;

import static Utilities.Common.EmailSenderUtilities.getCCEmailBasedOnRunType;
import static Utilities.Common.EmailSenderUtilities.getToEmailBasedOnRunType;

public class DistinguishedAwardOpsReminder2 {
    static AppConfig config = MasterConfig.getDataBasedOnActiveConfig(MasterConfig.activeEnvironment);
    public static void main(String[] args) {
        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(new ByteArrayOutputStream()));
        String emailBody = "";
        try {
            emailBody = DistinguishedAwardOperationsEmailBodyBuilderService.buildReminderEmailBody2();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.setOut(originalOut);
        }
        String sender = config.getSenderId();
        String subject = "Last Call: Don’t Ghost the Greats! Nominate Now! - Nominations for the Reward and Recognition (Distinguished Award) - Q" + EmailBodyUtilities.getCurrentQuarter() + " " + java.time.LocalDate.now().getYear() + " - Operational Support";
        EmailSenderUtilities.sendEmail(
                getToEmailBasedOnRunType(config.getLocalRunFor() , "op-to"),
                getCCEmailBasedOnRunType(config.getLocalRunFor() , "op-cc"),
                sender,
                subject,
                emailBody,
                "distinguished_operations"
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
        String filePath = System.getProperty("user.dir") + "/src/main/resources/distinguished_operations_message_id.txt";
        try (FileWriter writer = new FileWriter(filePath, false)) {
            writer.write("Expired");
            System.out.println("distinguished_operations_message_id.txt has been cleared.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void clearMessageIdFileJenkins() {
        String filePath = "/var/jenkins_home/shared/distinguished_operations_message_id.txt";
        try (FileWriter writer = new FileWriter(filePath, false)) {
            writer.write("Expired");
            System.out.println("distinguished_operations_message_id.txt has been cleared.");
        } catch (IOException e) {
            System.err.println("Failed to clear distinguished_operations_message_id.txt");
            e.printStackTrace();
        }
    }
}
