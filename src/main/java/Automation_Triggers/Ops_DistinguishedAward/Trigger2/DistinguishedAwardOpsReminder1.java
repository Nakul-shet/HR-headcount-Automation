package Automation_Triggers.Ops_DistinguishedAward.Trigger2;

import Utilities.Common.EmailBodyUtilities;
import Utilities.Configuration.AppConfig;
import Utilities.Configuration.MasterConfig;
import Utilities.Service.DistinguishedAwardOperationsEmailBodyBuilderService;
import Utilities.Service.DistinguishedAwardPracticeEmailBodyBuilderService;
import Utilities.Configuration.SpotAwardConfig;
import Utilities.Common.EmailSenderUtilities;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;

import static Utilities.Common.EmailSenderUtilities.getCCEmailBasedOnRunType;
import static Utilities.Common.EmailSenderUtilities.getToEmailBasedOnRunType;

public class DistinguishedAwardOpsReminder1 {
    static AppConfig config = MasterConfig.getDataBasedOnActiveConfig(MasterConfig.activeEnvironment);
    public static void main(String[] args) {
        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(new ByteArrayOutputStream()));
        String emailBody = "";
        try {
            emailBody = DistinguishedAwardOperationsEmailBodyBuilderService.buildReminderEmailBody1();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.setOut(originalOut);
        }
        String sender = config.getSenderId();
        String subject = "Reminder: Nominate or Regret! - Nominations for the Reward and Recognition (Distinguished Award) - Q" + EmailBodyUtilities.getCurrentQuarter() + " " + java.time.LocalDate.now().getYear() + " - Operational Support";
        EmailSenderUtilities.sendEmail(
                getToEmailBasedOnRunType(config.getLocalRunFor() , "op-to"),
                getCCEmailBasedOnRunType(config.getLocalRunFor() , "op-cc"),
                sender,
                subject,
                emailBody,
                "distinguished_operations"
        );
    }
}
