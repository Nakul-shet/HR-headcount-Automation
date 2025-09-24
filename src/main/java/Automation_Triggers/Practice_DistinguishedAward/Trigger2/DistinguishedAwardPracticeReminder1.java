package Automation_Triggers.Practice_DistinguishedAward.Trigger2;

import Utilities.Common.EmailBodyUtilities;
import Utilities.Common.EmailSenderUtilities;
import Utilities.Configuration.AppConfig;
import Utilities.Configuration.MasterConfig;
import Utilities.Configuration.SpotAwardConfig;
import Utilities.Service.DistinguishedAwardPracticeEmailBodyBuilderService;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;

import static Utilities.Common.EmailSenderUtilities.getCCEmailBasedOnRunType;
import static Utilities.Common.EmailSenderUtilities.getToEmailBasedOnRunType;

public class DistinguishedAwardPracticeReminder1 {
    static AppConfig config = MasterConfig.getDataBasedOnActiveConfig(MasterConfig.activeEnvironment);
    public static void main(String[] args) {
        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(new ByteArrayOutputStream()));
        String emailBody = "";
        try {
            emailBody = DistinguishedAwardPracticeEmailBodyBuilderService.buildReminderEmailBody1();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.setOut(originalOut);
        }
        String sender = config.getSenderId();
        String subject = "Reminder: Nominate or Regret! - Nominations for the Reward and Recognition (Distinguished Award) - Q" + EmailBodyUtilities.getCurrentQuarter() + " " + java.time.LocalDate.now().getYear();
        EmailSenderUtilities.sendEmail(
                getToEmailBasedOnRunType(config.getLocalRunFor() , "prac-to"),
                getCCEmailBasedOnRunType(config.getLocalRunFor() , "prac-cc"),
                sender,
                subject,
                emailBody,
                "distinguished_practice"
        );
    }
}
