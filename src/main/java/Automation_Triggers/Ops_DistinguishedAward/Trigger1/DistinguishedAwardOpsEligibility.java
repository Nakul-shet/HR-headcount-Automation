package Automation_Triggers.Ops_DistinguishedAward.Trigger1;

import Utilities.Common.EmailBodyUtilities;
import Utilities.Common.EmailSenderUtilities;

import Utilities.Configuration.AppConfig;
import Utilities.Configuration.MasterConfig;
import Utilities.Configuration.SpotAwardConfig;
import Utilities.Service.DistinguishedAwardOperationsEmailBodyBuilderService;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;


import static Utilities.Common.EmailSenderUtilities.getCCEmailBasedOnRunType;
import static Utilities.Common.EmailSenderUtilities.getToEmailBasedOnRunType;

public class DistinguishedAwardOpsEligibility {
    static AppConfig config = MasterConfig.getDataBasedOnActiveConfig(MasterConfig.activeEnvironment);
    public static void main(String[] args) {
        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(new ByteArrayOutputStream()));
        String emailBody = "";
        try {
            emailBody = DistinguishedAwardOperationsEmailBodyBuilderService.buildEligibilityEmailBody();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.setOut(originalOut);
        }
        String sender = config.getSenderId();
        String subject = "Nominations for the Reward and Recognition (Distinguished Award) - Q" + EmailBodyUtilities.getCurrentQuarter() + " " + java.time.LocalDate.now().getYear() + " - Operational Support";
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
