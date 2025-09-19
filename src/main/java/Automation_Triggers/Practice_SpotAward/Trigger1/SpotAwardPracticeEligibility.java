package Automation_Triggers.Practice_SpotAward.Trigger1;

import Utilities.Configuration.AppConfig;
import Utilities.Configuration.MasterConfig;
import Utilities.Configuration.SpotAwardConfig;
import Utilities.Common.EmailSenderUtilities;
import Utilities.Service.SpotAwardPracticeEmailBodyBuilderService;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.Locale;

import static Utilities.Common.EmailSenderUtilities.getCCEmailBasedOnRunType;
import static Utilities.Common.EmailSenderUtilities.getToEmailBasedOnRunType;

public class SpotAwardPracticeEligibility {

    static AppConfig config = MasterConfig.getDataBasedOnActiveConfig(MasterConfig.activeEnvironment);

    public static void main(String[] args) {
        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(new ByteArrayOutputStream()));
        String emailBody = "";
        try {
            emailBody = SpotAwardPracticeEmailBodyBuilderService.buildEligibilityEmailBody();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.setOut(originalOut);
        }

        String sender = config.getSenderId();
        String subject = "Spot Awards " + LocalDate.now().getMonth().getDisplayName(TextStyle.FULL, Locale.ENGLISH) + " " + LocalDate.now().getYear();

        EmailSenderUtilities.sendEmail(
                getToEmailBasedOnRunType(config.getLocalRunFor() , "prac-to"),
                getCCEmailBasedOnRunType(config.getLocalRunFor() , "prac-cc"),
                sender,
                subject,
                emailBody,
                "spot_practice"
        );
    }

}
