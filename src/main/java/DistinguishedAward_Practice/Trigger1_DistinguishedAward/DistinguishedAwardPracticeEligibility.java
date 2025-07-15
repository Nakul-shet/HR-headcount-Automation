package DistinguishedAward_Practice.Trigger1_DistinguishedAward;

import HR_Automation_Utilities.*;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;

import static HR_Automation_Utilities.SpotAwardEmailSenderUtility.getCCEmailBasedOnRunType;
import static HR_Automation_Utilities.SpotAwardEmailSenderUtility.getToEmailBasedOnRunType;

public class DistinguishedAwardPracticeEligibility {
    public static void main(String[] args) {
        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(new ByteArrayOutputStream()));
        String emailBody = "";
        try {
            emailBody = DistinguishedAwardPracticeEmailBodyBuilderService.buildEligibilityEmailBody();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.setOut(originalOut);
        }
        String sender = SpotAwardConfig.SENDER_ID;
        String subject = "Nominations for the Reward and Recognition (Distinguished Award) - Q" + CommonEmailBodyUtilities.getCurrentQuarter() + " " + java.time.LocalDate.now().getYear();
        SpotAwardEmailSenderUtility.sendEmail(
                getToEmailBasedOnRunType(SpotAwardConfig.localRunFor , "prac-to"),
                getCCEmailBasedOnRunType(SpotAwardConfig.localRunFor , "prac-cc"),
                sender,
                subject,
                emailBody,
                "distinguished_practice"
        );
    }
}
