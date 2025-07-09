package DistinguishedAward_Practice.Trigger1_DistinguishedAward;

import HR_Automation_Utilities.*;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;

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
//            Objects.requireNonNull(ExcelUtilities.getToEmailAddresses("prac-to")),
//            ExcelUtilities.getToEmailAddresses("prac-cc"),
                SpotAwardConfig.RECIPIENTS,
                null,
                sender,
                subject,
                emailBody
        );
    }
}
