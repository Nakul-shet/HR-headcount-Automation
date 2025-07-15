package DistinguishedAward_Ops.Trigger1_DistinguisedAward;

import HR_Automation_Utilities.*;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;
import java.util.Objects;

import static HR_Automation_Utilities.SpotAwardEmailSenderUtility.getCCEmailBasedOnRunType;
import static HR_Automation_Utilities.SpotAwardEmailSenderUtility.getToEmailBasedOnRunType;

public class DistinguishedAwardOpsEligibility {
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
        String sender = SpotAwardConfig.SENDER_ID;
        String subject = "Nominations for the Reward and Recognition (Distinguished Award) - Q" + CommonEmailBodyUtilities.getCurrentQuarter() + " " + java.time.LocalDate.now().getYear();
        SpotAwardEmailSenderUtility.sendEmail(
                getToEmailBasedOnRunType(SpotAwardConfig.localRunFor , "op-to"),
                getCCEmailBasedOnRunType(SpotAwardConfig.localRunFor , "op-cc"),
                sender,
                subject,
                emailBody,
                "distinguished_operations"
        );
    }
}
