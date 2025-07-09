package SpotAward_Ops.Trigger1_SpotAward_Ops;

import HR_Automation_Utilities.SpotAwardConfig;
import HR_Automation_Utilities.SpotAwardEmailSenderUtility;
import HR_Automation_Utilities.SpotAwardEmailBodyBuilderService;
import HR_Automation_Utilities.SpotAwardOperationsEmailBodyBuilderService;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;
import java.util.Objects;

public class SpotAwardOpsEligibility {
    public static void main(String[] args) {
        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(new ByteArrayOutputStream()));
        String emailBody = "";
        try {
            emailBody = SpotAwardOperationsEmailBodyBuilderService.buildEligibilityEmailBody();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.setOut(originalOut);
        }
        String sender = SpotAwardConfig.SENDER_ID;
        String subject = "Spot Awards " + java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH) + " " + java.time.LocalDate.now().getYear() + " - Operational Support";
        SpotAwardEmailSenderUtility.sendEmail(
            //Objects.requireNonNull(ExcelUtilities.getToEmailAddresses("op-to")),
            //ExcelUtilities.getToEmailAddresses("op-cc"),
                SpotAwardConfig.RECIPIENTS,
                null,
                sender,
                subject,
                emailBody
        );
    }
}
