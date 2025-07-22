package SpotAward_Ops.Trigger1;

import Utilities.Configuration.SpotAwardConfig;
import Utilities.Common.EmailSenderUtilities;
import Utilities.Service.SpotAwardOperationsEmailBodyBuilderService;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;

import static Utilities.Common.EmailSenderUtilities.getCCEmailBasedOnRunType;
import static Utilities.Common.EmailSenderUtilities.getToEmailBasedOnRunType;

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
        EmailSenderUtilities.sendEmail(
                getToEmailBasedOnRunType(SpotAwardConfig.localRunFor , "op-to"),
                getCCEmailBasedOnRunType(SpotAwardConfig.localRunFor , "op-cc"),
                sender,
                subject,
                emailBody,
                "spot_operations"
        );
    }
}
