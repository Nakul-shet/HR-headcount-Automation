package SpotAward_Automation2;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;

public class SpotAwardTestCase {
    public static void main(String[] args) {
        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(new ByteArrayOutputStream()));
        String emailBody = "";
        try {
            emailBody = SpotAwardEligibilityService.buildEligibilityEmailBody();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.setOut(originalOut);
        }
        String sender = SpotAwardConfig.SENDER_ID;
        String subject = "Spot Awards " + java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH) + " " + java.time.LocalDate.now().getYear();
        SpotAwardEmailUtility.sendEmail(
            SpotAwardConfig.RECIPIENTS,
            sender,
            subject,
            emailBody
//            SpotAwardConfig.SPOT_AWARD_FORMAT_FILE
        );
    }
}
