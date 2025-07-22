package SpotAward_Practice.Trigger4_Finance;

import Utilities.Configuration.SpotAwardConfig;
import Utilities.Service.SpotAwardPracticeEmailBodyBuilderService;
import Utilities.Common.EmailSenderUtilities;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;

public class SpotAwardFinanceNotify {

    public static void main(String[] args) {
        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(new ByteArrayOutputStream()));
        String emailBody = "";
        try {
            emailBody = SpotAwardPracticeEmailBodyBuilderService.buildFinanceEmailBody();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.setOut(originalOut);
        }
        String sender = SpotAwardConfig.SENDER_ID;
        String subject = "Spot Awards " + java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH) + " " + java.time.LocalDate.now().getYear();
        EmailSenderUtilities.sendEmail(
                SpotAwardConfig.RECIPIENTS,
                null,
                sender,
                subject,
                emailBody,
                "spot_practice"
        );
    }
}
