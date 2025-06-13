package Trigger3_Finance;

import HR_Automation_Utilities.SpotAwardConfig;
import HR_Automation_Utilities.SpotAwardEmailBodyBuilderService;
import HR_Automation_Utilities.SpotAwardEmailSenderUtility;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;

public class SpotAwardFinanceNotify {

    public static void main(String[] args) {
        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(new ByteArrayOutputStream()));
        String emailBody = "";
        try {
            emailBody = SpotAwardEmailBodyBuilderService.buildFinanceEmailBody();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.setOut(originalOut);
        }
        String sender = SpotAwardConfig.SENDER_ID;
        String subject = "Spot Awards " + java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH) + " " + java.time.LocalDate.now().getYear();
        SpotAwardEmailSenderUtility.sendEmail(
                SpotAwardConfig.RECIPIENTS,
                sender,
                subject,
                emailBody
        );
    }
}
