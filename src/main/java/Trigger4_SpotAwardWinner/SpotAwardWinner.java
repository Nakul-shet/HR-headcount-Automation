package Trigger4_SpotAwardWinner;

import HR_Automation_Utilities.SpotAwardConfig;
import HR_Automation_Utilities.SpotAwardEmailBodyBuilderService;
import HR_Automation_Utilities.SpotAwardEmailSenderUtility;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import jxl.read.biff.BiffException;

import static HR_Automation_Utilities.SpotAwardEmailBodyBuilderService.getSpotAwardWinnersEmail;

public class SpotAwardWinner {

    public static void main(String[] args) throws BiffException, IOException {

        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(new ByteArrayOutputStream()));
        String emailBody = "";
        try {
            emailBody = SpotAwardEmailBodyBuilderService.buildEmployeeConfirmationEmailBody();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.setOut(originalOut);
        }
        String sender = SpotAwardConfig.SENDER_ID;
        String subject = "Finance Confirmation for Spot Awards " + java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH) + " " + java.time.LocalDate.now().getYear();
        SpotAwardEmailSenderUtility.sendEmail(
                getSpotAwardWinnersEmail(),
                sender,
                subject,
                emailBody
        );
    }
}
