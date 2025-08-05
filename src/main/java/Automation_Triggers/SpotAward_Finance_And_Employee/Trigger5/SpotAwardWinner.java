package Automation_Triggers.SpotAward_Finance_And_Employee.Trigger5;

import Utilities.Configuration.SpotAwardConfig;
import Utilities.Service.CommonEmailBodyBuilderService;
import Utilities.Common.EmailSenderUtilities;

import java.io.FileWriter;
import java.io.IOException;

import jxl.read.biff.BiffException;

import static Utilities.Common.ExcelUtilities.getAwardWinnersEmail;


public class SpotAwardWinner {

    public static void main(String[] args) throws BiffException, IOException {

        String emailBody = "";
        try {
            emailBody = CommonEmailBodyBuilderService.buildEmployeeConfirmationEmailBodyForSpot();
        } catch (Exception e) {
            e.printStackTrace();
            }
        String sender = SpotAwardConfig.SENDER_ID;
        String subject = "Finance Confirmation for Spot Awards " + java.time.LocalDate.now().minusMonths(1).getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH) + " " + java.time.LocalDate.now().getYear();
        EmailSenderUtilities.sendEmail(
                getAwardWinnersEmail(SpotAwardConfig.FILE_SPOT_FINANCE_DATA),
                null,
                sender,
                subject,
                emailBody,
                "spot_practice spot_ops"
        );

        if(SpotAwardConfig.runEnvironment.equals("local")){
            clearMessageIdFileLocal();
        }
    }

    public static void clearMessageIdFileLocal() {
        String filePath = System.getProperty("user.dir") + "/src/main/resources/spot_practice_message_id.txt";
        try (FileWriter writer = new FileWriter(filePath, false)) {
            writer.write("");
            System.out.println("spot_practice_message_id.txt has been cleared.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
