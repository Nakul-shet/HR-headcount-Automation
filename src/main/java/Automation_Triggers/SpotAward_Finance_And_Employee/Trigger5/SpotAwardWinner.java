package Automation_Triggers.SpotAward_Finance_And_Employee.Trigger5;

import Utilities.Configuration.SpotAwardConfig;
import Utilities.Service.CommonEmailBodyBuilderService;
import Utilities.Common.EmailSenderUtilities;

import java.io.FileWriter;
import java.io.IOException;

import jxl.read.biff.BiffException;

import static Utilities.Common.EmailSenderUtilities.getCCEmailBasedOnRunType;
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
        String subject = "Congratulations | Updates | SPOT Award | " + java.time.LocalDate.now().minusMonths(1).getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH) + " " + java.time.LocalDate.now().getYear();
        EmailSenderUtilities.sendEmail(
                getAwardWinnersEmail(SpotAwardConfig.FILE_SPOT_FINANCE_DATA),
                getCCEmailBasedOnRunType(SpotAwardConfig.localRunFor , "finance-cc"),
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
        String basePath = System.getProperty("user.dir") + "/src/main/resources/";
        String[] fileNames = {
                "spot_practice_message_id.txt",
                "spot_operations_message_id.txt"
        };

        for (String fileName : fileNames) {
            String filePath = basePath + fileName;
            try (FileWriter writer = new FileWriter(filePath, false)) {
                writer.write("");
                System.out.println(fileName + " has been cleared.");
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

}
