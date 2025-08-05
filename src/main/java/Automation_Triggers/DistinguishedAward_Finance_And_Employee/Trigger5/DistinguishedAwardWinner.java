package Automation_Triggers.DistinguishedAward_Finance_And_Employee.Trigger5;

import Utilities.Common.EmailBodyUtilities;
import Utilities.Common.EmailSenderUtilities;
import Utilities.Configuration.SpotAwardConfig;
import Utilities.Service.CommonEmailBodyBuilderService;
import jxl.read.biff.BiffException;

import java.io.FileWriter;
import java.io.IOException;

import static Utilities.Common.ExcelUtilities.getAwardWinnersEmail;


public class DistinguishedAwardWinner {

    public static void main(String[] args) throws BiffException, IOException {

        String emailBody = "";
        try {
            emailBody = CommonEmailBodyBuilderService.buildEmployeeConfirmationEmailBodyForDistinguished();
        } catch (Exception e) {
            e.printStackTrace();
            }
        String sender = SpotAwardConfig.SENDER_ID;
        String subject = "Finance Confirmation for Reward and Recognition (Distinguished Award) - Q" + EmailBodyUtilities.getCurrentQuarter() + " " + java.time.LocalDate.now().getYear();
        EmailSenderUtilities.sendEmail(
                getAwardWinnersEmail(SpotAwardConfig.FILE_DISTINGUISHED_FINANCE_DATA),
                null,
                sender,
                subject,
                emailBody,
                "distinguished_practice distinguished_ops"
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
