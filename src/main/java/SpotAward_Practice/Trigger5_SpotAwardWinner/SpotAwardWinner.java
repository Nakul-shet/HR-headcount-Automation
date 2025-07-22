package SpotAward_Practice.Trigger5_SpotAwardWinner;

import Utilities.Configuration.SpotAwardConfig;
import Utilities.Service.SpotAwardPracticeEmailBodyBuilderService;
import Utilities.Common.EmailSenderUtilities;

import java.io.ByteArrayOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintStream;
import jxl.read.biff.BiffException;

import static Utilities.Service.SpotAwardPracticeEmailBodyBuilderService.getSpotAwardWinnersEmail;

public class SpotAwardWinner {

    public static void main(String[] args) throws BiffException, IOException {

        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(new ByteArrayOutputStream()));
        String emailBody = "";
        try {
            emailBody = SpotAwardPracticeEmailBodyBuilderService.buildEmployeeConfirmationEmailBody();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.setOut(originalOut);
        }
        String sender = SpotAwardConfig.SENDER_ID;
        String subject = "Finance Confirmation for Spot Awards " + java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH) + " " + java.time.LocalDate.now().getYear();
        EmailSenderUtilities.sendEmail(
                getSpotAwardWinnersEmail(),
                null,
                sender,
                subject,
                emailBody,
                "spot_practice"
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
