package SpotAward_Practice.Trigger2_Reminder;

import HR_Automation_Utilities.SpotAwardConfig;
import HR_Automation_Utilities.SpotAwardEmailSenderUtility;
import HR_Automation_Utilities.SpotAwardEmailBodyBuilderService;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;

import static HR_Automation_Utilities.SpotAwardEmailSenderUtility.getCCEmailBasedOnRunType;
import static HR_Automation_Utilities.SpotAwardEmailSenderUtility.getToEmailBasedOnRunType;

public class SpotAwardReminder1 {

    public static void main(String[] args) {
        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(new ByteArrayOutputStream()));
        String emailBody = "";
        try {
            emailBody = SpotAwardEmailBodyBuilderService.buildReminderEmailBody1();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.setOut(originalOut);
        }
        String sender = SpotAwardConfig.SENDER_ID;
        String subject = "Nominate or Regret! SPOT Award Reminder - Spot Awards " + java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH) + " " + java.time.LocalDate.now().getYear();
        SpotAwardEmailSenderUtility.sendEmail(
                getToEmailBasedOnRunType(SpotAwardConfig.localRunFor , "prac-to"),
                getCCEmailBasedOnRunType(SpotAwardConfig.localRunFor , "prac-cc"),
                sender,
                subject,
                emailBody,
                "spot_practice"
        );
    }
}
