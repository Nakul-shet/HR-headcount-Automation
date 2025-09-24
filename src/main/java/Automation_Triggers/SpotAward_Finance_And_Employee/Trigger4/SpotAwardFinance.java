package Automation_Triggers.SpotAward_Finance_And_Employee.Trigger4;

import java.time.*;
import java.time.format.TextStyle;
import java.util.*;

import Utilities.Configuration.AppConfig;
import Utilities.Configuration.MasterConfig;
import Utilities.Configuration.SpotAwardConfig;
import Utilities.Service.CommonEmailBodyBuilderService;
import Utilities.Service.SpotAwardPracticeEmailBodyBuilderService;
import Utilities.Common.EmailSenderUtilities;
import static Utilities.Common.EmailSenderUtilities.getCCEmailBasedOnRunType;
import static Utilities.Common.EmailSenderUtilities.getToEmailBasedOnRunType;
public class SpotAwardFinance {
    static AppConfig config = MasterConfig.getDataBasedOnActiveConfig(MasterConfig.activeEnvironment);
    public static void main(String[] args) {
        String emailBody = "";
        try {
            emailBody = CommonEmailBodyBuilderService.buildFinanceEmailBodyForSpot();
        } catch (Exception e) {
            e.printStackTrace();
        }
        String sender = config.getSenderId();
        String subject = "Spot Awards " + LocalDate.now().minusMonths(1).getMonth().getDisplayName(TextStyle.FULL, Locale.ENGLISH) + " " + java.time.LocalDate.now().getYear();
        EmailSenderUtilities.sendEmail(
                getToEmailBasedOnRunType(config.getLocalRunFor() , "finance-to"),
                getCCEmailBasedOnRunType(config.getLocalRunFor() , "finance-cc"),
                sender,
                subject,
                emailBody,
                "spot_practice spot_ops"
        );
    }
}
