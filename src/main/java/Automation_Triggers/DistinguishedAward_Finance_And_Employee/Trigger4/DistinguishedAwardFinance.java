package Automation_Triggers.DistinguishedAward_Finance_And_Employee.Trigger4;

import Utilities.Common.EmailBodyUtilities;
import Utilities.Common.EmailSenderUtilities;
import Utilities.Configuration.AppConfig;
import Utilities.Configuration.MasterConfig;
import Utilities.Configuration.SpotAwardConfig;
import Utilities.Service.CommonEmailBodyBuilderService;

import static Utilities.Common.EmailSenderUtilities.getCCEmailBasedOnRunType;
import static Utilities.Common.EmailSenderUtilities.getToEmailBasedOnRunType;

public class DistinguishedAwardFinance {
    static AppConfig config = MasterConfig.getDataBasedOnActiveConfig(MasterConfig.activeEnvironment);
    public static void main(String[] args) {
        String emailBody = "";
        try {
            emailBody = CommonEmailBodyBuilderService.buildFinanceEmailBodyForDistinguished();
        } catch (Exception e) {
            e.printStackTrace();
        }
        String sender = config.getSenderId();
        String subject = "Q" + EmailBodyUtilities.getCurrentQuarter() + " - Distinguished Performer - Credit $100 - " + java.time.LocalDate.now().getYear();
        EmailSenderUtilities.sendEmail(
                getToEmailBasedOnRunType(config.getLocalRunFor() , "finance-to"),
                getCCEmailBasedOnRunType(config.getLocalRunFor() , "finance-cc"),
                sender,
                subject,
                emailBody,
                "distinguished_practice distinguished_ops"
        );
    }
}
