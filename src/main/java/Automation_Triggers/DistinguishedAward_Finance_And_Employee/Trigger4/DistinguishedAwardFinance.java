package Automation_Triggers.DistinguishedAward_Finance_And_Employee.Trigger4;

import Utilities.Common.EmailBodyUtilities;
import Utilities.Common.EmailSenderUtilities;
import Utilities.Configuration.SpotAwardConfig;
import Utilities.Service.CommonEmailBodyBuilderService;

import static Utilities.Common.EmailSenderUtilities.getCCEmailBasedOnRunType;
import static Utilities.Common.EmailSenderUtilities.getToEmailBasedOnRunType;

public class DistinguishedAwardFinance {
    public static void main(String[] args) {
        String emailBody = "";
        try {
            emailBody = CommonEmailBodyBuilderService.buildFinanceEmailBodyForDistinguished();
        } catch (Exception e) {
            e.printStackTrace();
        }
        String sender = SpotAwardConfig.SENDER_ID;
        String subject = "Reward and Recognition (Distinguished Award) - Q" + EmailBodyUtilities.getCurrentQuarter() + " " + java.time.LocalDate.now().getYear();
        EmailSenderUtilities.sendEmail(
                getToEmailBasedOnRunType(SpotAwardConfig.localRunFor , "finance-to"),
                getCCEmailBasedOnRunType(SpotAwardConfig.localRunFor , "finance-cc"),
                sender,
                subject,
                emailBody,
                "distinguished_practice distinguished_ops"
        );
    }
}
