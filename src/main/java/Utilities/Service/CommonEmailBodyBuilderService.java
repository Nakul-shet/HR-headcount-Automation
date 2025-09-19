package Utilities.Service;

import Utilities.Common.EmailBodyUtilities;
import Utilities.Common.ExcelUtilities;
import Utilities.Configuration.SpotAwardConfig;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

public class CommonEmailBodyBuilderService {
    public static String buildFinanceEmailBodyForSpot() throws Exception {
        File file = new File("./DataFiles/EmployeeFinanceData/" + SpotAwardConfig.FILE_SPOT_FINANCE_DATA);

        String previousMonthYear = LocalDate.now().minusMonths(1).getMonth().getDisplayName(TextStyle.FULL, Locale.ENGLISH)
                + " " + LocalDate.now().getYear();

        List<List<String>> tableData = ExcelUtilities.readExcelData(file);
        String htmlTable = EmailBodyUtilities.generateHtmlTable(tableData);

        StringBuilder body = new StringBuilder();

        body.append("Hi Kishore,<br><br>")
                .append("Please credit the Spot Award Amount for the below mentioned Employees. ")
                .append("This is approved by the respective Practice Head for <b>")
                .append(previousMonthYear)
                .append("</b>.<br><br>")
                .append("Please do confirm once done.<br><br>")
                .append(htmlTable)
                .append("<br>")
                .append(EmailBodyUtilities.getEmailSignature());

        return body.toString();
    }
    public static String buildEmployeeConfirmationEmailBodyForSpot(){

        LocalDate today = LocalDate.now();
        int month =today.getMonthValue();
        int year = today.getYear();
        LocalDate statementDate = LocalDate.of(year, month, 10);

        String formattedStatementDate = statementDate.format(DateTimeFormatter.ofPattern("dd MMMM yyyy"));

        StringBuilder htmlBuilder = new StringBuilder();

        htmlBuilder.append("<html><body>");

        htmlBuilder.append("Dear All,<br><br>")
                .append("Hope you are doing good!<br><br>")
                .append("<b><span style='color:blue'>Congratulations</span></b> on winning a <b>SPOT Award</b> for <b>")
                .append(java.time.LocalDate.now().minusMonths(1).getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH))
                .append(" ")
                .append(java.time.LocalDate.now().getYear())
                .append("</b>!<br><br>")
                .append("The award amount of <b>1K</b> has been credited to your <b style= 'background-color: yellow;'>Pluxee card Account</b>. ")
                .append("<b>It will take 24 hours to reflect in your account. Please check accordingly</b> ")
                .append("and revert in case of any discrepancies on or after <b>")
                .append(formattedStatementDate)
                .append("</b>.<br><br>")
                .append("<b>Note:</b><br>")
                .append("1. Reach out to your reporting manager/ L1 managers for the award certificates.<br>")
                .append("2. The SPOT Awards certificates are shared with L1 Managers for your RMs reference.<br>")
                .append("3. Post the certificate in LinkedIn and tag <b>TEKsystems Global Services In India</b>.<br>")
                .append("4. Amount is not included in the Pluxee card or not having the card; ")
                .append("For additional credit related queries, contact <a href='mailto:kiskala@teksystems.com'>kiskala@teksystems.com</a>.<br>");

        htmlBuilder.append("</div>");
        htmlBuilder.append(EmailBodyUtilities.getEmailSignature());
        htmlBuilder.append("</div>");

        htmlBuilder.append("</body></html>");
        return htmlBuilder.toString();
    }

    public static String buildFinanceEmailBodyForDistinguished() throws Exception {
        File file = new File("./DataFiles/EmployeeFinanceData/" + SpotAwardConfig.FILE_DISTINGUISHED_FINANCE_DATA);

        List<List<String>> tableData = ExcelUtilities.readExcelData(file);
        String htmlTable = EmailBodyUtilities.generateHtmlTable(tableData);

        StringBuilder body = new StringBuilder();

        body.append("Hi Kishore,<br><br>")

                .append("Please credit the Distinguished Performer amount for the below-mentioned Employees, this is approved by the respective Practice Head for ")
                .append("<b>Q"+EmailBodyUtilities.getCurrentQuarter()+"-"+LocalDate.now().getYear()+"</b>")
                .append("<br><br>")
                .append("Please do confirm once done.<br><br>")
                .append(htmlTable)
                .append("<br>")
                .append(EmailBodyUtilities.getEmailSignature());

        return body.toString();
    }

    public static String buildEmployeeConfirmationEmailBodyForDistinguished(){

        LocalDate today = LocalDate.now();
        int month = today.getMonthValue();
        int year = today.getYear();
        LocalDate statementDate = LocalDate.of(year, month, 14);

        String formattedStatementDate = statementDate.format(DateTimeFormatter.ofPattern("dd MMMM yyyy"));

        StringBuilder htmlBuilder = new StringBuilder();

        htmlBuilder.append("<html><body>");

        htmlBuilder.append("Dear All,<br><br>")
                .append("Hope you are doing good!<br><br>")
                .append("<b><span style='color:blue'>Congratulations</span></b> on winning a Distinguished Award for <b>Q")
                .append(EmailBodyUtilities.getCurrentQuarter())
                .append(" ")
                .append(java.time.LocalDate.now().getYear())
                .append("</b>!<br><br>")
                .append("The award amount of <b>$100</b> has been credited to your <b style= 'background-color: yellow;'>Pluxee card Account</b>. ")
                .append("<b>It will take 24 hours to reflect in your account. Please check accordingly</b> ")
                .append("and revert in case of any discrepancies on or after <b>")
                .append(formattedStatementDate)
                .append("</b>.<br><br>")
                .append("<b>Note:</b><br>")
                .append("1. Reach out to your reporting manager/ L1 managers for the award certificates.<br>")
                .append("2. The Distinguished Awards certificates are shared with L1 Managers for your RMs reference.<br>")
                .append("3. Post the certificate in LinkedIn and tag <b>TEKsystems Global Services In India</b>.<br>")
                .append("4. Amount is not included in the Pluxee card or not having the card; ")
                .append("For additional credit related queries, contact <a href='mailto:kiskala@teksystems.com'>kiskala@teksystems.com</a>.<br>")
                .append("5. Attached the mail which consists the details and process of Pluxee card in case of the following( new card/KYC related/Deactivated)");


        htmlBuilder.append("</div>");
        htmlBuilder.append(EmailBodyUtilities.getEmailSignature());
        htmlBuilder.append("</div>");

        htmlBuilder.append("</body></html>");
        return htmlBuilder.toString();
    }

}
