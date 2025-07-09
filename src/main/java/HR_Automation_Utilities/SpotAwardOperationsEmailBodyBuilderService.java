package HR_Automation_Utilities;

import jxl.Cell;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Base64;

public class SpotAwardOperationsEmailBodyBuilderService {
    public static String buildEligibilityEmailBody() throws Exception {


        StringBuilder htmlBuilder = new StringBuilder();

        htmlBuilder.append("<html><body>");

        htmlBuilder.append("<p>Dear Leaders,</p>");
        htmlBuilder.append("<p>Below mentioned are the <b>SPOT Award Eligibility </b>for the month of ")
                .append(java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH))
                .append(" ")
                .append(java.time.LocalDate.now().getYear())
                .append("</p>");
        htmlBuilder.append("<p>Kindly share the nominations as per the <b>Practice</b> structure by clicking the <b>Nominate Employees</b> button below, on or before 28 ")
                .append(java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH))
                .append(" ")
                .append(java.time.LocalDate.now().getYear())
                .append("</p>");

        htmlBuilder.append(CommonEmailBodyUtilities.nominateEmployeesButton(SpotAwardConfig.ORG_OPERATIONS));

        htmlBuilder.append("<table border='1' style='border-collapse: collapse; width: auto; font-family: Arial, sans-serif; background-color: white; color: black;'>");

// Header
        htmlBuilder.append("<tr style=' text-align: center; font-weight: bold; background-color:orange;'>")
                .append("<th style='padding: 2px 20px;'>Department</th>")
                .append("<th style='padding: 2px 20px;'>SPOT Eligibility</th>")
                .append("</tr>");
        htmlBuilder.append(ExcelUtilities.readDistinguishedOperationsData("Ops_Spot"));

        htmlBuilder.append("</table>");

        htmlBuilder.append(CommonEmailBodyUtilities.getEmailSignature());

        htmlBuilder.append("</body></html>");

        return htmlBuilder.toString();
    }

}
