package HR_Automation_Utilities;

public class DistinguishedAwardPracticeEmailBodyBuilderService {
    public static String buildEligibilityEmailBody() throws Exception {

        StringBuilder htmlBuilder = new StringBuilder();

        htmlBuilder.append("<html><body>");

        htmlBuilder.append("<p>Dear All,</p>");

        htmlBuilder.append("<p>Please share the R&R nominations from your respective teams for Quarter ")
                .append(CommonEmailBodyUtilities.getCurrentQuarter())
                .append(", ")
                .append(java.time.LocalDate.now().getYear())
                .append("</p>");

        htmlBuilder.append("<p>Kindly share the nominations as per the <b>New Org</b> structure by clicking the <b>Nominate Employees</b> button below, on or before 28 ")
                .append(java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH))
                .append(" ")
                .append(java.time.LocalDate.now().getYear())
                .append("</p>");

        htmlBuilder.append("<ul style='font-family:Arial,sans-serif; font-size:14px; color:#333;'>")
                .append("<li><strong>Distinguished Performer of the Quarter</strong> ")
                .append("(Individuals who have demonstrated outstanding dedication, achievements, and contributed significantly to TGS success).</li>")
                .append("</ul>");

        htmlBuilder.append(CommonEmailBodyUtilities.nominateEmployeesButton(SpotAwardConfig.ORG_PRACTICE));

        htmlBuilder.append("<p><b>Distinguished Performer Award Eligibility</b></p>");

        htmlBuilder.append("<table border='1' style='border-collapse: collapse; width: auto; border-width: 2px; text-align:center;'>");
        htmlBuilder.append("<tr style='background-color:yellow'>");
        htmlBuilder.append("<th style='padding: 2px 30px; border: 2px solid black;'>New Org</th>");
        htmlBuilder.append("<th style='padding: 2px 30px; border: 2px solid black;'>")
                .append(" Eligibility</th>");
        htmlBuilder.append("</tr>");

        htmlBuilder.append(ExcelUtilities.readHeadCountData());
        htmlBuilder.append("</table>");

        htmlBuilder.append("<ul style='font-family:Arial,sans-serif; font-size:14px; color:#333;'>")
                .append("<p><strong>Please Note:</strong></p>")
                .append("<li>The citations provided for award winners will be shared to all employees during the award announcement.</li>")
                .append("<li>Request you to add details like why this award is bestowed to the individual, highlight the achievements and contributions in Q")
                .append(CommonEmailBodyUtilities.getCurrentQuarter())
                .append(".</li>")
                .append("<li>Two liner citations will not be accepted. Word limit: 100–150 words.</li>")
                .append("</ul>");

        htmlBuilder.append(CommonEmailBodyUtilities.getEmailSignature());
        htmlBuilder.append("</body></html>");
        return htmlBuilder.toString();
    }
}
