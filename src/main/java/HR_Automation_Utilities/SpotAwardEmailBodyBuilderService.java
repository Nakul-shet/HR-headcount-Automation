package HR_Automation_Utilities;

import jxl.Cell;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;
import java.util.Locale;

public class SpotAwardEmailBodyBuilderService {
    public static String buildEligibilityEmailBody() throws Exception {

        StringBuilder htmlBuilder = new StringBuilder();

        htmlBuilder.append("<html><body>");

        htmlBuilder.append("<p>Dear All,</p>");
        htmlBuilder.append("<p>Below mentioned are the <b>SPOT award Eligibility </b>for the month of ")
                .append(java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH))
                .append(" ")
                .append(java.time.LocalDate.now().getYear())
                .append("</p>");
        htmlBuilder.append("<p>Kindly share the nominations as per the <b>New Org</b> structure by clicking the <b>Nominate Employees</b> button below, on or before 28 ")
                .append(java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH))
                .append(" ")
                .append(java.time.LocalDate.now().getYear())
                .append("</p>");

        htmlBuilder.append(CommonEmailBodyUtilities.nominateEmployeesButton(SpotAwardConfig.ORG_PRACTICE));

        htmlBuilder.append("<table border='1' style='border-collapse: collapse; width: auto; border-width: 2px; text-align:center;'>");
        htmlBuilder.append("<tr style='background-color:yellow'>");
        htmlBuilder.append("<th style='padding: 2px 30px; border: 2px solid black;'>New Org</th>");
        htmlBuilder.append("<th style='padding: 2px 30px; border: 2px solid black;'>")
                .append(java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH))
                .append(" ")
                .append(java.time.LocalDate.now().getYear())
                .append(" Eligibility</th>");
        htmlBuilder.append("</tr>");


        htmlBuilder.append(ExcelUtilities.readHeadCountData());
        htmlBuilder.append("</table>");
        htmlBuilder.append(CommonEmailBodyUtilities.getEmailSignature());
        htmlBuilder.append("</body></html>");
        return htmlBuilder.toString();
    }

    public static String buildReminderEmailBody1() throws Exception {

        StringBuilder htmlBuilder = new StringBuilder();

        htmlBuilder.append("<html><body>");
        htmlBuilder.append("<p>Dear Managers,</p>");

        htmlBuilder.append("<p>This is your <b style='color : purple;'>gentle-but-not-so-gentle reminder</b> to submit those <b style='color : purple;'>SPOT Award Nominations</b> before 28 ")
                .append(java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH))
                .append(" ")
                .append(java.time.LocalDate.now().getYear())
                .append("</p>");

        htmlBuilder.append("<p>Don’t let those <b style='color : purple;'>silent rockstars go unnoticed</b>. It’s your chance to shine a light on your team’s awesomeness 🌟.</p>");

        htmlBuilder.append("<p>If you've already submitted, <b style='color : purple;'> you’re officially awesome </b> and can proudly ignore this message.</p>");

        htmlBuilder.append("<p>If not – tick tock ⏰… recognition season is calling!</p>");

        htmlBuilder.append("<p>Got questions or need help? Ping the ever-helpful folks at ")
                .append("<b><a href='mailto:TGSHRIndiaOps@teksystems.com'>TGSHRIndiaOps@teksystems.com</a></b>")
                .append("</p>")
                .append("<p>Let the nominations roll in!</p>");

        htmlBuilder.append(CommonEmailBodyUtilities.getEmailSignature());
        htmlBuilder.append("</body></html>");

        return htmlBuilder.toString();

    }

    public static String buildReminderEmailBody2() throws Exception {

        StringBuilder htmlBuilder = new StringBuilder();

        htmlBuilder.append("<html><body>");
        htmlBuilder.append("<p>Dear All,</p>");

        htmlBuilder.append("<p>This is your <b>final, no-kidding, last-chance, curtain-call reminder</b> to submit your <b>SPOT Award nominations</b> before 28 ")
            .append(java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH))
            .append(" ")
            .append(java.time.LocalDate.now().getYear())
            .append("</p>");

        htmlBuilder.append("<p>The nomination window <b>slams shut</b> at the end of the day on 28 ")
            .append(java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH))
            .append(" ")
            .append(java.time.LocalDate.now().getYear())
            .append("</p>");

        htmlBuilder.append("After that, even puppy eyes or “I forgot” won’t work. 😅</p>");

        htmlBuilder.append("<p><b>Still haven’t nominated?</b><br>");
        htmlBuilder.append("Please don’t be that manager whose team says, “Recognition? Never heard of it.”</p>");

        htmlBuilder.append("<p>Let’s not disappoint the unsung heroes quietly saving the day in your team!</p>");

        htmlBuilder.append("<p><b>Already submitted?</b> You’re a legend – please ignore this email and go treat yourself to a coffee. ☕</p>");

        htmlBuilder.append("<p>For any last-minute confusion or friendly SOS, reach out to the award-wielding champs at:<br>");
        htmlBuilder.append("📧 <b><a href='mailto:TGSHRIndiaOps@teksystems.com'>TGSHRIndiaOps@teksystems.com</a></b></p>");

        htmlBuilder.append("<p><b>Let’s make those nominations count</b> (before the HR ops team starts chasing you with memes)! 😄</p>");

        htmlBuilder.append(CommonEmailBodyUtilities.getEmailSignature());
        htmlBuilder.append("</body></html>");

        return htmlBuilder.toString();

    }

    public static String buildFinanceEmailBody() throws Exception {
        File file = new File("./DataFiles/" +SpotAwardConfig.FINANCE_DATA_FILENAME);

        String monthYear = LocalDate.now().getMonth().getDisplayName(TextStyle.FULL, Locale.ENGLISH)
                + " " + LocalDate.now().getYear();

        List<List<String>> tableData = readExcelData(file);
        String htmlTable = generateHtmlTable(tableData);

        StringBuilder body = new StringBuilder();

        body.append("Hi Kishore,<br><br>")
                .append("Please credit the Spot Award Amount for the below mentioned Employees. ")
                .append("This is approved by the respective Practice Head for <b>")
                .append(monthYear)
                .append("</b>.<br><br>")
                .append("Please do confirm once done.<br><br>")
                .append(htmlTable)
                .append("<br>")
                .append(CommonEmailBodyUtilities.getEmailSignature());

        return body.toString();
    }

    private static List<List<String>> readExcelData(File file) throws Exception {
        List<List<String>> data = new ArrayList<>();
        Workbook workbook = Workbook.getWorkbook(file);
        Sheet sheet = workbook.getSheet(0);

        for (int row = 0; row < sheet.getRows(); row++) {
            List<String> rowData = new ArrayList<>();
            for (int col = 0; col < sheet.getColumns(); col++) {
                Cell cell = sheet.getCell(col, row);
                rowData.add(cell.getContents());
            }
            data.add(rowData);
        }
        workbook.close();
        return data;
    }

    private static String generateHtmlTable(List<List<String>> tableData) {
        StringBuilder table = new StringBuilder("<table border='1' cellspacing='0' cellpadding='5'>");

        for (int i = 0; i < tableData.size(); i++) {
            table.append("<tr>");
            for (String cell : tableData.get(i)) {
                table.append(i == 0 ? "<th style='background-color : yellow;'>" : "<td>").append(cell).append(i == 0 ? "</th>" : "</td>");
            }
            table.append("</tr>");
        }
        table.append("</table>");
        return table.toString();
    }

    public static String buildEmployeeConfirmationEmailBody(){


        LocalDate today = LocalDate.now();
        int nextMonth = today.getMonthValue() == 12 ? 1 : today.getMonthValue() + 1;
        int year = today.getMonthValue() == 12 ? today.getYear() + 1 : today.getYear();
        LocalDate statementDate = LocalDate.of(year, nextMonth, 8);

        String formattedStatementDate = statementDate.format(DateTimeFormatter.ofPattern("dd MMMM yyyy"));

        StringBuilder htmlBuilder = new StringBuilder();

        htmlBuilder.append("<html><body>");

        htmlBuilder.append("Dear All,<br><br>")
                .append("Hope you are doing good!<br><br>")
                .append("<b><span style='color:blue'>Congratulations</span></b> on winning a SPOT award for <b>")
                .append(java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH))
                .append(" ")
                .append(java.time.LocalDate.now().getYear())
                .append("</b>!<br><br>")
                .append("The award amount of 1K has been credited to your Salary Account. ")
                .append("<b>It will take 24 hours to reflect in your bank account. Please check the bank statement accordingly</b> ")
                .append("and revert in case of any discrepancies on or after <b>")
                .append(formattedStatementDate)
                .append("</b>.<br><br>")
                .append("<b>Note:</b><br>")
                .append("1. Reach out to your reporting manager/ L1 managers for the award certificates.<br>")
                .append("2. The SPOT awards certificates are shared with L1 Managers for your RMs reference.<br>")
                .append("3. Post the certificate in LinkedIn and tag <b>TEKsystems Global Services In India</b>.<br>")
                .append("4. Amount is not included in the salary; it is credited to your salary account separately. ")
                .append("For additional credit related queries, contact <a href='mailto:kiskala@teksystems.com'>kiskala@teksystems.com</a>.<br>");

        htmlBuilder.append("</div>");
        htmlBuilder.append(CommonEmailBodyUtilities.getEmailSignature());
        htmlBuilder.append("</div>");

        htmlBuilder.append("</body></html>");
        return htmlBuilder.toString();
    }

    public static String[] getSpotAwardWinnersEmail() throws BiffException, IOException {
        File file = new File("./DataFiles/" + SpotAwardConfig.FINANCE_DATA_FILENAME);
        Workbook workbook = Workbook.getWorkbook(file);
        Sheet sheet = workbook.getSheet(0);
        int mailIdColumn = -1;
        Cell[] headerRow = sheet.getRow(0);
        for (int i = 0; i < headerRow.length; i++) {
            if (headerRow[i].getContents().equalsIgnoreCase("Mailid")) {
                mailIdColumn = i;
                break;
            }
        }
        if (mailIdColumn == -1) {
            System.out.println("Mailid column not found.");
        }
        List<String> emailList = new ArrayList<>();
        for (int row = 1; row < sheet.getRows(); row++) {
            Cell cell = sheet.getCell(mailIdColumn, row);
            String email = cell.getContents().trim();
            if (!email.isEmpty()) {
                emailList.add(email);
            }
        }
        workbook.close();
        String[] emailArray = emailList.toArray(new String[0]);
        return emailArray;
    }
}
