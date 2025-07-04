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
        File file = new File("./DataFiles/" + SpotAwardConfig.HEADCOUNT_DATA_FILENAME);
        Workbook workbook = Workbook.getWorkbook(file);
        Sheet sheet = workbook.getSheet(1);
        int mergedCellRow = -1;
        int mergedCellCol = -1;
        for (Range range : sheet.getMergedCells()) {
            Cell topLeft = range.getTopLeft();
            if (topLeft.getContents().trim().equalsIgnoreCase(SpotAwardConfig.HEADCOUNT_DATE_TABLE_NAME)) {
                mergedCellRow = topLeft.getRow();
                mergedCellCol = topLeft.getColumn();
                break;
            }
        }
        if (mergedCellRow == -1) {
            workbook.close();
            throw new Exception("Merged cell with 'New Org Headcount' not found.");
        }
        int headerRow = mergedCellRow + 1;
        while (headerRow < sheet.getRows() && sheet.getCell(mergedCellCol, headerRow).getContents().trim().isEmpty()) {
            headerRow++;
        }
        int dataStartRow = headerRow + 1;
        int latestCol = 2;
        while (latestCol + 1 < sheet.getColumns()) {
            Cell next = sheet.getCell(latestCol + 1, dataStartRow);
            if (!next.getContents().trim().isEmpty()) {
                latestCol++;
            } else {
                break;
            }
        }
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

        htmlBuilder.append("<a href='")
                .append(SpotAwardConfig.SHAREPOINT_LINK)
                .append("' style='display: inline-block; ")
                .append("background-color: #0066cc; ")
                .append("color: white; ")
                .append("padding: 12px 25px; ")
                .append("text-decoration: none; ")
                .append("border-radius: 5px; ")
                .append("font-weight: bold; ")
                .append("margin: 10px 0;'>")
                .append("Nominate Employees")
                .append("</a>");

        htmlBuilder.append("<table border='1' style='border-collapse: collapse; width: 50%; border-width: 2px; text-align:center;'>");
        htmlBuilder.append("<tr style='background-color:yellow'>");
        htmlBuilder.append("<th style='padding-left: 10px; border: 2px solid black;'>New Org</th>");
        htmlBuilder.append("<th style='padding-left: 10px; border: 2px solid black;'>")
                .append(java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH))
                .append(" ")
                .append(java.time.LocalDate.now().getYear())
                .append(" Eligibility</th>");
        htmlBuilder.append("</tr>");
        for (int row = dataStartRow; row < sheet.getRows(); row++) {
            String department = sheet.getCell(0, row).getContents().trim();
            String countStr = sheet.getCell(latestCol, row).getContents().trim();
            if (department.isEmpty() && countStr.isEmpty()) break;
            try {
                int headCount = Integer.parseInt(countStr);
                int twoPercent = (int) Math.floor(headCount * SpotAwardConfig.ELIGIBILITY_PERCENTAGE);
                htmlBuilder.append(String.format(
                        "<tr><td style='padding-left: 10px; border: 2px solid black;'>%s</td>" +
                                "<td style='padding-left: 10px; border: 2px solid black;'>%d</td></tr>",
                        department, twoPercent));
            } catch (NumberFormatException e) {
                htmlBuilder.append(String.format(
                        "<tr><td style='padding-left: 10px; border: 2px solid black;'>%s</td>" +
                                "<td style='padding-left: 10px; border: 2px solid black;'>Invalid</td>" +
                                "<td style='padding-left: 10px; border: 2px solid black;'>N/A</td></tr>",
                        department));
            }
        }
        htmlBuilder.append("</table>");

        htmlBuilder.append("</div>");
        htmlBuilder.append(getEmailSignature());
        htmlBuilder.append("</div>");

        htmlBuilder.append("</body></html>");
        workbook.close();
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

        htmlBuilder.append(getEmailSignature());
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

        htmlBuilder.append(getEmailSignature());
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
                .append(getEmailSignature());

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
        htmlBuilder.append(getEmailSignature());
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

    private static String getEmailSignature() {
        return new StringBuilder()
                .append("<div style='border-top: 1px solid #cccccc; padding-top: 15px; margin-top: 20px;'>")
                .append("<p style='margin: 0; line-height: 1.5;'>Thanks & Regards,</p>")
                .append("<p style='margin: 5px 0; line-height: 1.5;'><strong>TGS India HR</strong></p>")
                .append("<img src='data:image/png;base64,")
                .append(getBase64Image("/src/main/resources/signature/TGSSignature1.jpg"))
                .append("' alt='Company Logo' style='width: 510px; height: 55px; margin-bottom: 5px;'><br>")
                .append("<img src='data:image/png;base64,")
                .append(getBase64Image("/src/main/resources/signature/TGSSignature2.png"))
                .append("' alt='Company Logo' style='width: 600px; height: 26px; margin-bottom: 10px;'><br>")
                .append("</div>")
                .toString();
    }

    private static String getBase64Image(String imagePathLocation) {
        try {
            String imagePath = System.getProperty("user.dir") + imagePathLocation;
            byte[] imageBytes = Files.readAllBytes(Paths.get(imagePath));
            return Base64.getEncoder().encodeToString(imageBytes);
        } catch (IOException e) {
            System.err.println("Failed to load signature image: " + e.getMessage());
            return "";
        }
    }
}
