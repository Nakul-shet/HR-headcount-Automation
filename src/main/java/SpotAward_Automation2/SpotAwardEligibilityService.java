package SpotAward_Automation2;

import jxl.Cell;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Base64;

public class SpotAwardEligibilityService {
    public static String buildEligibilityEmailBody() throws Exception {
        File file = new File(SpotAwardConfig.HEADCOUNT_DATA_FILENAME);
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
        htmlBuilder.append("<p>Kindly share the nominations in the <b>attached format only</b> as per the <b>New Org</b> on or before 28 ")
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

        htmlBuilder.append("<table border='1' style='border-collapse: collapse; width: 50%; border-width: 2px;'>");
        htmlBuilder.append("<tr style='background-color:yellow'>");
        htmlBuilder.append("<th style='padding-left: 10px; border: 2px solid black;'>New Org</th>");
        //htmlBuilder.append("<th style='padding-left: 10px; border: 2px solid black;'>Headcount</th>");
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
//                htmlBuilder.append(String.format(
//                        "<tr><td style='padding-left: 10px; border: 2px solid black;'>%s</td>" +
//                                "<td style='padding-left: 10px; border: 2px solid black;'>%d</td>" +
//                                "<td style='padding-left: 10px; border: 2px solid black;'>%d</td></tr>",
//                        department, headCount, twoPercent));
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

    private static String getEmailSignature() {
        return new StringBuilder()
                .append("<div style='border-top: 1px solid #cccccc; padding-top: 15px; margin-top: 20px;'>")
                .append("<p style='margin: 0; line-height: 1.5;'>Thanks and Regards,</p>")
                .append("<p style='margin: 5px 0; line-height: 1.5;'>")
                .append("<strong>Karthik M K</strong> | Associate Engineer | Global Engineering Application Hub")
                .append("</p>")
                .append("<p style='margin: 5px 0; line-height: 1.5;'>")
                .append("M: <a href='tel:+918078254741' style='color: #0066cc; text-decoration: none;'>+91 8078254741</a>")
                .append("</p>")
                .append("<p style='margin: 5px 0; line-height: 1.5;'><strong>Office:</strong> TEKsystems Global Services Pvt. Ltd.</p>")
                .append("<p style='margin: 5px 0; line-height: 1.5;'>#801, 8B, 8th Floor, Arliga Ecoworld Campus (earlier - RMZ Ecoworld)<br>")
                .append("Outer Ring Road, Devarabeesanahalli<br>")
                .append("Bengaluru - 560 103.</p>")
                .append("<p style='margin: 5px 0; line-height: 1.5;'>")
                .append("Email: <a href='mailto:kmk@teksystems.com' style='color: #0066cc; text-decoration: none;'>kmk@teksystems.com</a>")
                .append("</p>")
                .append("<img src='data:image/png;base64,")
                .append(getBase64Image())
                .append("' alt='Company Logo' style='width: 300px; height: 50px; margin-bottom: 10px;'><br>")
                .append("</div>")
                .toString();
    }

    private static String getBase64Image() {
        try {
            String imagePath = System.getProperty("user.dir") + "/src/main/resources/signature/TGSSignature.png";
            byte[] imageBytes = Files.readAllBytes(Paths.get(imagePath));
            return Base64.getEncoder().encodeToString(imageBytes);
        } catch (IOException e) {
            System.err.println("Failed to load signature image: " + e.getMessage());
            return "";
        }
    }
}
