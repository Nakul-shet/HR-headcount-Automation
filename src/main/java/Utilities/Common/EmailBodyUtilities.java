package Utilities.Common;

import Utilities.Configuration.SpotAwardConfig;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.Base64;
import java.util.List;

public class EmailBodyUtilities {
    public static String nominateEmployeesButton(String orgStructure) {
        String sharepointLink = orgStructure.equalsIgnoreCase("practice") ? SpotAwardConfig.SHAREPOINT_LINK_PRACTICE : SpotAwardConfig.SHAREPOINT_LINK_OPERATIONS;
        return new StringBuilder()
                .append("<!--[if mso]>")
                .append("<table border='0' cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:30px;'><tr><td align='left'>")
                .append("<v:roundrect xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w='urn:schemas-microsoft-com:office:word' href='")
                .append(sharepointLink)
                .append("' style='height:45px;v-text-anchor:middle;width:280px;' arcsize='10%' stroke='f' fillcolor='#0066cc'>")
                .append("<w:anchorlock/>")
                .append("<center style='color:#ffffff;font-family:Arial,sans-serif;font-size:15px;font-weight:bold;'>")
                .append("Nominate Employees")
                .append("</center>")
                .append("</v:roundrect>")
                .append("</td></tr></table>")
                .append("<![endif]-->")
                .append("<!--[if !mso]><!-- -->")
                .append("<table border='0' cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:30px;'><tr><td align='left'>")
                .append("<a href='").append(sharepointLink).append("' target='_blank' ")
                .append("style='background-color:#0066cc;border-radius:5px;color:#ffffff;display:inline-block;font-family:Arial,sans-serif;")
                .append("font-size:15px;font-weight:bold;line-height:45px;text-align:center;text-decoration:none;width:220px;'>")
                .append("Nominate Employees")
                .append("</a>")
                .append("</td></tr></table>")
                .append("<!--<![endif]-->")
                .toString();
    }

    public static String getEmailSignature() {
        return new StringBuilder()
                .append("<div style='border-top: 1px solid #cccccc; padding-top: 15px; margin-top: 20px;'>")
                .append("<p style='margin: 0; line-height: 1.5;'>Thanks & Regards,</p>")
                .append("<p style='margin: 5px 0; line-height: 1.5;'><strong>TGS India HR</strong></p>")
                .append("<img src='data:image/png;base64,")
                .append(getBase64Image("/src/main/resources/Signature/TGSSignature1.jpg"))
                .append("' alt='Company Logo' width='510' height='55' style='margin-bottom: 5px;'><br>")
                .append("<img src='data:image/png;base64,")
                .append(getBase64Image("/src/main/resources/Signature/TGSSignature2.png"))
                .append("' alt='Company Logo' width='600' height='26' style='margin-bottom: 10px;'><br>")
                .append("</div>")
                .toString();
    }

    public static String getBase64Image(String imagePathLocation) {
        try {
            String imagePath = System.getProperty("user.dir") + imagePathLocation;
            byte[] imageBytes = Files.readAllBytes(Paths.get(imagePath));
            return Base64.getEncoder().encodeToString(imageBytes);
        } catch (IOException e) {
            System.err.println("Failed to load signature image: " + e.getMessage());
            return "";
        }
    }

    public static int getCurrentQuarter() {
        int month = LocalDate.now().getMonthValue(); // 1 (Jan) to 12 (Dec)
        if (month >= 1 && month <= 3) {
            return 4; // Q1: Jan - Mar
        } else if (month >= 4 && month <= 6) {
            return 1; // Q2: Apr - Jun
        } else if (month >= 7 && month <= 9) {
            return 2; // Q3: Jul - Sep
        } else {
            return 3; // Q4: Oct - Dec
        }
    }
    public static String generateHtmlTable(List<List<String>> tableData) {
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

}
