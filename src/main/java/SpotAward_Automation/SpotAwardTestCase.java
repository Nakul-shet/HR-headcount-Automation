package SpotAward_Automation;

import jxl.Cell;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.PrintStream;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.Locale;

public class SpotAwardTestCase {

    public static String getCurrentMonthName() {
        LocalDate currentDate = LocalDate.now();
        return currentDate.getMonth().getDisplayName(TextStyle.FULL, Locale.ENGLISH);
    }

    public static void main(String[] args) {

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        PrintStream printStream = new PrintStream(outputStream);
        PrintStream originalOut = System.out;
        System.setOut(printStream);

        String emailBody = "";

        try {
            File file = new File("head_count_excel.xls");
            Workbook workbook = Workbook.getWorkbook(file);
            Sheet sheet = workbook.getSheet(1);

            int mergedCellRow = -1;
            int mergedCellCol = -1;

            // Step 1: Find the merged cell with "New Org Headcount"
            for (Range range : sheet.getMergedCells()) {
                Cell topLeft = range.getTopLeft();
                if (topLeft.getContents().trim().equalsIgnoreCase("New Org Headcount")) {
                    mergedCellRow = topLeft.getRow();
                    mergedCellCol = topLeft.getColumn();
                    break;
                }
            }

            if (mergedCellRow == -1) {
                System.out.println("Merged cell with 'New Org Headcount' not found.");
                return;
            }

            // Step 2: Find the first non-empty row below the merged cell (table header)
            int headerRow = mergedCellRow + 1;
            while (headerRow < sheet.getRows() && sheet.getCell(mergedCellCol, headerRow).getContents().trim().isEmpty()) {
                headerRow++;
            }

            // Step 3: Data starts one row below the header
            int dataStartRow = headerRow + 1;
            int dataStartCol = 2; // Column C

            // Step 4: Find the latest non-empty column in the first data row
            int latestCol = dataStartCol;
            while (latestCol + 1 < sheet.getColumns()) {
                Cell next = sheet.getCell(latestCol + 1, dataStartRow);
                if (!next.getContents().trim().isEmpty()) {
                    latestCol++;
                } else {
                    break;
                }
            }

            // Step 5: Print the data
            System.out.println("Latest headcount column index: " + latestCol);
            System.out.printf("%-20s | %-10s | %-13s%n", "Department", "Headcount", "Eligibility");
            System.out.println("---------------------+------------------------------");

            StringBuilder htmlBuilder = new StringBuilder();
            htmlBuilder.append("<html><body>");
            htmlBuilder.append("<h5>Spot Award Eligibility</h5>");
            htmlBuilder.append("<p>Below mentioned are the <b>SPOT award Eligibility </b>for the month of ").append(getCurrentMonthName()).append(" 2025 </p>");
            htmlBuilder.append("<p>Kindly share the nominations in the <b>attached format only</b> as per the <b>New Org</b> on or before 28 ").append(getCurrentMonthName()).append(" 2025 </p>");
            htmlBuilder.append("<table border='1' style='border-collapse: collapse; width: 100%;'>");
            htmlBuilder.append("<tr style='background-color:yellow'><th>New Org</th><th>Headcount</th><th>Eligibility</th></tr>");

            for (int row = dataStartRow; row < sheet.getRows(); row++) {
                String department = sheet.getCell(0, row).getContents().trim();
                String countStr = sheet.getCell(latestCol, row).getContents().trim();
                int twoPercent = (int) Math.ceil(Integer.parseInt(countStr) * 0.02);

                if (department.isEmpty() && countStr.isEmpty()) break;

                try {
                    int headCount = Integer.parseInt(countStr);
                    htmlBuilder.append(String.format("<tr><td>%s</td><td>%d</td><td>%d</td></tr>", department, headCount, twoPercent));
                } catch (NumberFormatException e) {
                    htmlBuilder.append(String.format("<tr><td>%s</td><td>Invalid</td><td>N/A</td></tr>", department));
                }
            }

            htmlBuilder.append("</table>");
            htmlBuilder.append("</body></html>");

            emailBody = htmlBuilder.toString();

            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            System.setOut(originalOut);
        }

        String[] recipients =
                {
                    "nshet@teksystems.com",
                    "kmk@teksystems.com",
                    "aytiwari@teksystems.com",
                    "dmaddala@teksystems.com",
                    "shati@teksystems.com"
                };
        String sender = "";
        String subject = "Spot Award Eligibility";
        String attachmentPath = "spot-award-format.xls";

        SpotAwardEmailUtility.sendEmail(recipients, sender, subject, emailBody, attachmentPath);
    }
}
