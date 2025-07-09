package HR_Automation_Utilities;

import jxl.Cell;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtilities {
    public static String readHeadCountData() throws Exception {
        StringBuilder htmlBuilder = new StringBuilder();
        File file = new File("./DataFiles/" + SpotAwardConfig.HEADCOUNT_DATA_FILENAME);
        Workbook workbook = Workbook.getWorkbook(file);
        Sheet sheet = workbook.getSheet(0);
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

        for (int row = dataStartRow; row < sheet.getRows(); row++) {
            String department = sheet.getCell(0, row).getContents().trim();
            String countStr = sheet.getCell(latestCol, row).getContents().trim();
            if (department.isEmpty() && countStr.isEmpty()) break;
            try {
                int headCount = Integer.parseInt(countStr);
                int twoPercent = (int) Math.floor(headCount * SpotAwardConfig.ELIGIBILITY_PERCENTAGE);
                htmlBuilder.append(String.format(
                        "<tr><td style='padding: 2px 30px; border: 2px solid black;'>%s</td>" +
                                "<td style='padding: 2px 30px; border: 2px solid black;'>%d</td></tr>",
                        department, twoPercent));
            } catch (NumberFormatException e) {
                htmlBuilder.append(String.format(
                        "<tr><td style='padding: 2px 30px; border: 2px solid black;'>%s</td>" +
                                "<td style='padding: 2px 30px; border: 2px solid black;'>Invalid</td>" +
                                "<td style='padding: 2px 30px; border: 2px solid black;'>N/A</td></tr>",
                        department));
            }
        }
        workbook.close();
        return htmlBuilder.toString();
    }

    public static String readDistinguishedOperationsData(String sheetName) throws BiffException, IOException {

        StringBuilder htmlBuilder = new StringBuilder();

        File file = new File("./DataFiles/" + SpotAwardConfig.OPS_ELIGIBILITY);
        Workbook workbook = Workbook.getWorkbook(file);
        int sheetNumber = sheetName.equalsIgnoreCase("Ops_Spot") ? 0 : 1;
        Sheet sheet = workbook.getSheet(sheetNumber);
        int row = 1;
        while (row < sheet.getRows()) {
            String dept = sheet.getCell(0, row).getContents().trim();
            String elig = sheet.getCell(1, row).getContents().trim();

            if (dept.isEmpty() && elig.isEmpty()) {
                row++;
                continue;
            }

            // Only proceed for non-blank eligibility rows (start of merge group)
            if (!elig.isEmpty()) {
                int rowspan = 1;

                // Check forward to count how many empty eligibility rows follow
                for (int i = row + 1; i < sheet.getRows(); i++) {
                    String nextDept = sheet.getCell(0, i).getContents().trim();
                    String nextElig = sheet.getCell(1, i).getContents().trim();
                    if (!nextElig.isEmpty()) break;
                    if (nextDept.isEmpty() && nextElig.isEmpty()) break;
                    rowspan++;
                }

                // Row styling
                boolean isTotal = dept.equalsIgnoreCase("Total Numbers");
                String style = isTotal
                        ? "background-color:green; text-align: center; color:white;"
                        : "background-color:white;";

                // First row with eligibility
                htmlBuilder.append("<tr style='").append(style).append("'>")
                        .append("<td style='padding: 2px 20px; text-align: left; border: 2px solid black;'>").append(dept).append("</td>")
                        .append("<td rowspan='").append(rowspan).append("' style='padding: 3px 8px; text-align:center; border: 2px solid black;'>").append(elig).append("</td>")
                        .append("</tr>");

                // Next rows without eligibility
                for (int j = 1; j < rowspan; j++) {
                    String nextDept = sheet.getCell(0, row + j).getContents().trim();
                    htmlBuilder.append("<tr style='").append(style).append("'>")
                            .append("<td style='padding: 2px 20px; border: 2px solid black;'>").append(nextDept).append("</td>")
                            .append("</tr>");
                }

                row += rowspan;
            } else {
                // Print single row without eligibility (to handle cases if any are missed)
                String style = "text-align: left;";
                htmlBuilder.append("<tr style='").append(style).append("'>")
                        .append("<td style='padding: 2px 20px; border: 2px solid black;'>").append(dept).append("</td>")
                        .append("</tr>");
                row++;
            }
        }
        workbook.close();
        return htmlBuilder.toString();
    }

    public static String[] getToEmailAddresses(String addressType){

        try {
            File file = new File("./DataFiles/Spot Award-Spoc List.xls");
            Workbook workbook = Workbook.getWorkbook(file);
            Sheet sheet;

            if(addressType.equals("prac-to")){
                sheet = workbook.getSheet(0);
            }else if(addressType.equals("prac-cc")){
                sheet = workbook.getSheet(1);
            }else if(addressType.equals("op-to")){
                sheet = workbook.getSheet(2);
            }else{
                sheet = workbook.getSheet(3);
            }

            List<String> emails = new ArrayList<>();

            // Loop starts from row 1 (index 1) to skip the header
            for (int row = 1; row < sheet.getRows(); row++) {
                Cell emailCell = sheet.getCell(2, row); // Column index 2 (OfficialMail)
                String email = emailCell.getContents().trim();
                if (!email.isEmpty()) {
                    emails.add(email);
                }
            }

            String[] emailArray = emails.toArray(new String[0]);

            for (String e : emailArray) {
                System.out.println(e);
            }

            return emailArray;

        } catch (Exception e) {
            e.printStackTrace();
        }

        return null;
    }

    public static void main(String[] args) {
        getToEmailAddresses("op-cc");
    }
}
