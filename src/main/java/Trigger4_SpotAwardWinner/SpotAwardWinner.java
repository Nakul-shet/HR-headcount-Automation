package Trigger4_SpotAwardWinner;

import HR_Automation_Utilities.SpotAwardConfig;
import HR_Automation_Utilities.SpotAwardEmailBodyBuilderService;
import HR_Automation_Utilities.SpotAwardEmailSenderUtility;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.PrintStream;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class SpotAwardWinner {

    public static void main(String[] args) throws BiffException, IOException {

        File file = new File(SpotAwardConfig.FINANCE_DATA_FILENAME);
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
            return;
        }
        List<String> emailList = new ArrayList<>();
        for (int row = 1; row < sheet.getRows(); row++) {
            Cell cell = sheet.getCell(mailIdColumn, row);
            String email = cell.getContents().trim();
            if (!email.isEmpty()) {
                emailList.add(email);
            }
        }

        String[] emailArray = emailList.toArray(new String[0]);
        workbook.close();

        PrintStream originalOut = System.out;
        System.setOut(new PrintStream(new ByteArrayOutputStream()));
        String emailBody = "";
        try {
            emailBody = SpotAwardEmailBodyBuilderService.buildEmployeeConfirmationEmailBody();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.setOut(originalOut);
        }
        String sender = SpotAwardConfig.SENDER_ID;
        String subject = "Finance Confirmation for Spot Awards " + java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH) + " " + java.time.LocalDate.now().getYear();
        SpotAwardEmailSenderUtility.sendEmail(
                emailArray,
                sender,
                subject,
                emailBody
        );
    }
}
