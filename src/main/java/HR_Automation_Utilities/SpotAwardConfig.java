package HR_Automation_Utilities;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class SpotAwardConfig {
    public static final String runEnvironment = "local";
    public static final double ELIGIBILITY_PERCENTAGE = 0.02;
    public static final String HEADCOUNT_DATA_FILENAME = "head_count_excel.xls";
    public static final String FINANCE_DATA_FILENAME = "SpotAwardDataFinance.xls";
    public static final String HEADCOUNT_DATE_TABLE_NAME = "New Org Headcount";
    public static final String SENDER_ID = "TGSHRIndiaOps@teksystems.com";
    public static final String SENDER_PASSWORD = "****";
    public static final String SHAREPOINT_LINK = "https://allegiscloud.sharepoint.com/teams/TEKGlobalServices-TGS_INDIA_HR/Shared%20Documents/Forms/AllItems.aspx?id=%2Fteams%2FTEKGlobalServices%2DTGS%5FINDIA%5FHR%2FShared%20Documents%2FTGS%20INDIA%20DATA%2FRewards%20and%20Recognition%20SharePoint%2FNomination%20Sheet&viewid=7fc6f4be%2D249f%2D428e%2D8bd8%2Dc903a32a8cc7&noAuthRedirect=1&startedResponseCatch=true&ovuser=371cb917%2Db098%2D4303%2Db878%2Dc182ec8403ac%2Cnshet%40teksystems%2Ecom&OR=Teams%2DHL&CT=1751975428460&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiI0OS8yNTA2MTIxNjQyMSIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D";
    public static final String[] RECIPIENTS = {
//       "kmk@teksystems.com",
//        "susk@teksystems.com",
//        "aytiwari@teksystems.com",
//        "dmaddala@teksystems.com",
//        "shati@teksystems.com",
//        "sumanvekar@teksystems.com",
//        "raakki@teksystems.com",
//        "kshafiuddin@teksystems.com",
//        "spolegar@teksystems.com",
//        "snayak@teksystems.com",
//        "vsada@teksystems.com",
//        "anmenon@teksystems.com",
//        "nshet@teksystems.com"
    };

//    public static String[] getToEmailAddresses(String addressType){
//
//        try {
//            File file = new File("./DataFiles/Spot Award-Spoc List.xls");
//            Workbook workbook = Workbook.getWorkbook(file);
//            Sheet sheet;
//
//            if(addressType.equals("to")){
//                sheet = workbook.getSheet(0);
//            }else{
//                sheet = workbook.getSheet(1);
//            }
//
//            List<String> emails = new ArrayList<>();
//
//            // Loop starts from row 1 (index 1) to skip the header
//            for (int row = 1; row < sheet.getRows(); row++) {
//                Cell emailCell = sheet.getCell(2, row); // Column index 2 (OfficialMail)
//                String email = emailCell.getContents().trim();
//                if (!email.isEmpty()) {
//                    emails.add(email);
//                }
//            }
//
//            String[] emailArray = emails.toArray(new String[0]);
//
//            for (String e : emailArray) {
//                System.out.println(e);
//            }
//
//            return emailArray;
//
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//
//        return null;
//    }

//    public static void main(String[] args) {
//        getToEmailAddresses("cc");
//    }
}
