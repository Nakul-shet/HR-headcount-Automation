package Utilities.Configuration;

public class TestHRConfig implements AppConfig {

    public String getRunEnvironment() { return "local"; }
    public String getLocalRunFor() { return "test"; }
    public double getEligibilityPercentage() { return 0.02; }
    public String getFilePracticeEligibilityData() { return "Practice_Eligibility.xls"; }
    public String getFileSpotFinanceData() { return "Spot Awards"; }
    public String getFileDistinguishedFinanceData() { return "Distinguished Award Q"; }
    public String getFileOpsEligibilityData() { return "Ops_Eligibility.xls"; }
    public String getPracticeTableName() { return "New Org Headcount"; }
    public String getOrgPractice() { return "practice"; }
    public String getOrgOperations() { return "operations"; }
    public String getSenderId() { return "kmk@teksystems.com"; }
    public String getSenderPassword() { return "****"; }
    public String getSharepointLinkPractice() { return "https://allegiscloud.sharepoint.com/...practice"; }
    public String getSharepointLinkOperations() { return "https://allegiscloud.sharepoint.com/...operations"; }
    public String[] getRecipients() { return new String[]{
            "nshet@teksystems.com"
    }; }
    public String spot(){return "spot";}
    public String distinguished(){return "distinguished";}
}
