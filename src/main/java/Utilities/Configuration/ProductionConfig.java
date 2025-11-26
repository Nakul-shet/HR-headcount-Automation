package Utilities.Configuration;

public class ProductionConfig implements AppConfig {

    public String getRunEnvironment() { return "local"; }
    public String getLocalRunFor() { return "prod"; }
    public double getEligibilityPercentage() { return 0.02; }
    public String getFilePracticeEligibilityData() { return "Eligibility.xls"; }
    public String getFileSpotFinanceData() { return "Spot Awards"; }
    public String getFileDistinguishedFinanceData() { return "Distinguished Award Q"; }
    public String getFileOpsEligibilityData() { return "Eligibility.xls"; }
    public String getPracticeTableName() { return "New Org Headcount"; }
    public String getOrgPractice() { return "practice"; }
    public String getOrgOperations() { return "operations"; }
    public String getSenderId() { return "TGSHRIndiaOps@teksystems.com"; }
    public String getSenderPassword() { return "****"; }
    public String getSharepointLinkPractice() { return "https://allegiscloud.sharepoint.com/...practice"; }
    public String getSharepointLinkOperations() { return "https://allegiscloud.sharepoint.com/...operations"; }
    public String[] getRecipients() { return new String[]{}; }
    public String spot(){return "spot";}
    public String distinguished(){return "distinguished";}
}
