package Utilities.Configuration;

public interface AppConfig {
    String getRunEnvironment();
    String getLocalRunFor();
    double getEligibilityPercentage();
    String getFilePracticeEligibilityData();
    String getFileSpotFinanceData();
    String getFileDistinguishedFinanceData();
    String getFileOpsEligibilityData();
    String getPracticeTableName();
    String getOrgPractice();
    String getOrgOperations();
    String getSenderId();
    String getSenderPassword();
    String getSharepointLinkPractice();
    String getSharepointLinkOperations();
    String[] getRecipients();
}

