package Utilities.Configuration;

public class MasterConfig {

    public static final String activeEnvironment = "Test-Internal";

    public static AppConfig getDataBasedOnActiveConfig(String env) {
        return switch (env) {
            case "Test-Internal" -> new TestInternalConfig();
            case "Test-HR" -> new TestHRConfig();
            case "Production" -> new ProductionConfig();
            default -> throw new IllegalArgumentException("Unknown environment: " + env);
        };
    }
}
