package Utilities.Configuration;

public class MasterConfig {

    public static final String activeEnvironment = "Test-HR";

    public static AppConfig getDataBasedOnActiveConfig(String env) {
        return switch (env) {
            case "Test-Internal" -> new TestInternalConfig();
            case "Test-HR" -> new TestHRConfig();
            default -> throw new IllegalArgumentException("Unknown environment: " + env);
        };
    }
}
