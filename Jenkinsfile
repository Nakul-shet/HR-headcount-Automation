pipeline {
    agent any
 
    parameters {
        choice(name: 'ENVIRONMENT',
               choices: ['Test-HR', 'Test-Internal'],
               description: 'Select environment')
    }
 
    tools {
        jdk 'Jdk21'
        maven 'Maven3'
    }
 
    stages {
        stage('Checkout repository') {
            steps {
                checkout scm
            }
        }
 
        stage('Update MasterConfig.java with environment') {
            steps {
                sh '''
                sed -i 's|public static final String activeEnvironment = ".*";|public static final String activeEnvironment = "${ENVIRONMENT}";|' src/main/java/Utilities/Configuration/MasterConfig.java
                echo "✅ Updated environment to ${ENVIRONMENT}"
                grep "activeEnvironment" src/main/java/Utilities/Configuration/MasterConfig.java
                '''
            }
        }
 
        stage('Build with Maven') {
            steps {
                sh 'mvn -B package --file pom.xml'
            }
        }
 
        stage('Run SpotAwardPracticeEligibility') {
            steps {
                sh 'mvn -q exec:java -Dexec.mainClass="Automation_Triggers.Practice_SpotAward.Trigger1.SpotAwardPracticeEligibility"'
            }
        }
    }
}
