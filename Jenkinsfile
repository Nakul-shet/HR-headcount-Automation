pipeline {
    agent any
 
    parameters {
        choice(name: 'ENVIRONMENT',
               choices: ['Test-HR', 'Test-Internal'],
               description: 'Select environment')

        choice(name: 'TEST_TO_RUN',
               choices: [
                   'Spot Award Eligibility - Practice',
                   'Spot Award Eligibility - Operations',
                   'Spot Award Reminder 1 - Practice',
                   'Spot Award Reminder 1 - Operations',
		           'Spot Award Reminder 2 - Practice',
                   'Spot Award Reminder 2 - Operations',

                   'Spot Award Finance - Practice & Operations',
                   'Spot Award Final Mail - Employees'

                    'Distinguished Award Eligibility - Practice',
                    'Distinguished Award Eligibility - Operations',
                    'Distinguished Award Reminder 1 - Practice',
                    'Distinguished Award Reminder 1 - Operations',
                    'Distinguished Award Reminder 2 - Practice',
                    'Distinguished Award Reminder 2 - Operations',

                    'Distinguished Award Finance - Practice & Operations',
                    'Distinguished Award Final Mail - Employees'
               ],
               description: 'Select which test/automation to run')
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
                sh """
                    sed -i 's|public static final String activeEnvironment = ".*";|public static final String activeEnvironment = "${ENVIRONMENT}";|' src/main/java/Utilities/Configuration/MasterConfig.java
                    echo "✅ Updated environment to ${ENVIRONMENT}"
                    echo "✅ Selected test to run: ${TEST_TO_RUN}"
                    grep "activeEnvironment" src/main/java/Utilities/Configuration/MasterConfig.java
                """
            }
        }

        stage('Build with Maven') {
            steps {
                sh 'mvn -B package --file pom.xml'
            }
        }

        stage('Run Selected Test') {
            steps {
                script {
                    def mainClassMap = [
                        'Spot Award Eligibility - Practice': 'Automation_Triggers.Practice_SpotAward.Trigger1.SpotAwardPracticeEligibility',
                        'Spot Award Eligibility - Operations': 'Automation_Triggers.Ops_SpotAward.Trigger1.SpotAwardOpsEligibility',
                        'Spot Award Reminder 1 - Practice': 'Automation_Triggers.Practice_SpotAward.Trigger2.SpotAwardPracticeReminder1',
                        'Spot Award Reminder 1 - Operations': 'Automation_Triggers.Ops_SpotAward.Trigger2.SpotAwardOpsReminder1',
                        'Spot Award Reminder 2 - Practice': 'Automation_Triggers.Practice_SpotAward.Trigger3.SpotAwardPracticeReminder2',
                        'Spot Award Reminder 2 - Operations': 'Automation_Triggers.Ops_SpotAward.Trigger3.SpotAwardOpsReminder2'

                        'Distinguished Award Eligibility - Practice' : 'Automation_Triggers.Practice_DistinguishedAward.Trigger1.DistinguishedAwardPracticeEligibility',
                        'Distinguished Award Eligibility - Operations' : 'Automation_Triggers.Ops_DistinguishedAward.Trigger1.DistinguishedAwardOpsEligibility',
                        'Distinguished Award Reminder 1 - Practice' : 'Automation_Triggers.Practice_DistinguishedAward.Trigger2.DistinguishedAwardPracticeReminder1',
                        'Distinguished Award Reminder 1 - Operations' : 'Automation_Triggers.Ops_DistinguishedAward.Trigger2.DistinguishedAwardOpsReminder1',
                        'Distinguished Award Reminder 2 - Practice' : 'Automation_Triggers.Practice_DistinguishedAward.Trigger3.DistinguishedAwardPracticeReminder2',
                        'Distinguished Award Reminder 2 - Operations' : 'Automation_Triggers.Ops_DistinguishedAward.Trigger3.DistinguishedAwardOpsReminder2',
                    ]

                    def selectedMainClass = mainClassMap[params.TEST_TO_RUN]

                    if (selectedMainClass) {
                        echo "🚀 Running ${params.TEST_TO_RUN} on ${params.ENVIRONMENT} environment"
                        sh "mvn -q exec:java -Dexec.mainClass=\"${selectedMainClass}\""
                        echo "✅ Completed execution of ${params.TEST_TO_RUN}"
                    } else {
                        error "❌ Unknown test selection: ${params.TEST_TO_RUN}"
                    }
                }
            }
        }
    }

    post {
        always {
            echo "Pipeline completed for Environment: ${params.ENVIRONMENT}, Test: ${params.TEST_TO_RUN}"
        }
        success {
            echo "✅ Pipeline executed successfully!"
        }
        failure {
            echo "❌ Pipeline failed. Please check the logs."
        }
    }
}


