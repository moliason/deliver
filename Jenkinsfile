pipeline {
    agent any

    options {
        timestamps()
        disableConcurrentBuilds()
        buildDiscarder(logRotator(numToKeepStr: '20'))
    }

    parameters {
        string(name: 'TARGET_SERIES', defaultValue: '', trim: true, description: 'Redmine 项目名，逗号分隔；留空交付全部项目')
        string(name: 'CONFIG_CREDENTIALS_ID', defaultValue: 'delivery-config', description: 'JSON 配置的 Secret file 凭据 ID')
        string(name: 'MAPPING_CREDENTIALS_ID', defaultValue: 'delivery-mcu-mapping', description: 'MCU 选型 INI 的 Secret file 凭据 ID')
        string(name: 'GIT_CREDENTIALS_ID', defaultValue: 'delivery-git', description: 'Git HTTPS 用户名和密码凭据 ID')
        string(name: 'GIT_TOOL_NAME', defaultValue: 'Default', description: 'Jenkins 中配置的 Git 工具名称')
        string(name: 'GITLAB_TOKEN_ID', defaultValue: 'delivery-gitlab-token', description: 'GitLab API token 的 Secret text 凭据 ID')
        string(name: 'REDMINE_KEY_ID', defaultValue: 'delivery-redmine-key', description: 'Redmine API key 的 Secret text 凭据 ID')
        string(name: 'DINGTALK_APP_ID', defaultValue: '', description: '可选：钉钉 appKey/appSecret 的 Username with password 凭据 ID')
        string(name: 'DINGTALK_UNION_ID', defaultValue: '', description: '可选：钉钉 unionId 的 Secret text 凭据 ID')
        string(name: 'ZIP_PASSWORD_ID', defaultValue: '', description: '可选：ZIP 密码的 Secret text 凭据 ID，需安装 7-Zip')
    }

    environment {
        PYTHONIOENCODING = 'utf-8'
    }

    stages {
        stage('Prepare') {
            steps {
                sh '''
                    python3 -m venv .venv
                    .venv/bin/python -m pip install -r requirements.txt
                '''
            }
        }
        stage('Deliver') {
            steps {
                script {
                    def bindings = [
                        file(credentialsId: params.CONFIG_CREDENTIALS_ID, variable: 'DELIVERY_CONFIG'),
                        file(credentialsId: params.MAPPING_CREDENTIALS_ID, variable: 'DELIVERY_MAPPING'),
                        gitUsernamePassword(credentialsId: params.GIT_CREDENTIALS_ID, gitToolName: params.GIT_TOOL_NAME),
                        string(credentialsId: params.GITLAB_TOKEN_ID, variable: 'GITLAB_TOKEN'),
                        string(credentialsId: params.REDMINE_KEY_ID, variable: 'REDMINE_API_KEY')
                    ]
                    if (params.DINGTALK_APP_ID) {
                        bindings.add(usernamePassword(credentialsId: params.DINGTALK_APP_ID, usernameVariable: 'DINGTALK_CLIENT_ID', passwordVariable: 'DINGTALK_CLIENT_SECRET'))
                        bindings.add(string(credentialsId: params.DINGTALK_UNION_ID, variable: 'DINGTALK_UNION_ID'))
                    }
                    if (params.ZIP_PASSWORD_ID) {
                        bindings.add(string(credentialsId: params.ZIP_PASSWORD_ID, variable: 'ZIP_PASSWORD'))
                    }
                    withCredentials(bindings) {
                        sh '''
                            set +x
                            mkdir -p "build/$BUILD_NUMBER"
                            cp "$DELIVERY_CONFIG" delivery.local.json
                            cp "$DELIVERY_MAPPING" mcu_selection_mapping.local.ini
                            .venv/bin/python deliver.py --check-config
                            status=0
                            .venv/bin/python deliver.py --series "$TARGET_SERIES" \
                                --output "build/$BUILD_NUMBER/workspace" \
                                > "build/$BUILD_NUMBER/deliver_output.txt" 2>&1 || status=$?
                            cat "build/$BUILD_NUMBER/deliver_output.txt"
                            .venv/bin/python package_delivery.py \
                                --source "build/$BUILD_NUMBER/workspace" \
                                --output "build/$BUILD_NUMBER/package" \
                                --log "build/$BUILD_NUMBER/deliver_output.txt" \
                                --report "build/$BUILD_NUMBER/html_report/index.html" || status=$?
                            exit "$status"
                        '''
                    }
                }
            }
        }
    }

    post {
        always {
            archiveArtifacts artifacts: "build/${env.BUILD_NUMBER}/package/**/*,build/${env.BUILD_NUMBER}/deliver_output.txt", allowEmptyArchive: true
            publishHTML(target: [
                allowMissing: true,
                alwaysLinkToLastBuild: true,
                keepAll: true,
                reportDir: "build/${env.BUILD_NUMBER}/html_report",
                reportFiles: 'index.html',
                reportName: 'Delivery Report'
            ])
        }
    }
}
