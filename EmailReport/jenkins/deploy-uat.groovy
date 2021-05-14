#!groovy

library 'etax-jenkins-libraries'
deployLamda(
    emailReceiver: 'chayakorn.taktuan@scb.co.th',
    environment: 'uat',
    fileName: 'EmailReport.zip',
    aws_profile: 'etax-uat',
    aws_region: 'ap-southeast-1',
    aws_bucket: 's3-sourcecode-scbs-nonprod',
    aws_bucket_path: 'deployment-package',
    aws_lambda_func_name: 'scbs-lambda-etax-email-report-uat',
)

