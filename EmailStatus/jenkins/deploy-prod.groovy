#!groovy

library 'etax-jenkins-libraries'
deployLamda(
    emailReceiver: 'chayakorn.taktuan@scb.co.th',
    environment: 'prod',
    fileName: 'EmailStatus.zip',
    aws_profile: 'etax-prod',
    aws_region: 'ap-southeast-1',
    aws_bucket: 's3-sourcecode-scbs-prod',
    aws_bucket_path: 'deployment-package',
    aws_lambda_func_name: 'scbs-lambda-etax-email-status-prod',
)

