#!groovy

library 'etax-jenkins-libraries'
deployDotnetLambda(
    emailReceiver: 'chayakorn.taktuan@scb.co.th',
    environment: 'uat',
    fileName: 'EmailNotify.zip',
    aws_profile: 'etax-uat',
    aws_region: 'ap-southeast-1',
    aws_bucket: 's3-sourcecode-scbs-nonprod',
    aws_bucket_path: 'deployment-package',
    aws_lambda_func_name: 'scbs-lambda-etax-email-notify-uat',
)

