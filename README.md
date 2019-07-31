Export security groups to .xlsx file
======================================

Introduction
------------
Python script that let you import all Security Groups form specific VPC in AWS account to .xlsx file.

Few variables to be filled in before run:
awsAccID = '123456789012' # AWS account number
assumeRole = 'arn:' # ARN of role that you want to assume on account above
vpcId = 'vpc-' # VPC ID from account above
fileName = 'SecurityGroups - AWS ' + awsAccID + '.xlsx' # Name of .xlsx file
tokenSerial = '' # In case of MFA - Serial number of hardware token

Contact
-------
Ariel Syrko


