.SYNOPSIS
    Generates an HTML report for Windows LAPS deployment status across domain machines.

.DESCRIPTION
    This script queries Active Directory for all computer objects and checks their LAPS status.
    It supports both Legacy LAPS and Windows LAPS (native).
    It differentiates between Windows Clients and Windows Servers and generates a detailed HTML report.

.NOTES
    Requires: Active Directory PowerShell module
    Legacy LAPS Attribute: ms-Mcs-AdmPwdExpirationTime
    Windows LAPS Attribute: msLAPS-PasswordExpirationTime

Make sure your user have the right permissions to query the above attributes.
The script will generate the report in the same folder.

<img width="1669" height="976" alt="image" src="https://github.com/user-attachments/assets/e230d905-09b9-4fee-b9aa-65073f0d30f5" />
