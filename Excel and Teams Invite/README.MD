# Send Teams invite based on Excel table data

## Video link

[Demo video](https://youtu.be/HyBdx52NOE8)


## Scenario: 

Example scenario:

* A HR recruiter manages the interview schedule of candidates in a Excel file.
* Recruiter needs to send Teams meeting invite out to candidate and interviewers to eligible candidates. The business rules are to select: (a) invites to only those for whom the invite is already not sent as recorded in the Excel column (b) interview date is in the future (no past dates).
* Recruiter needs to update the Excel file with the confirmation that all Teams meetings have been sent for the eligible records. 


## Solution 

The solution has 3 parts: 

1. Office Script to extract data from Excel table based on conditions and returns an array of objects as JSON data. 
1. The data is then sent to Teams **Create a Teams meeting** action to send invites. Send one Teams meeting per instance in the JSON array. 
1. Send the same JSON data to another Office Script to update the status of the invitation. 

## Input Excel

This is the input table being used. 

## Power Automate 

Full Power Automate flow. 

## Office Scripts

The solution uses two Office Scripts. 

1. [Select filtered rows from table as JSON](SelectFilteredRowsFromTableAsJSON.ts)
1. [Mark as invited](MarkAsInvited.ts)


