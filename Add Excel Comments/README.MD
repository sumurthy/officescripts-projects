# Add and manage comments in Excel sheet

Add comments to Excel cell with at-mention of a contact. 

## Video link

[Demo video](https://youtu.be/CpR78nkaOFw)


## Scenario: 

Example scenario:

* Team lead maintains shift schedule information. Team lead assigns employee ID to shift record. 
* Team lead wishes to send send email reminder by adding a comment with @mention of the employee by cross referencing the employee email address along with a custom message from the sheet. 
* Employee can then address/respond to the comment. 

## Solution 

1. Office Script that extracts employee information from the employee sheet. 
1. The script then cross references the shift record and adds comment to the appropriate cell by including relevant employee email, team lead comment. 
1. It also removes any existing comment before adding new comment. 

## Input Excel file
Download the input Excel file used in this sample and try it out yourself! 

This is the [input Excel file](Excel-Comments.xlsx) being used. 

## Office Scripts

The solution uses two Office Scripts. 

1. [Add comment and at-mention](AddComment.ts)


