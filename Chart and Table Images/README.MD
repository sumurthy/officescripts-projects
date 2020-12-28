# Use Office Scripts Email Chart and Table Images

This project shows how to use Office Scripts and Power Automate actions to re-calculate, create chart and extract images as base-64 string to be used later to send Email.

## Video link
[![Watch step by step video](v_charts_image.jpg)](https://youtu.be/152GJyqc-Kw "Watch step by step video")

## Scenario: 

Example scenario:

* Calculate to get latest results
* Create Chart 
* Get Chart and Table images
* Email the images to recipient using Power Automate action

## Screenshots

#### Input data
![Input data](Input-Data.png) 

### Output chart 
![Chart created](Chart-Created.png) 

## Email that was received through Power Automate flow 
![Email received](Email-Received.png) 

## Solution 

The solution has 2 parts: 

1. Office Script to calculate and extract Excel chart and tables. 
1. A power Automate flow to invoke script and email the results. 

## Input Excel file
Download the input Excel file used in this sample and try it out yourself! 

This is the [input Excel file](Email-Chart-Table.xlsx) being used. 

## Office Scripts

The solution uses two Office Scripts. 

1. [Report Images](ReportImages.ts)


