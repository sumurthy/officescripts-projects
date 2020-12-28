# Filter Excel Table and get Visible Range as array of objects

This project shows how to use Office Scripts to filter Excel table to get list of rows visible as array of objects. 

## Video link

[![Watch step by step video](v_visible_range_as_objects.jpg)](https://youtu.be/Mv7BrvPq84A "Watch step by step video")

## Scenario: 

Example scenario:

* Apply filter on a column of a table.
* Extract visible range post filter and extract data as array of objects
* Assemble an object with following structure and return for a future action to use. 

Sample JSON sample. Each key represents uniqie value of a table. Each array instance represents the row that is visible when the corresponding filter is applied. 

```json
{
	"Station-1": [{
		"Station": "Station-1",
		"Shift": "Morning",
		"Date": "27-Oct-20",
		"Responsible": "Debra Berger",
		"Reason": ""
	}, {
		"Station": "Station-1",
		"Shift": "Afternoon",
		"Date": "27-Oct-20",
		"Responsible": "Delia Dennis",
		"Reason": ""
	}, {
		"Station": "Station-1",
		"Shift": "Afternoon",
		"Date": "28-Oct-20",
		"Responsible": "Lidia Holloway",
		"Reason": ""
	}],
	"Station-2": [{
		"Station": "Station-2",
		"Shift": "Morning",
		"Date": "27-Oct-20",
		"Responsible": "Gerhart Moller",
		"Reason": ""
	}, {
		"Station": "Station-2",
		"Shift": "Afternoon",
		"Date": "28-Oct-20",
		"Responsible": "Grady Archie",
		"Reason": ""
	}],
	"Station-3": [{
		"Station": "Station-3",
		"Shift": "Morning",
		"Date": "27-Oct-20",
		"Responsible": "Isaiah Langer",
		"Reason": ""
	}]
}
```

## Solution 

The solution has a single Office Script: 

1. Office Script that perfoms the above mentioned transformation and extracts required data. 

## Input Excel file

This is the [input Excel file](Excel-Filter.xlsx) being used. 

## Office Scripts

The solution uses two Office Scripts. 

1. [Excel Filter Visible Range](ExcelFilterVisibleRange.ts)


