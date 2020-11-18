# Output Excel table data as array of objects (as JSON) for usage in Power Automate 

Often it is beneficial to extract Excel table data as array of objects (each item representing a row) in the form of a JSON using Office Scripts. This helps with extracting the data from Excel in the same format that is visible to the user. Columns such as date and date-time can then be fed into other systems using Power Automate flows. 


## Input table data
![Input table](Table-Input.png) 

## Output - Excel table data as JSON: array of rows 

```json
[{
	"Event ID": "E107",
	"Date": "2020-12-10",
	"Location": "Montgomery",
	"Capacity": "10",
	"Speakers": "Debra Berger"
}, {
	"Event ID": "E108",
	"Date": "2020-12-11",
	"Location": "Montgomery",
	"Capacity": "10",
	"Speakers": "Delia Dennis"
}, {
	"Event ID": "E109",
	"Date": "2020-12-12",
	"Location": "Montgomery",
	"Capacity": "10",
	"Speakers": "Diego Siciliani"
}, {
	"Event ID": "E110",
	"Date": "2020-12-13",
	"Location": "Boise",
	"Capacity": "25",
	"Speakers": "Gerhart Moller"
}, {
	"Event ID": "E111",
	"Date": "2020-12-14",
	"Location": "Salt Lake City",
	"Capacity": "20",
	"Speakers": "Grady Archie"
}, {
	"Event ID": "E112",
	"Date": "2020-12-15",
	"Location": "Fremont",
	"Capacity": "25",
	"Speakers": "Irvin Sayers"
}, {
	"Event ID": "E113",
	"Date": "2020-12-16",
	"Location": "Salt Lake City",
	"Capacity": "20",
	"Speakers": "Isaiah Langer"
}, {
	"Event ID": "E114",
	"Date": "2020-12-17",
	"Location": "Salt Lake City",
	"Capacity": "20",
	"Speakers": "Johanna Lorenz"
}]
```


## Office Scripts source code

1. [Return table data as array of objects/JSON](TableAsArrayOfObjects.ts)

To suit your needs/data, chage the `interface TableData` structure to match your table columns. Note that for column names that contains spaces, be sure to name your key within quotes as with `"Event ID"` in the sample. 

## Excel files used

See the image and re-create it on your own. 
