# Top tips to accelarate your learning 

## Tip 1 - Use code-completion/IntelliSense in the editor 

Per [VS Code document](https://code.visualstudio.com/docs/editor/intellisense), IntelliSense is a general term for various code editing features including: code completion, parameter info, quick info, and member lists. IntelliSense features are sometimes called by other names such as "code completion", "content assist", and "code hinting."

IntelliSense features in Office Scripts are powered by TypeScript language. TypeScript knows how to read & interpret the script you write and the APIs that Excel offers. When its possible to provide code completion hints, the IntelliSense suggestions will pop up as you type. If you continue typing characters, the list of members (variables, methods, etc.) is filtered to only include members containing your typed characters. Pressing Tab or Enter will insert the selected member.

You can trigger IntelliSense by pressing `ctrl space` keys or by typing dot character (.) next to a variable or type. IntelliSence is most useful when completing a method. Among other things, the method signature will contain list of arguments it needs (if any), each argument's date type (can accept many types), whether a given argument required or optional, and the return type of the method. 

Use this code completion feature to explore the object model, discover API you need and learn how to use it. 

## Tip 2 - Type assertion of range values

Range is the foundational object within Office Script's Excel object model. [Range APIs](https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts) allows access to both data and format that available on the grid and links other key objects within Excel such as worksheets, tables, charts, etc. 

Among other things each cell consists of `value` field. It contains the underlying value of the cell, which may be different from the `text` you see in the cell. Date or time are good example of value being different from the `text` you see in the cell. A cell can contain one of the following types: `string`, `number`, `boolean`. Even error codes are stored as strings. 

When you are using the cell value it is important to tell TypeScript what value you are expecting to get from a cell. For example, I would like to use this function that subtracts two dates present in cells and get the number of days separating them. The function that helps me do this expects a number type as its argument. However, when I call this function, I have to tell TypeScript that value being read from the cell using `getValue()` is a number using the `as number` **Type Assertion**. Otherwise, the TypeScript compiler points out the potential mis-match using a warning message (red underline). Let's see how to do that. 

```ts
function main(workbook: ExcelScript.Workbook) {

  const xlDate1value = workbook.getActiveWorksheet().getRange('A2').getValue();
  const xlDate2value = workbook.getActiveWorksheet().getRange('B2').getValue();

  const jsDate1 = xlDateToJSDate(xlDate1value as number); // Tell TypeScript compiler that the value will be a number
  console.log(jsDate1.toUTCString())

  const jsDate2 = xlDateToJSDate(xlDate2value as number); // Tell TypeScript compiler that the value will be a number
  console.log(jsDate2.toUTCString())
}

Similarly, you can use `as string`, `as boolean` **Type Assertions**. 


/**
 * Function to return the JS date from Excel date
 * 
 */
function xlDateToJSDate(serialDate: number): Date {
  var days = Math.floor(serialDate);
  var hours = Math.floor((serialDate % 1) * 24);
  var minutes = Math.floor((((serialDate % 1) * 24) - hours) * 60)
  const returnDate = new Date(Date.UTC(0, 0, serialDate, hours - 17, minutes));
  return returnDate;
}
```

Note that such compile time warnings will not affect the actual running of the script as at runtime you are really passing the numerical argument. However, TypeScript compiler doesn't know the values at runtime and hence it shows the warning in the editor. 

If indeed the cell had a non-numeric value at run-time, the helper function will fail to execute and throw an error. So, if you are unsure about the type of data present then you must do a run-time check before calling this function that expects a numeric argument. Let's see how to do that. 

```ts
function main(workbook: ExcelScript.Workbook) {

  const xlDate1value = workbook.getActiveWorksheet().getRange('A2').getValue();
  const xlDate2value = workbook.getActiveWorksheet().getRange('B2').getValue();

  if (typeof xlDate1value === 'number') {
    const jsDate1 = xlDateToJSDate(xlDate1value as number);
    console.log(jsDate1.toUTCString())
  }

  if (typeof xlDate2value === 'number') {
    const jsDate2 = xlDateToJSDate(xlDate2value as number);
    console.log(jsDate2.toUTCString())
  }
}


/**
 * Function to return the JS date from Excel date
 * 
 */
function xlDateToJSDate(serialDate: number): Date {
  var days = Math.floor(serialDate);
  var hours = Math.floor((serialDate % 1) * 24);
  var minutes = Math.floor((((serialDate % 1) * 24) - hours) * 60)
  const returnDate = new Date(Date.UTC(0, 0, serialDate, hours - 17, minutes));
  return returnDate;
}
```

## Tip 3 - Understand arguments - object type

The Office Scripts methods accepts zero or many arguments depending on the type of the method. `workbook.getActiveWorksheet()` or `workbook.getActiveCell()` doesn't require any argument for them to work; however `workbook.getWorksheet()` method requires the name or ID of the worksheet such as `workbook.getWorksheet('Sheet1')`. The arguments passed to methods can be of many data types such as: 
* String
* Number
* Boolean
* Excel objects as Range, Table, etc. 
* Object parameter 

Then there is a special parameter known as enumerator - which basically allows standardzied argument that Excel accepts for a given method. Example of enumerator (ENUM) argument passed to the method:

* workbook.getActiveCell().clear(ExcelScript.ClearApplyTo.contents)
* workbook.getActiveCell().clear(ExcelScript.ClearApplyTo.formats)
* workbook.getActiveCell().clear(ExcelScript.ClearApplyTo.all)

In the above lines, different standardized enmerator arguments are passed to the range's clear method. The name of the enumerator is `ExcelScript.ClearApplyTo` and it offers options such as `contents1`, `formats` and `all` in the last segment of the enumerator. 

## Tip 4 - `null` or `undefined` check and optional chaining 

The `null` or `undefined` types crop up as part of API return values and can cause quite a bit of headache if not handled properly. Fortunately, TypeScript/JavaScript language offers an easy way to check if a variable is of type undefined or null easily. 

For example, the `findAll` API can either return the range areas where the string "locks" is found or it can return `undefined` (if there is no such string in the worksheet). 

```ts
function main(workbook: ExcelScript.Workbook) {
  const results = workbook.getActiveWorksheet().findAll("Locks", {matchCase: true});
  const cellAddressArray = results.getAddress().split(',');
  console.log(cellAddressArray);
}
```

The `getAddress()` method called on the `results` variable can either work or throw an error such as below:

```
`Line 3: Cannot read property 'getAddress' of undefined`
```
So, it is important to check if the `results` 

Per [docs](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/Optional_chaining), the optional chaining operator (`?.`) permits reading the value of a property located deep within a chain of connected objects without having to expressly validate that each reference in the chain is valid. The ?. operator functions similarly to the . chaining operator, except that instead of causing an error if a reference is nullish (null or undefined), the expression short-circuits with a return value of undefined. When used with function calls, it returns undefined if the given function does not exist.


## Tip 5 - Arrays 


