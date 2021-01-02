# Getting started 

Office Scripts is written using [TypeScript language](https://www.typescriptlang.org/), which is an open-source language that builds on JavaScript, one of the world’s most used tools, by adding static type definitions. As the website says, `Types` provide a way to describe the shape of an object, providing better documentation, and allowing TypeScript to validate that your code is working correctly. 

The language syntax itself is written using [JavaScript](https://developer.mozilla.org/en-US/docs/Web/JavaScript) with additional typings defined in the script using TypeScript conventions. For the most part, you can think of Office Script as written in JavaScript. It is essential that you understand basics of JavaScript language to begin your Office Scripts journey; though you don't need to be proficient at it to get started. Office Scripts' action recorder makes it easy to understand the script statements by adding code comments and allowing you to follow along.  

## Definining types 

Type declarations help users understand the type of variable they are dealing with. It helps with auto completion of methods and assists in development time quality checks. 
You can fine type declarations in the script in various places including function declaration, variable declaration, IntelliSence definitions, etc. 

Examples: 

* `function main(workbook: ExcelScript.Workbook)`
* `let myRange: ExcelScript.Range;` 
* `function getMaxAmount(range: ExcelScript.Range): number`

You can identify the types easily in the code editor as it appears distinctly usually in a different color. The color `:` usually preceds the type declaration.  

Writing types can be optional in TypeScript, because type inference allows you to get a lot of power without writing additional code. For the most part TypeScript language is good at inferring the types of variables. However, in certain cases, Office Script will require the type declarations be explicitly defined if the language is unable to clearly identify the type. Also, explicit or implicit `any` is not allowed in Office Script.  

## Script's starting point 

Office Scripts' execution begins at the `main` function. A script is a single file containining one of many functions along with type/interface declarations, variable declaration, etc. To follow along the script, begin with `main` function as Excel always invokes the `main` function when you execute any script. The `main` function will always have at least one argument (or parameter) named `workbook`, which is just a variable name identifying the current workbook against which the script is running. One can define additional argument for usage with Power Automation (offline) execution. 

* `function main(workbook: ExcelScript.Workbook)`

Script can be organized into many smaller functions to aid with code re-usability, clarity, etc. Other functions can inside or outside of the main function - but always in the same file. A script is self contained and can only use functions defined in the same file. Scripts cannot invoke or call another Office Script. 


## Objects  

All of the Excel's object model is defined in a hierarchical structure of objects, beginning with workbook object of type `ExcelScript.Workbook`. An object can contain methods, and other objects (rarely) within it. Objects are linked to each other using the methods - an object's method can return another object or collection of objects. Using code editor's IntelliSence (code completion) feature is a great way to explore the object hierarchy. You can also use the [official reference documentation site](https://docs.microsoft.com/en-us/office/dev/scripts/overview/excel) to follow along the relationships among objects. 

### Get object reference 

You can get object reference by fetching the object using `get` method or by retrieving the collection of objects as shown below: 

```ts
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    const sheet = workbook.getActiveWorksheet();
    // If you want to fetch using an ID or key
    const sheet = workbook.getWorksheet('SomeSheetName');
    // Invoke methods on the object 
    sheet.setPosition(0); 
    
    // Get collection of methods
    const tables = sheet.getTables();
    console.log('Total tables in this sheet: ' + tables.length);
}
```

### Check if an object exists, delete and Add

For creating an object, say with a pre-defined name, it is always better to remove similar object that may exist and then add it. You can accompliush that using this pattern:

```ts
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added. 
  let name = "Index";
  // Check if the worksheet already exists. If not, add the the worksheet.
  let sheet = workbook.getWorksheet('Index');
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Call the delete method on the object to remove it. 
    sheet.delete();
  } 
    // Add a blank worksheet. 
  console.log(`Adding the worhseet named  ${name}.`)
    workbook.addWorksheet("Index");
}

```

Alternately, for deleting an object that may or may not exist, use this pattern: 

```ts
    workbook.getWorksheet('Index')?.delete(); // The ? preceding the delete() will ensure that the API is only invoked if the object exists. 
```


## Collections 

Collections are objects such as tables, charts, columns, etc. that can be retrived as an array and iterated over for processing. You can retrieve a collection using corresponding `get` method and process the data in a loop using one of many TypeScript array traversal techniques such as:

* [`for` or `while`](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Guide/Loops_and_iteration)
* [`for..of`](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Statements/for...of)
* [`forEach`](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/forEach), etc.

* [Language basics of arrays](https://developer.mozilla.org/en-US/docs/Learn/JavaScript/First_steps/Arrays)


This script demonstrates how to use collections used in Office Scripts Excel APIs. 

```ts 

function main(workbook: ExcelScript.Workbook) {

  // Get all sheets as a collection. 
  const sheets = workbook.getWorksheets();
  const names = sheets.map ((sheet) => sheet.getName());
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  // Get information from specific sheets within the collection 
  console.log(`First sheet name is: ${names[0]}`);
  if (sheets.length > 1) {
    console.log(`Last sheet's Id is: ${sheets[sheets.length -1].getId()}`);
  }  
  // Color each worksheet with random color
  for (const sheet of sheets) {
    sheet.setTabColor(`#${Math.random().toString(16).substr(-6)}`);
  }
}
```

## Range basics 

Checkout the [Range Basics](../Range%20Basics) before you go further on your journey. 

## Basic performance considerations

### Avoid slow operations in the loop

Certain operations when done inside/around the loop statements such as `for`, `for..of`, `map`, `forEach`, etc. can lead to slow performance. Avoid the following patterns:

* `get*` APIs 

Read all the data you need outside of the loop rather than reading it inside of the loop. At times, it is hard to avoid reading inside of loops - in such a case, make sure your loop counts are  not too large or manage them in batches to avoid having to loop through a large data structure. 


[![Read and write optimization tip](v_perf.jpg)](https://youtu.be/lsR_GvVW3Pg "Read and write optimization tip")

* `console.log` in the loop slows down the script 

```ts
    // Color each cell with random color
    for (let row = 0; row < rows; row++) {
        for (let col = 0; col < cols; col++) {
            range
                .getCell(row, col)
                .getFormat()
                .getFill()
                .setColor(`#${Math.random().toString(16).substr(-6)}`);
            /* Avoid such console.log in side loop */    
            // console.log("Updating" + range.getCell(row, col).getAddress()); 
        }
    }
```

* `try {} catch ()` statement 

Avoid exception handling for loops. Both inside and outside loops. 

## Error handling 

### Do pre-checks at the beginning of the script

As a best practice, always ensure that all your inputs are present in the Excel file prior to running your script. You may have made certian assumptions about objects being present in the workbook. If those objects don't exist, your script may encouter an error when you read the object or its data. Rather than beginning the processing and erroring in the middle after part of the updates/processing has already finished, it is better to do all pre-checks at the start of the script. 

Example: 

The following script requires two tables by the name of Table1 and Table2 to be present. Hence it checks their presence and ends the script with `return` statement if they are not present with appropriate message. 


```ts
function main(workbook: ExcelScript.Workbook) {

  // Tables that should be part of the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);


  if (!targetTable || !sourceTable) {
    console.log(`Tables missing - Check to make sure both source (${TargetTableName}) and target table (${SourceTableName}) are present before running the script. `);
    return;
  }
  // Continue....

```

If the verification to ensure input data is present or not is happening in a separate function, it is important to end the script by issueing `return` statement from the `main` function. 

For example, `main` function calls `inputPresent` function to do the pre-checks. The `inputPresent` just returns a boolean (true or false) indicating whether all required inputs are present or not. It is the then the responsibility of the `main` function to issue `return` statement (from within same the `main` function) to end the script immediately. 


```ts
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects
  if (!inputPresent(workbook)) {
    return;
  }
  // Continue....

}

function inputPresent( workbook: ExcelScript.Workbook): boolean {

  // Tables that should be part of the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    console.log(`Tables missing - Check to make sure both source (${TargetTableName}) and target table (${SourceTableName}) are present before running the script. `);
    return false;
  }
  return true; 

}
```

### When to abort the script - applicable for scripts used in Power Automation flow

If your script is running as part of Power Automation, you may wish to abort the flow if certain conditions are not met. It is therefore important to not `return` upon an error but rather issue `throw` statement to abort the script so that following step doesn't run. 

```ts
function main(workbook: ExcelScript.Workbook) {

  // Tables that should be part of the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);


  if (!targetTable || !sourceTable) {
    // Abort script
    throw `Tables missing - Check to make sure both source (${TargetTableName}) and target table (${SourceTableName}) are present before running the script. `;
  }
  
```


### When to use try..catch (throw exception)

For the most part, you don't need to use [`try..catch`](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Statements/try...catch) in your script. This is because the script's purpose here is to inform the user that script failed to run due to the absence of input data. Returning with an error message is sufficient in most cases and it lot simpler to simply `return` out of the `main` function. 

However under certain conditions, it may be important to check the return value of an API to verify that it completed successfully. 

Example: consider following snippet -

```ts
function main(workbook: ExcelScript.Workbook) {

  // somewhere in the script, perform a large data update
  range.setValues(someLargeValues);

}
```

The `setValues()` may fail resulting in the script failure. You may wish to handle this condition in your code and perhaps customize the error message or break-up the update into smaller units, etc. In that case, it is important to know that the API returned an error and interpret or handle that error. 

```ts
    try {
        range.setValues(someLargeValues);
    } catch (error) {
        console.log(`The script failed to update the values at location ____. Please inspect and run again.`);        
        console.log(error);      
        return; // End script (assuming this is in main function)  
    }

    // or...

    try {
        range.setValues(someLargeValues);
    } catch (error) {
        console.log(`The script failed to update the values at location ____. Trying a different approach`);        
        handleUpdatesInSmallerChunks(someLargeValues);
    }
    // Continue

```

