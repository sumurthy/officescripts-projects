# Top tips to accelarate your learning 

## Tip 1 - Use code completion and IntelliSence  

Per [VS Code document](https://code.visualstudio.com/docs/editor/intellisense), IntelliSense is a general term for various code editing features including: code completion, parameter info, quick info, and member lists. IntelliSense features are sometimes called by other names such as "code completion", "content assist", and "code hinting."

IntelliSense features are powered by TypeScript language service. TypeScript knows how to read the script you write and the APIs that Excel offers. When its possible to provide code completion hints, the IntelliSense suggestions will pop up as you type. If you continue typing characters, the list of members (variables, methods, etc.) is filtered to only include members containing your typed characters. Pressing Tab or Enter will insert the selected member.

You can trigger IntelliSense by pressing `ctrl space` keys or by typing dot character (.) next to a variable or type. IntelliSence is most useful when completing a method. Among other things, the method signature will contain list of arguments it needs (if any), each argument's date type (can accept many types), whether a given argument required or optional, and the return type of the method. 

Use this code completion feature to explore the object model, discover API you need and learn how to use it. 

## Tip 3 - Range value type

Range is the foundational object within Office Script's Excel object model. [Range APIs](https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts) allows access to both data and format that available on the grid and links other key objects within Excel such as worksheets, tables, charts, etc. 

Among other things each cell consists of `value` field. It contains the underlying value of the cell, which may be different from the `text` you see in the cell. Date or time are good example of value being different from the `text` you see in the cell. A cell can contain one of the following types: `string`, `number`, `boolean`. Even error codes are stored as strings. 

When you are using the cell value it is important to tell TypeScript what value you are expecting to get from a cell. For example, I would like to use this function that subtracts two dates and gives me the number of days separating them. The function that helps me do 


## Tip 4 - Understand arguments - object type

## Tip 4 - Optional chaining 

## Tip 5 - Arrays 


