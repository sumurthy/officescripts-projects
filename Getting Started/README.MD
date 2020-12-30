# Getting started 

## Objects  

## Collections 

## Error handling 

## Basic performance 

### Avoid slow operations in the lopp

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


* `get*` APIs 

Read all the data you need outside of the loop rather than reading it inside of the loop. At times, it is hard to avoid reading inside of loops - in such a case, make sure your loop counts are  not too large or manage them in batches to avoid having to loop through a large data structure. 



## 
