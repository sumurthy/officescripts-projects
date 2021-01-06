# Read and write performance optimization


## Basic perf optimization 

Please refer to the [getting started](../Getting%20Started#basic-performance-considerations) performance section for the basics of Office Scripts. 

## Large data write scenario

The `setValues()` API of range allows setting the values of a range. This API will have data limitation depending on various factors such as data size, network settings, etc. In order to reliably update a large range of data, you'll need to think about doing data updates in smaller chunks. This script attempts to do this and writes rows of a range in chunks so that if a large range needs to be updated, it can be done in smaller parts. It has not been tested across various sizes... be aware of that if you want to use this in your script. I'll update with findings around how it performs for various data sizes. 

It selects 1K cells per chunk -- but you can override to test out how it works for you. 

* [Updating rows in chunks](UpdatingRowsInChunks.ts): It updates 1k rows with 6 columns of data. Run this on a blank sheet to examine. 

[![Large range data update](v_largedata.png)](https://youtu.be/BP9Kp0Ltj7U)
