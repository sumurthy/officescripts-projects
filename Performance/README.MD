# Read and write performance optimization

This script writes rows of a range in chunks so that if a large range needs to be updated, it can be done in smaller chunks. 

It selects 1K cells per chunk -- but you can override to test out how it works for you. 

* [Updating rows in chunks](UpdatingRowsInChunks.ts)
