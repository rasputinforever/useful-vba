# useful-vba
Some handy scripts for VBA

## Procedurally go through many excel files and do the same task on each file
Preparation: Create a list of filepaths, their complete filepaths. This range will be referenced in the For Each loop
Note: All changes are the same so keep that in mind in terms of how similar the files are to one another.

### On "Source" variables
1. You may have one source range by moving the "set source" above For Each so that isn't being re-called every loop.
2. You may have source items, such as wb and sht, = targets as well if it is local source
3. You may not need source items if you are not doing "this = that" 
