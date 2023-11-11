# MSOfficeMacros
Visual basic macros and configurations for MS Office Productivity
# A Repository for my MS Word Macros
## List of MS Word Macros
1. Remove CR/LF/CRLF from MS Word Tables
2. Insert an ISO formatted date - time stamp
3. Create a new worklog word document and store it in a file in the working directory
## List of assumptions
1. The normal template contains a standard module named Macros.vb 
2. The normal template contains a class module named DocFormats.cls
# excel_macros
## Current Excel Macros
### 1.Doors Import Macros and Template
  - Map Macros To Keys: Macro to link the DOORS macros to keyboard shortcuts as follows: 
     - Ctrl + Shift + U = MergeCellsLeft
     - Ctrl + Shift + A = BreakoutOneWordHeader
     - Ctrl + Shift + Q = MergeCellsUp
     - Ctrl + Shift + O = AppendColon
     - Ctrl + Shift + E = CarryDown
  - Merge Cells Left: Macro to merge cells to the left where the first and second cell contents gets concatinated into the first cell and The remaining cells shift left; the second cell is deleted.
  - Breakout a one word header: Macro to pull the first word from a cell into its own row. The existing row gets pushed down one and loses the first word.This effectively pull the header text out of the beginning of a cell
  - Merge Cells Up: Macro to merge the words of the currently selected cell and cell below it then delete the row below effectively merging the text up
  - AppendColon: Macro to add a colon at the end of a cell's text. This is helpful for formatting
  - CarryDown: Macro to enter a formula to copy the value of the cell above, effectively carrying the value down one cell.


#Issues / Resolution
## Issue #1: DOORsModuleGrouping
Created branch "DOORsImportMacros/AddMerge" to add the new method
