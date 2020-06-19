# VBA
For VB6, VBscript, and VBA code projects.
Mostly for Excel.

**Wrap Function**

  Wrap Function was designed to add another function around an existing spreadsheet formula in multiple cells. For example, adding an iferror() or a round() to a bunch of disparate formulae.

This was an early script and isn't optimized. Should pull into an array, iterate there, then put the whole thing back.

**Powerpoint Pagination by Sections**

  This was created to add page numbers to a PowerPoint file based on ppt's sections. For booklets etc. where you may add a new page in the middle or beginning and straight through page numbers would now be wrong. This adds a textbox toward the bottom with "Section Name page n of x". Rerunning it clears out the old ones and puts in the new ones. Skips a section called "Default section" so you can have the title page contents etc. excluded.
  
**Option Selection Worksheet Switch**

  This is a worksheet module function which works with two named ranges "Options" and "selections." Changing the selection to one or more cells within the Options range will add any values within which are not in the selections range to the next available cells in the selections range. Changing the selection to cells in the selections range will remove. Selections including cells in these ranges, but also cells outside, will not make these changes.    
    
**Fill missing time series rows**

  This is designed for changing time series data in a worksheet that was saved 'on change' to include the implied datapoints. Designed for timestamps in seconds. Is not yet designed to make changes related to multiple timestamps in the same second. Though, it can handle their presence.
  
**Word Title Fixer**
  Recreating a now lost (can't find Normal.dotm backup) macro which cleans up selection to be useable as a title and filename. Useful for downloading PDF papers where the filename is useless gibberish. Copy the title from the file, but it may be in all caps, have  double spaces  between  all  words, have line returns, special characters, etc.
  
**Estimated End Date**

  Worksheet function designed to take integer hours (from provided range, in the same rows) and add them to the date from the cell above to give the date when that would occur. By default, skips weekend days. Also, a range of cells with dates (without times, currently) to be excluded, such as holidays and personal days off.
