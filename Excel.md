Shortcuts

- CTRL plus arrow keys

- Collapse pivot table - Alt+A+H

- Expand- Alt + A + J

  

Questions to ask in class: 

- Why my conditional formating isn't working in the rounded numbers
  - I wasn't able to adjust the rule in 1.2.5 - something isn't working

- How to properly remove filters



Look up syntax for IFS



**VLOOKUP**

=VLOOKUP(A2,$E$4:$G$7,3,FALSE)

- Lookup a cell
- Provide a data range
- Index number

**HLOOKUP**

- Looking for values horizontally
- 

## Syntax

HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])

The HLOOKUP function syntax has the following arguments:

- **Lookup_value**  Required. The value to be found in the first row of the table. Lookup_value can be a value, a reference, or a text string.
- **Table_array**  Required. A table of information in which data is looked up. Use a reference to a range or a range name.
  - The values in the first row of table_array can be text, numbers, or logical values.
  - If range_lookup is TRUE, the values in the first row of table_array must be placed in ascending order: ...-2, -1, 0, 1, 2,... , A-Z, FALSE, TRUE; otherwise, HLOOKUP may not give the correct value. If range_lookup is FALSE, table_array does not need to be sorted.
  - Uppercase and lowercase text are equivalent.
  - Sort the values in ascending order, left to right. For more information, see [Sort data in a range or table](https://support.office.com/en-us/f1/topic/sort-data-in-a-range-or-table-62d0b95d-2a90-4610-a6ae-2e545c4a4654?NS=EXCEL&Version=90).
- **Row_index_num**  Required. The row number in table_array from which the matching value will be returned. A row_index_num of 1 returns the first row value in table_array, a row_index_num of 2 returns the second row value in table_array, and so on. If row_index_num is less than 1, HLOOKUP returns the #VALUE! error value; if row_index_num is greater than the number of rows on table_array, HLOOKUP returns the #REF! error value.
- **Range_lookup**  Optional. A logical value that specifies whether you want HLOOKUP to find an exact match or an approximate match. If TRUE or omitted, an approximate match is returned. In other words, if an exact match is not found, the next largest value that is less than lookup_value is returned. If FALSE, HLOOKUP will find an exact match. If one is not found, the error value #N/A is returned.

# XLOOKUP function				 

Use the **XLOOKUP** function to find things in a table or range by row. For example, look up the price of an automotive part by the part number, or find an employee name based on their employee ID. With XLOOKUP, you can look in one column for a search term, and return a result from the same row in another column, regardless of which side the return column is on.

## Syntax

The XLOOKUP function searches a range or an array, and then returns the item corresponding to the first match it finds. If no match exists, then XLOOKUP can return the closest (approximate) match. 

**=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])**     

| Argument                                | Description                                                  |
| --------------------------------------- | ------------------------------------------------------------ |
| **lookup_value**             Required*  | The value to search for  			*If omitted, XLOOKUP returns blank cells it finds in **lookup_array**. |
| **lookup_array**             Required   | The array or range to search                                 |
| **return_array**             Required   | The array or range to return                                 |
| **[if_not_found]**             Optional | Where a valid match is not found, return the [if_not_found] text you supply.If a valid match is not found, and [if_not_found] is missing, **#N/A** is returned. |
| **[match_mode]**             Optional   | Specify the match type:0 - Exact match. If none found, return #N/A. This is the default.-1 - Exact match. If none found, return the next smaller item.1 - Exact match. If none found, return the next larger item.2 - A wildcard match where *, ?, and ~ have [special meaning](https://support.office.com/en-us/f1/topic/using-wildcard-characters-in-searches-ef94362e-9999-4350-ad74-4d2371110adb?NS=EXCEL&Version=90). |
| **[search_mode]**             Optional  | Specify the search mode to use:1 - Perform a search starting at the first item. This is the default.-1 - Perform a reverse search starting at the last item.2 - Perform a binary search that relies on lookup_array being sorted in *ascending* order. If not sorted, invalid results will be returned.-2 - Perform a binary search that relies on lookup_array being sorted in *descending* order. If not sorted, invalid results will be returned. |