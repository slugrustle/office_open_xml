## BasicWorkbook

BasicWorkbook permits the creation of multi-sheet Workbooks containing cells of only three types:

1. Cell containing a numeric value
2. Cell with a formula
3. Cell with a string or text value

Different horizontal and vertical alignments of the value within a cell are supported, as are merged cells, bold text, wrapped text, various number formats, custom row heights, and custom column widths.

The feature set is deliberately kept minimal to avoid the trap of reimplementing the entire office open xml specification in C++.

The file test1.xlsx was produced by the code in BasicWorkbookDemo.cpp.
