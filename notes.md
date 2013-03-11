Notes on XML format:

- Cell ss:Index and Row ss:Index is used if the Cell column number or Row number
  are not sequential from 1 to X eg: if the first Cell of a row is in the 8th
  Column, the Cell ss:Index is 8.

- ExpandedColumnCount and ExpandedRowCount needs to be increased whenever a
  Row is added or a Column is added.
