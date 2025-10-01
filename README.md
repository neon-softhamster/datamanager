This small library was created spontaneously for highly specialized table manipulation tasks. Its main purpose is to act as a wrapper for pandas, eliminating the need to think (or re-learn) how to work with it.

The library can read .ods tables, which are not particularly large, like the ones I often encounter when processing data from physics experiments. It can also write the calculated data to .ods files.

Library also has other useful functions for working with data. For example, you can take derivatives, find their extremes, and smooth the data.

In the future, I hope to add a convenient set of functions to reduce the amount of code that accumulates when preparing scientific illustrations in matplotlib.

### Installation

Remember to enable right pip venv before installation.

```
git clone https://github.com/neon-softhamster/datamanager.git
cd datamanager
pip install .
```

Add datamanager to your script `import datamanager as dm `

### Features

In order to load spreadsheet.ods do `table = dm.Reader('spreadsheet.ods')`. Now you can get columns from table object `col = dm.Reader(sheetName, colName)`. Pair sheetName, colName could be indexes like 0, 1 (which means 0-sheet, 1-column) or real names used in table such as "Sheet1", "DataCol1". Do not mix these types of arguments.

Sometimes it's useful to cycle through all columns on all sheets. You can use special Reader methods for this.

```
for sheet in table.sheetNames():
    for column in table.colNames(sheet):
        print(table.column(sheet, column)
```




