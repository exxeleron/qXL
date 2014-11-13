[:arrow_backward:](Worksheet-Examples.md) | [:arrow_forward:](RTD-API-Functions.md)

# VBA examples

- [Preliminaries](VBA-Examples.md#preliminaries)
- [Opening and closing connection](VBA-Examples.md#opening-and-closing-connection)
- [Querying data](VBA-Examples.md#querying-data)
- [Converting data types](VBA-Examples.md#converting-data-types)
- [Working with lists](VBA-Examples.md#working-with-lists)
- [Working with dictionaries and tables](VBA-Examples.md#working-with-dictionaries-and-tables)

<!--------------------------------------------------------------------------------------------------------------------->
## Preliminaries

> :white_check_mark: Hint:

> All examples are available in Excel format from [here](examples/).

#### Notes

1. The VBA examples shown here produce exactly the same results as in the [worksheet examples](Worksheet-Examples.md), 
thus the screenshots have been omitted as more emphasis is given to the VBA code itself
1. Please note that when using VBA code, the VBA function name needs to be proceeded with the object itself, in all
examples which follow `qXL` is used:
  ```VBA
  msg = qXL.qOpen(connAlias, hostName, portNumber, userName, password)
  ```
  
1. In case of changes to the worksheet names and / or ranges, these need to be updated accordingly to ensure correct 
functioning of the given examples
1. It is worth looking at the [worksheet examples](Worksheet-Examples.md) first as they provide more details and 
relevant screenshots to illustrate particular functionality
1. For clarity VBA code snippets used below have been aligned to make those easier to read.

#### Q process

Before running examples it is assumed that q process is running on `localhost`, port `5001`.

#### Q test dataset

Sample functions and datasets for VBA examples are exactly the same as in 
[worksheet examples](Worksheet-Examples.md#q-test-dataset). These can be read from the worksheet:

```VBA
INfunc      = Replace(Sheets("Definitions").Range("A3"), "=", "")
INret       = Replace(Sheets("Definitions").Range("A4"), "=", "")
INdtTest    = Replace(Sheets("Definitions").Range("A5"), "=", "", 1, 1)
INmixedList = Replace(Sheets("Definitions").Range("A6"), "=", "")
INt1        = Replace(Sheets("Definitions").Range("A7"), "=", "")
INt2        = Replace(Sheets("Definitions").Range("A8"), "=", "")
INgetMax    = Replace(Sheets("Definitions").Range("A9"), "=", "")
```

and executed:

```VBA
func      = qXL.qQuery(connAlias, INfunc)
ret       = qXL.qQuery(connAlias, INret)
dtTest    = qXL.qQuery(connAlias, INdtTest)
mixedList = qXL.qQuery(connAlias, INmixedList)
t1        = qXL.qQuery(connAlias, INt1)
t2        = qXL.qQuery(connAlias, INt2)
getMax    = qXL.qQuery(connAlias, INgetMax)
```

#### Excel list separators

For functions in VBA only comma `,` has to be used :

```VBA
= qXL.qOpen(connAlias,hostName,portNumber,userName,password)
```

Other separators, for example semicolon `(;)`, will return a syntax error.

<!--------------------------------------------------------------------------------------------------------------------->
## Opening and closing connection

Following snippets show how to connect to a q process from VBA code.

At first, `connAlias` and `qXL` are defined at global level as these can be then used across different VBA modules:

```VBA
Public connAlias As String
Public qXL       As Object
```

Within `Connect()` sub-procedure, some local variables are defined and initialized using ranges from the worksheet:

```VBA
Dim hostName   As String
Dim portNumber As Integer
Dim userName   As String
Dim password   As String
Dim msg        As String
Dim addIn      As COMAddIn

connAlias  = Sheets("qOpen qClose").Range("B1")
hostName   = Sheets("qOpen qClose").Range("B2")
portNumber = Sheets("qOpen qClose").Range("B3")
userName   = Sheets("qOpen qClose").Range("B4")
password   = Sheets("qOpen qClose").Range("B5")
```

Initial connection to a q process is done in three steps:

1. Excel's `COMAddIn` object retrieves `qXL` add-in
1. `COMAddIn` is used to create `qXL` object
1. Connection to q is opened using `qOpen` function

```VBA
Set addIn = Application.COMAddIns("qXL")
Set qXL   = addIn.Object

msg = qXL.qOpen(connAlias, hostName, portNumber, userName, password)
```

If connection was successful, string `Connected` will be shown in cell `G2`, otherwise an error message will appear:

```VBA
If msg = connAlias Then
    Sheets("qOpen qClose").Range("G2") = "Connected"
Else
    Sheets("qOpen qClose").Range("G2") = msg
End If
End Sub
```

To close a connection to a q process, use the alias in the `qClose` function:

```VBA
Sub Disconnect()

Dim msg As String

If connAlias = "" Then
    msg = "Connection alias does not exist"
Else
    msg = qXL.qClose(connAlias)
End If
Sheets("qOpen qClose").Range("G2") = msg
End Sub
```

<!--------------------------------------------------------------------------------------------------------------------->
## Querying data

### Simple query

```VBA
Sheets("qQuery").Range("C3")    = qXL.qQuery(connAlias, "1b")
Sheets("qQuery").Range("C5")    = qXL.qQuery(connAlias, "2 + 2")
Sheets("qQuery").Range("C4")    = qXL.qQuery(connAlias, ".z.D")
Sheets("qQuery").Range("C6:E6") = qXL.qQuery(connAlias, "10* 1 2 3")
```

While scalar values only need to be mapped to one cell, arrays need to be mapped to the number of the array entries.
In the last example above, three cells will be used which will result in overwriting any previous content.

> :heavy_exclamation_mark: Note:

> Array elements (as well as content of dictionaries and tables) cannot be changed once the values are displayed in 
> Excel. Please read 
> [this section](http://office.microsoft.com/en-001/excel-help/guidelines-and-examples-of-array-formulas-HA010228458.aspx#BM2) 
> from Microsoft for more details about array constants.

### Function call

To call a function already defined in a q process:

```VBA
Sheets("qQuery").Range("C8") = qXL.qQuery(connAlias, "func", 3, 10)
``` 

Custom functions can also be dynamically defined on a q process:

```VBA
triple = qXL.qQuery(connAlias, "triple:{[p] :3*p}")
Sheets("qQuery").Range("C10") = qXL.qQuery(connAlias, "triple", 40)
```

This function can then be later used in the worksheet:

```
| =qQuery(C1,"triple",40) | 120 |
```

 or in q environment itself:

```q
q)triple[40] 
120
```

> :heavy_exclamation_mark: Note:

> Only atoms (single values) can be used as parameters in function calls. In cases where more complex structures are 
> needed, use Excel ranges instead. For example:
>
> ```
> Dim arrayMax As Variant
> arrayMax = Sheets("qQuery").Range("D11:F11")
>
> ' Print result
> Sheets("qQuery").Range("C11") = qXL.qQuery(connAlias, "getMax", qXL.qList(arrayMax, "i"))
> ```
>
> instead of:
>
> ```
> Sheets("qQuery").Range("C11") = qXL.qQuery(connAlias, "getMax", qXL.qList({11,55,10},"i"))
> ```
>
> This is especially important for handling lists, dictionaries and tables, which will be covered below in more detail.

### q-sql

`qQuery` can also be used to display content of the table. Simple table can be shown in the following way:

```VBA
Dim resultQueryTable1 As Range
Set resultQueryTable1 = Sheets("qQuery").Range("C13:D16")

resultQueryTable1 = qXL.qQuery(connAlias, "t1")
```

> :warning: Note:

> Quotes around table name (`"t1"`) indicate that this string will be treated as q statement. Omitting those will cause
> Excel to use `t1` as cell address and produce an error, for example here:
>
> ```VBA
> Sheets("qQuery").Range("C18") = qXL.qQuery(connAlias, t1)
> ```

Subset of the table can also be selected from the table, for example:

```VBA
Dim resultQueryTable2Select1 As Range
Set resultQueryTable2Select1 = Sheets("qQuery").Range("C20:D21")

resultQueryTable2Select1 = qXL.qQuery(connAlias, "select from t2 where colA=`d")
```

However, there is one point to remember - any complex content (list, dictionary, another table within the result, etc.) 
will not be displayed, instead Excel will show `#Value!` message. In case of `t2`:

```q
q)select from t2 where colA=`c
colA colB
---------------
c    1
c    `sym1`sym2
```

In the above snippet `colB` contains a list (`` `sym`sym2 ``) which will not be shown in the worksheet:

```VBA
Dim resultQueryTable2Select2 As Range
Set resultQueryTable2Select2 = Sheets("qQuery").Range("C23:D25")

resultQueryTable2Select2 = qXL.qQuery(connAlias, "select from t2 where colA=`c")
```

If needed, content can always be inspected by 'tweaking' query, providing more details, for example:

```VBA
Dim resultQueryTable2Select3 As Range
Set resultQueryTable2Select3 = Sheets("qQuery").Range("C27:D29")

resultQueryTable2Select3 = qXL.qQuery(connAlias, "flip last select from t2 where colA=`c")
```

> :white_check_mark: Hint:

> [`Ungroup`](http://code.kx.com/wiki/Reference/ungroup) function can also be used, for example:
> ```VBA
> Dim resultQueryTable2Select4 As Range
> Set resultQueryTable2Select4 = Sheets("qQuery").Range("C31:D45")
>
> resultQueryTable2Select4 = qXL.qQuery(connAlias, "ungroup 0!select raze raze enlist colB by colA from t2")
> ```

<!--------------------------------------------------------------------------------------------------------------------->
## Converting data types

`qAtom` is only used to convert data types between Excel and q. Please recall that function requires two parameters - 
value to be parsed and data type:

```
Object qAtom ( value, type )
```

In order to use this function in Excel, one will always need to combine it with some other function(s). For example:

```VBA
Sheets("qAtom").Range("C3") = qXL.qQuery(connAlias, "func", qXL.qAtom(10,    "i"), 20)
Sheets("qAtom").Range("C4") = qXL.qQuery(connAlias, "func", qXL.qAtom(10.99, "i"), 20)
```

where q is 'forced' to treat `10` as an `int` and use `func` to multiply both arguments. 
Similarly, enforcing simple data type checks can be done this way:

```
Sheets("qAtom").Range("C5") = qXL.qQuery(connAlias, "dtTest", qXL.qAtom(255, "i"), -6)
Sheets("qAtom").Range("C6") = qXL.qQuery(connAlias, "dtTest", qXL.qAtom(255, "i"), -8)
Sheets("qAtom").Range("C7") = qXL.qQuery(connAlias, "dtTest", qXL.qAtom(255, "e"), -8)
```

> :white_check_mark: Hint:

> Please visit [code.kx.com](http://code.kx.com/wiki/Reference/Datatypes) for details about available q data types. 
[Type mapping](Type-Mapping.md) lists strings used for conversions, also apply to `qList`, `qDict` and `qTable` 
functions described below.

<!--------------------------------------------------------------------------------------------------------------------->
## Working with lists

> :heavy_exclamation_mark: Note:

> `qList` can only operate using Excel ranges. Using values directly in this function can result in unexpected 
> behaviour!

`qList` performs conversion of Excel values to lists as understood by q. In general, if one-dimensional range is given,
`qList` will convert it to simple list regardless of the vertical / horizontal range. However, if two-dimensions are 
used, `qList` will convert those values to list of lists. For example, for the following input data:

```
| row / col |  K |  L |  M |
|-----------|----|----|----|
| 1         |    |  1 |    |    ( L0 )
| 2         |    |  2 |    |    ( L0 )
| 3         |    |  3 |    |    ( L0 )
| 4         |  4 |  5 |  6 |    ( L1 )
| 5         | 10 | 11 | 12 |    ( L2 )
| 6         | 13 | 14 | 15 |    ( L2 )
| 7         | 16 | 17 | 18 |    ( L2 )
| 8         |  7 |  8 |  9 |    ( L3 )
| 9         |  a |  b |  c |    ( L4 )
```

the range that `qList` can receive might be either:

1. Vertical - where input is processed as simple q list, for example from the following query:

  ```VBA
  ' Print result
  Dim resultList1 As Range
  Set resultList1 = Sheets("qList").Range("C3")

  ' Get input values
  Dim arrayList1 As Variant
  arrayList1 = Sheets("qList").Range("L1:L3")

  ' Execute function 
  resultList1 = qXL.qQuery(connAlias, "{.tst.vL:x}", qXL.qList(arrayList1, "i"))
  resultList1.Value = "{.tst.vL:x}"

  ' -> Additionally, execute '.tst.vL' in q environment
  ```

1. Horizontal - where input is also processed as simple lists:

  ```
  ' Print result
  Dim resultList2 As Range
  Set resultList2 = Sheets("qList").Range("C4")

  ' Get input values
  Dim arrayList2 As Variant
  arrayList2 = Sheets("qList").Range("K4:M4")

  ' Execute function 
  resultList2 = qXL.qQuery(connAlias, "{.tst.hL:x}", qXL.qList(arrayList2, "i"))
  resultList2.Value = "{.tst.hL:x}"

  ' -> Additionally, execute '.tst.hL' in q environment
  ```

1. Two-dimensional - range is treated as list of lists:

  ```
  ' Print result
  Dim resultList3 As Range
  Set resultList3 = Sheets("qList").Range("C5")

  ' Get input values
  Dim arrayList3 As Variant
  arrayList3 = Sheets("qList").Range("K5:M7")

  ' Execute function 
  resultList3 = qXL.qQuery(connAlias, "{.tst.2D:x}", qXL.qList(arrayList3, "i"))
  resultList3.Value = "{.tst.2D:x}"

  ' -> Additionally, execute '.tst.2D' in q environment
  ```

### Multiplication of lists

 ```VBA
 Dim resultList4 As Range
 Set resultList4 = Sheets("qList").Range("C6:E6")

 Dim arrayList4 As Variant
 arrayList4 = Sheets("qList").Range("K4:M4")

 Dim arrayList5 As Variant
 arrayList5 = Sheets("qList").Range("K8:M8")

 resultList4 = qXL.qQuery(connAlias, "func", qXL.qList(arrayList4, "i"), qXL.qList(arrayList5, "i"))
 ```

### Concatenation of mixed lists 

 ```VBA 
 Dim resultList5 As Range
 Set resultList5 = Sheets("qList").Range("C7:H7")

 Dim arrayList6 As Variant
 arrayList6 = Sheets("qList").Range("K8:M8")

 Dim arrayList7 As Variant
 arrayList7 = Sheets("qList").Range("K9:M9")

 resultList5 = qXL.qQuery(connAlias, "mixedList", qXL.qList(arrayList6, "i"), qXL.qList(arrayList7, "s"))
 ```

<!--------------------------------------------------------------------------------------------------------------------->
## Working with dictionaries and tables

> :heavy_exclamation_mark: Note:

> `qDict` and `qTable` can only operate using Excel ranges. Using values directly in bodies of these functions 
> can result in unexpected behaviour!

### `qDict`

`qDict` is used to build a dictionary. For example:

```
Dim arrayDictColumnNames As Variant
arrayDictColumnNames = Sheets("data").Range("A1:E1")

Dim arrayDictDataMatrix As Variant
arrayDictDataMatrix = Sheets("data").Range("A2:E5")

Dim resultDict1 As Range
Set resultDict1 = Sheets("qDict").Range("C3:G7")

resultDict1 = qXL.qQuery(connAlias, "{[x]:x}", qXL.qDict(arrayDictColumnNames, arrayDictDataMatrix, "ssfff"))
```

where:
- `arrayDictColumnNames ` - defines keys for dictionary
- `arrayDictDataMatrix` - defines data range
- `ssfff` - specifies data types for values for each of the keys

> :white_check_mark: Hint:

> For more details on dictionaries please visit [code.kx.com](http://code.kx.com/wiki/JB:QforMortals2/dictionaries) page

<!--------------------------------------------------------------------------------------------------------------------->
### `qTable`

Table can be created in similar fashion as in case of `qDict`. Using the same dataset as before following statements can
be used:

```
Dim arrayTableColumnNames As Variant
arrayTableColumnNames = Sheets("data").Range("A1:E1")

Dim arrayTableDataMatrix As Variant
arrayTableDataMatrix = Sheets("data").Range("A2:E5")

Dim arrayTableKey As Variant
arrayTableKey = Sheets("data").Range("A1")

Dim resultTable1 As Range
Set resultTable1 = Sheets("qTable").Range("C3:G7")

resultTable1 = qXL.qQuery(connAlias, "{[x]:x}", qXL.qTable(arrayTableColumnNames, arrayTableDataMatrix, "ssfff", arrayTableKey))
```

There are two small differences when comparing `qTable` to `qDict`:

1. `arrayTableColumnNames` - first parameter describes column names (vs. keys in `qDict`)
2. `arrayTableKey` - last parameter specifies key(s) for the table, in the above case column `Symbol` was used as a key

This difference can be seen clearly when comparing function definition for `qDict` and `qTable`:

```
Object qXL.qDict  ( keys,        values, types )
Object qXL.qTable ( columnNames, values, types, keys )
```

> :white_check_mark: Hint:

> For more details on keyed tables please visit 
[code.kx.com](http://code.kx.com/wiki/JB:QforMortals2/tables#Primary_Keys_and_Keyed_Tables) page.
