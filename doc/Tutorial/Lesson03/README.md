[:arrow_backward:](../Lesson02/README.md)


#                                         **Lesson 3 - VBA functions**

<!--------------------------------------------------------------------------------------------------------------------->


## Goal of the lesson

The goal of the lesson is to present basic functionalities of `qXL` VBA functions

- opening a connection
- running a query and printing to the Excel range

You can download the Excel [Workbook](../Lesson03/Lesson03.xlsm) with the content of the lesson.  


<!--------------------------------------------------------------------------------------------------------------------->
## Prerequisites
Although this whole tutorial is built based on the assumption that `Exxeleron` 
[system](https://github.com/exxeleron/enterprise-components) is installed locally, it is possible to use with other 
systems. All parts of code which need amending will be explicitly mentioned.


<!--------------------------------------------------------------------------------------------------------------------->
## Global variables

As all VBA functions are properties of qXL add-in object it is useful to declare it as a global variable. Similarly with the connection alias(es). In the example Workbook these two variables are initialized in the initialization sub:

```VBA

Public Sub initialization()
    Set qXL = Application.COMAddIns("qXL").Object
    Set wsConnection = ThisWorkbook.Worksheets("kdb+ connection")
    gsRdbConn = wsConnection.Range("nrRdbConnection").Value
    ...
```

This allows to use for all `q` calls general form `qXL.exampleFunction(conn, ...)` without separate initialization in
each sub. 


<!--------------------------------------------------------------------------------------------------------------------->
## Opening connection

Please note in the VBA code above that we used connection alias taken from Worksheet, assuming that connection was opened using Excel formulas, as in the [Lesson1](../Lesson01/README.md). An alternative approach would be to use VBA formula as below. 

```VBA

msg = qXL.qOpen(connAlias, hostName, portNumber, userName, password)

```

But we decided to use Excel connection because:

- it allows to use the same connection for VBA and Excel functions
- the modification is more user friendly

<!--------------------------------------------------------------------------------------------------------------------->

##Getting data from `q` to Excel with VBA

In this lesson we will cover the same examples as in the [Lesson1](../Lesson01/README.md) getting table as a result of 
custom q-query or function result. The example code to retrieve data from q is presented below: 

```VBA
dim vKdbQuery as Variant
vKdbQuery = qXL.qQuery(conn, sQuery)
```

Naturally, `qXL` object needs to be initialized first. The `sQuery` can be any `q` statement which returns a `q` object. The
result will be casted to type of `vKdbQuery` variable. `sQuery` can also be a function. In such a case all parameters need to listed inside of `qXL.qQuery`.

```VBA
Dim vKdbQuery As Double
vKdbQuery = qXL.qQuery(gsRdbConn, "+", qXL.qAtom("2", "f"), qXL.qAtom("2", "f"))
```

The most common case is getting a `q` table and printing in the Excel range. The easiest way to do this is to declare `vKdbQuery` as a Variant and then writing to properly sized range as can be seen below:

```VBA
Dim vKdbQuery As Variant
Dim rngResult As Range
Dim rngFirstCell as Range
Set rngFirstCell = Range("A1")
vKdbQuery = qXL.qQuery(gsRdbConn, sQuery)
Set rngResult = firstCell.Resize(UBound(vKdbQuery, 1), UBound(vKdbQuery, 2) + 1)
```   
where `rngFirstCell` is going to be the top left cell of the range. 

This conecpt is used in the `rngKdbQueryToRange` function which writes the result to range and it returns the range as a result. 

<!--------------------------------------------------------------------------------------------------------------------->

##Working with q tables

When repeatedly querying data to the same Worksheet it is very convenient to use Excel Tables. It enables to manage the result as one object with known size. Thus, we don't need to worry about overlapping results. 

Subroutine `kdbQueryToTable` creates Excel Table with given name using the range returned by `rngKdbQueryToRange`. If a table with the same name already exists, it is deleted.  

Applying the `kdbQueryToTable` function we are able to create customize queries in one line and display results in specified Excel range. In the Workbook [Lesson3](../Lesson03/Lesson03.xlsm) we created two example subroutines, which cover the same functionalities as in [Lesson1](../Lesson01/Lesson01.xlsx). 

- first which gets the result of q-sql query
- second which gets OHLC function result with input paramateres from Excel

The extra functionalities we have by using VBA are:

- displaying results as Excel tables
- error handling
- log messages
- button click actions


