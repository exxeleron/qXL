[:arrow_backward:](Installation.md) | [:arrow_forward:](Worksheet-Examples.md)

# Worksheet and VBA API Functions

- [qOpen](Worksheet-VBA-API-Functions.md#qopen)
- [qClose](Worksheet-VBA-API-Functions.md#qclose)
- [qQuery](Worksheet-VBA-API-Functions.md#qquery)
- [qAtom](Worksheet-VBA-API-Functions.md#qatom)
- [qList](Worksheet-VBA-API-Functions.md#qlist)
- [qDict](Worksheet-VBA-API-Functions.md#qdict)
- [qTable](Worksheet-VBA-API-Functions.md#qtable)


> :white_check_mark: Note:

> Given functions can be applied directly from Excel's Worksheet or through VBA code.

<!--------------------------------------------------------------------------------------------------------------------->
### qOpen

Function used to open a connection to a kdb+ process:

```
String qOpen ( alias, hostname, port, username, password, reEval )
```

where:
- `alias` [`String`] - alias name that should be assigned to the connection currently opened; this alias can later 
be used by other functions to reference specific connections (as multiple connections can be opened from the same 
`Workbook` object)
- `hostname` [`String`] - name or IP address of the host to which connection should be opened
- `port` [`Int`] - port number of q process to connect to
- `username` [`String`] - username used to connect to q process (optional)
- `password` [`String`] - password used to connect to q process (optional)
- `reEval` [`Cell`] - cell address used for re-evaluating connection to kdb+ server, any data type inside the cell 
can be used (optional)

Returns:
- alias bound with currently opened connection (when opened successfully)
- description of error (in case connection could not be established)

> :heavy_exclamation_mark: Note:

> Internally, `qXL` will store connection details as per given alias only. Therefore, one needs to be careful not to 
> overwrite connection alias with different host and/or port number. For example if at first this call is executed:
> 
> ```
> =qOpen("testConnection","localhost",5001)
> ```
>
> followed by:
>
> ```
> =qOpen("testConnection",172.16.254.1,17000)
> ```
> 
> then all functions using `testConnection` alias will refer to host `172.16.254.1` on port `17000`.

**Examples:** [**worksheet**](Worksheet-Examples.md#opening-and-closing-connection),
[**VBA**](VBA-Examples.md#opening-and-closing-connection)

<!--------------------------------------------------------------------------------------------------------------------->
### qClose

Function used to correctly close the connection with given alias:

```
String qClose ( alias )
```

where:
- `alias` [`String`] - closes the connection with given alias

Returns:
- string `Closed` if connection is closed successfully 
- error description in case of failure

**Examples:** [**worksheet**](Worksheet-Examples.md#opening-and-closing-connection),
[**VBA**](VBA-Examples.md#opening-and-closing-connection)

<!--------------------------------------------------------------------------------------------------------------------->
### qQuery

Function used to query data or call a kdb+ function using the connection with given alias:

```
Object qQuery ( alias, query, p1, p2, p3, p4, p5, p6, p7, p8 )
```

where:
- `alias` [`String`] - connection alias to be used
- `query` [`String`] - query / function to be called
- `p1` to `p8` [`String`] - optional parameters

Returns the query or function call result.

**Examples:** [**worksheet**](Worksheet-Examples.md#querying-data),
[**VBA**](VBA-Examples.md#querying-data)

<!--------------------------------------------------------------------------------------------------------------------->
### qAtom

Function performs conversion of incoming value to specified q type and stores it in global container returning unique 
identifier for the data:

```
Object qAtom ( value, type )
```

where:
- `value` [`Object`] - value to be converted
- `type` [`String`] - conversion string; see [type mapping](../doc/Type-Mapping.md) section for more details

Function returns the following array: `(conversionKey,marker,error)`.

**Examples:** [**worksheet**](Worksheet-Examples.md#converting-data-types),
[**VBA**](VBA-Examples.md#converting-data-types)

<!--------------------------------------------------------------------------------------------------------------------->
### qList

Function performs conversion of incoming value (range) to specified q list and stores it in global container returning unique 
identifier for the data:

```
Object qList ( value, type )
```

where:
- `value` [`Object`] - value to be converted
- `type` [`String`] - conversion string; see [type mapping](../doc/Type-Mapping.md) section for more details

Function returns the following array: `(conversionKey,marker,error)`.

**Examples:** [**worksheet**](Worksheet-Examples.md#working-with-lists),
[**VBA**](VBA-Examples.md#working-with-lists)

<!--------------------------------------------------------------------------------------------------------------------->
### qDict

Function creates q dictionary from provided data:

```
Object qDict ( keys, values, types )
```

where:
- `keys` [`Object`] - dictionary keys
- `values` [`Object`] - values to be converted
- `types` [`String`] - conversion strings; see [type mapping](../doc/Type-Mapping.md) section for more details

Function returns the following array: `(conversionKey,marker,error)`.

**Examples:** [**worksheet**](Worksheet-Examples.md#working-with-dictionaries-and-tables),
[**VBA**](VBA-Examples.md#working-with-dictionaries-and-tables)

<!--------------------------------------------------------------------------------------------------------------------->
### qTable

Function creates q table from provided data:

```
Object qTable ( columnNames, values, types, keys )
```

where:
- `columnNames` [`Object`] - name of the columns
- `values` [`Object`] - data for the table
- `types` [`String`] - type specification (conversion strings) of the columns; see [type mapping](../doc/Type-Mapping.md) section for more details 
- `keys` [`Object`] - sublist of columns that should be considered as keys

**Examples:** [**worksheet**](Worksheet-Examples.md#working-with-dictionaries-and-tables),
[**VBA**](VBA-Examples.md#working-with-dictionaries-and-tables)
