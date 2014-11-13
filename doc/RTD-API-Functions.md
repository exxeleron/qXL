[:arrow_backward:](VBA-Examples.md) | [:arrow_forward:](RTD-Examples.md)

# RTD API Functions

- [qRtdOpen](RTD-API-Functions.md#qrtdopen)
- [qRtdClose](RTD-API-Functions.md#qrtdclose)
- [qRtdConfigure](RTD-API-Functions.md#qrtdconfigure)
- [RTD](RTD-API-Functions.md#rtd)

> :warning: Note:

> RTD functionality can only be used from the Excel's Worksheet (i.e. using these calls in VBA code will not work).

<!--------------------------------------------------------------------------------------------------------------------->
### qRtdOpen

Function used to open a RTD connection to a kdb+ process:

```
String qRtdOpen ( alias, hostname, port, username, password )
```

where:
- `alias` [`String`] - alias name that should be assigned to the connection currently being opened; this alias can 
later be used by other functions to reference specific connections (as multiple connections can be opened from the 
same Workbook object)
- `hostname` [`String`] - name or IP address of the host to which connection should be opened
- `port` [`Int`] - port number of q process to connect to
- `username` [`String`] - username used to connect to q process
- `password` [`String`] - password used to connect to q process

Returns either:
- alias associated with currently opened connection - in case opened successfully
- description of error - in case connection could not be established

The connection is opened only once. When using the same alias for all subsequent function calls, the same existing 
connection will be used.

> Note: 

>To force re-opening of the connection use different alias.

[**Examples**](RTD-Examples.md#opening-and-closing-connection)

<!--------------------------------------------------------------------------------------------------------------------->
### qRtdClose

Function used to correctly close the connection associated with given alias:

```
String qRtdClose ( alias )
```

where:
- `alias` [`String`] - closes the connection associated with given alias

Returns string `Closed` if connection closed successfully or error description in case of failure.

[**Examples**](RTD-Examples.md#opening-and-closing-connection)

<!--------------------------------------------------------------------------------------------------------------------->
### qRtdConfigure

Function used to configure various aspects of RTD server behaviour:

```
Object qRtdConfigure ( paramName, paramValue )
```

where:
- `paramName` [`String`] - name of parameter to be configured; currently supported configurable parameters are:
```
| Parameter name      | Parameter type | Description                                   | Default value |
|---------------------|----------------|-----------------------------------------------|---------------|
| sym.column.name     | String         | Name of the column containing primary         | sym           |
|                     |                |    instrument identifier (so called symbol)   |               |
| function.add        | String         | Name of the function that should be present   | .u.add        |
|                     |                |    on q process allowing addition of          |               |
|                     |                |    instruments to current subscription list   |               |
| function.sub        | String         | Name of the function that should be present   | .u.sub        |
|                     |                |    on q process allowing subscription         |               |
| function.del        | String         | Name of the function that should be present   | .u.del        |
|                     |                |    on q process allowing removal of the       |               |
|                     |                |    instruments from current subscription list |               |
| data.history.length | Int            | Length of the history vector                  | 1             |
```
- `paramValue` [`Object`] - value of the parameter to be set

Returns value of the parameter in question. In case it has been correctly set, it will be equal to `ParamValue` 
otherwise will return `old` value of the parameter.

[**Examples**](RTD-Examples.md#configuration)

<!--------------------------------------------------------------------------------------------------------------------->

### RTD

Function used to subscribe to certain value from within a cell in workbook so that it automatically gets populated with 
new values as soon as they are available:

```
Object RTD ( progID, server, topic1, topic2, topic3, topic4, topic5, topic6)
```

where:
- `progID` [`String`] - ID of RTD server implementation; should be set to `qxlrtd`; required
- `server` [`String`] - not used; required
- `topic1` [`String`] - connection alias returned via `qSubscribe`; required
- `topic2` [`String`] - name of the table from which instruments should be subscribed; required
- `topic3` [`String`] - instrument name; required
- `topic4` [`String`] - column name within the table to which the subscription is being made; required
- `topic5` [`String`] - history index; optional - only needed when `Topic3` equals back tick symbol (\`); when filling 
`Topic6` we have to have `Topic5` (can be empty)
- `topic6` [`Number`] - Symbol Mapping; optional

> Note:

> Since `Topic6` can be used only for back tick subscription, system cannot determine where to put data for 
particular symbols. That is why `Topic6` has been introduced – it is used to map incoming symbols to numbers. 
First symbol which is delivered from database is mapped with first empty number of `Topic6` and it stays mapped 
until all `RTD` formulas with given value of `Topic6` are present in the sheet.

Updates the content of cell to which this function call is linked with the value matching subscription criteria.


[**Examples**](RTD-Examples.md#subscribe-single-symbol)
