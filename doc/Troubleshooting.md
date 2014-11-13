###                                           **Troubleshooting**

<!--------------------------------------------------------------------------------------------------------------------->
`Troubleshooting` document describes known problems and solutions encountered during usage of `qXL` COM add-in. 

> Note:
  
> Please use our [google group](https://groups.google.com/d/forum/exxeleron) 
or open a [ticket](https://github.com/exxeleron/enterprise-components/issues) 
in case you enocounter any installation/startup problem which is not covered in this document.


#### Issue 1 - memory leak while using charts based on RTD data

##### Problem
It was observed that memory usage is constanly growing for Excel Workbooks which contain charts built dynamically based on the data from `RTD` formula. The increase of memory usage depends on the charts types and amount of data require. 
The issue is caused by the memory leak for Excel charts with external data source. 

##### Solution
The only possible solution is to close and open again the Workbook to free the memory. 

### Issue 2 - no data updates after calling `qRTDClose` and `qRTDOpen`. 

##### Problem
In case of calling `qRTDClose` function the subscription is properly closed and the `RTD` formuls is not updating. However, calling `qRTDOpen` with the same alias again does not result in subsribing for new values. 

##### Solution
Calling `qRTDOpen` with different alias refreshes the connection and results in subscribing for new values. 

### Issue 3 - different display of `qQueryRange` and `qQuery` nested results. 

#####Problem
Nested results (with more than two dimensions) are displayed differently in Excel in `qQuery` and `qQueryResult` formulas.
Whenever we try to display something different than `q` atom in single cell `qQuery` prints `#VALUE` error and `qQueryRange`
leaves the cell blank. 

Please see the example below: 

We query nested list which have atoms on all position except for the first element of the second row where there is another list. 

`qQuery` displays nested result as `#VALUE` error in Excel as can be seen below:

![qQueryNested](../doc/img/qQueryNested.png)


`qQueryRange` displays nested result as empty cell as can can seen below:


![qQueryRangeNested](../doc/img/qQueryRangeNested.png)


##### Solution

To have the same display `qQuery` can be wrappred with `IFERROR` formula as can be seen below: 

![qQueryNestedAligned](../doc/img/qQueryNestedAligned.png)

