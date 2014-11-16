// 
// Copyright (c) 2011-2014 Exxeleron GmbH
// 
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// 
//   http://www.apache.org/licenses/LICENSE-2.0
// 
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
// 

//
// Based on Excel-DNA by Govert van Drimmelen
// https://exceldna.codeplex.com/
//

#region

using System;
using System.IO;
using System.Runtime.Caching;
using System.Windows.Forms;
using ExcelDna.ComInterop;
using ExcelDna.Integration;

#endregion

namespace qXL
{
    // ReSharper disable UnusedMember.Global
    // ReSharper disable InconsistentNaming
    public class qXLAddIn : IExcelAddIn
    // ReSharper restore InconsistentNaming
    // ReSharper restore UnusedMember.Global
    {
        // ReSharper disable InconsistentNaming
        private static readonly qXLShared _qXL = new qXLShared();
        // ReSharper restore InconsistentNaming

        private qXLComAddIn _comAddin;

        #region Excel-DNA

        public void AutoOpen()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex => "!!! EXCEPTION: " + ex.ToString());
            ComServer.DllRegisterServer();
            try
            {
                _comAddin = new qXLComAddIn();
                ExcelComAddInHelper.LoadComAddIn(_comAddin);
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"Error loading COM AddIn: " + ex);
            }
        }

        /// On sheet closing event realize all active handle to q
        public void AutoClose()
        {
            _qXL.qCloseAll();
            ComServer.DllUnregisterServer();
        }

        #endregion

        #region AddIn

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Opens connection to specified host and binds it with provided alias.
        /// </summary>
        /// <param name="alias">logical identifier for the connection</param>
        /// <param name="hostname">hostname,fqdn or ip of the machine running q process to which connection should be established</param>
        /// <param name="port">port number of the q process</param>
        /// <param name="username">username </param>
        /// <param name="password">password</param>
        /// <param name="reEval">reEval</param>
        /// <returns>requested logical identifier for the connection in case connection can be established</returns>
        [ExcelFunction(Description = "Opens a connection to specified host and binds it with provided alias.",
            Category = "qXL", Name = "qOpen")]
        // ReSharper disable UnusedMember.Global
        public static object Open([ExcelArgument("Logical identifier for the connection.")] string alias,
            // ReSharper restore UnusedMember.Global
            [ExcelArgument(
                "Hostname, fqdn or ip of the machine running q process to which connection should be established."
                )] string hostname,
            [ExcelArgument("Port number of the q process.")] object port,
            [ExcelArgument("Username (optional).")] string username = null,
            [ExcelArgument("Password (optional).")] string password = null,
            // ReSharper disable UnusedParameter.Global
            [ExcelArgument("reEval (optional).")] object reEval = null)
        // ReSharper restore UnusedParameter.Global
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            {
                return ExcelEmpty.Value;
            }

            try
            {
                return _qXL.qOpen(alias, hostname, port, username, password);
            }
            catch (Exception e)
            {
                return "ERR: " + e.Message;
            }
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Closes the connection associated with given alias.
        /// </summary>
        /// <param name="alias"></param>
        /// <returns>"Closed" or error message in case connection could not be closed.</returns>
        [ExcelFunction(Description = "Closes the connection associated with given alias.",
            Category = "qXL", Name = "qClose")]
        // ReSharper disable UnusedMember.Global
        public static object Close([ExcelArgument("Logical identifier for the connection.")] string alias)
        // ReSharper restore UnusedMember.Global
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            {
                return ExcelEmpty.Value;
            }

            try
            {
                return _qXL.qClose(alias);
            }
            catch (Exception e)
            {
                return "ERR: " + e.Message;
            }
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Function for sending queries to q process.
        /// </summary>
        /// <param name="alias"> alias of the connection, this value can be received by calling Open</param>
        /// <param name="query">name of the function to be called or string to be evaluated within q process </param>
        /// <param name="p1">(optional) first parameter to the function call</param>
        /// <param name="p2">(optional) second parameter to the function call</param>
        /// <param name="p3">(optional) third parameter to the function call</param>
        /// <param name="p4">(optional) fourth parameter to the function call</param>
        /// <param name="p5">(optional) fifth parameter to the function call</param>
        /// <param name="p6">(optional) sixth parameter to the function call</param>
        /// <param name="p7">(optional) seventh parameter to the function call</param>
        /// <param name="p8">(optional) eighth parameter to the function call</param>
        /// <returns></returns>
        [ExcelFunction(Description = "Function for sending queries to kdb+.", Category = "qXL",
            Name = "qQuery")]
        // ReSharper disable UnusedMember.Global
        public static object Query(
            // ReSharper restore UnusedMember.Global
            [ExcelArgument("Alias of the connection, this value can be received by calling Open.")] string alias,
            [ExcelArgument("Name of the function to be called or string to be evaluated within q process.")] object
                query,
            [ExcelArgument("First parameter to the function call (optional).")] object p1 = null,
            [ExcelArgument("Second parameter to the function call (optional).")] object p2 = null,
            [ExcelArgument("Third parameter to the function call (optional).")] object p3 = null,
            [ExcelArgument("Fourth parameter to the function call (optional).")] object p4 = null,
            [ExcelArgument("Fifth parameter to the function call (optional).")] object p5 = null,
            [ExcelArgument("Sixth parameter to the function call (optional).")] object p6 = null,
            [ExcelArgument("Seventh parameter to the function call (optional).")] object p7 = null,
            [ExcelArgument("Eighth parameter to the function call (optional).")] object p8 = null)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            {
                return ExcelEmpty.Value;
            }

            try
            {
                var key = alias + query;
                object[,] r;
                if (MemoryCache.Default.Contains(key))
                {
                    r = MemoryCache.Default[key] as object[,];
                }
                else
                {
                    var result = _qXL.qQuery(alias, query, p1, p2, p3, p4, p5, p6, p7, p8);

                    if (result == null) return query; //null gets returned only when function definition has been sent to q.
                    if (!(result is object[,])) return result;

                    r = result as object[,];
                    MemoryCache.Default.Add(key, r, DateTimeOffset.Now.AddSeconds(30));
                }
                return ArrayResizer.ResizeCached(r, key);
            }
            catch (IOException io)
            {
                //this normally means that the process has been terminated on the receiving site
                // so clear the connection alias.
                return "ERR: " + io.Message;
            }
            catch (Exception e)
            {
                return "ERR: " + e.Message;
            }
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Alternative function for sending queries to q process.
        /// </summary>
        /// <param name="alias"> alias of the connection, this value can be received by calling Open</param>
        /// <param name="query">name of the function to be called or string to be evaluated within q process </param>
        /// <param name="p1">(optional) first parameter to the function call</param>
        /// <param name="p2">(optional) second parameter to the function call</param>
        /// <param name="p3">(optional) third parameter to the function call</param>
        /// <param name="p4">(optional) fourth parameter to the function call</param>
        /// <param name="p5">(optional) fifth parameter to the function call</param>
        /// <param name="p6">(optional) sixth parameter to the function call</param>
        /// <param name="p7">(optional) seventh parameter to the function call</param>
        /// <param name="p8">(optional) eighth parameter to the function call</param>
        /// <returns></returns>
        [ExcelFunction(Description = "Function for sending queries to kdb+.", Category = "qXL",
            Name = "qQueryRange")]
        // ReSharper disable UnusedMember.Global
        public static object QueryRange(
            // ReSharper restore UnusedMember.Global
            [ExcelArgument("Alias of the connection, this value can be received by calling Open.")] string alias,
            [ExcelArgument("Name of the function to be called or string to be evaluated within q process.")] object
                query,
            [ExcelArgument("First parameter to the function call (optional).")] object p1 = null,
            [ExcelArgument("Second parameter to the function call (optional).")] object p2 = null,
            [ExcelArgument("Third parameter to the function call (optional).")] object p3 = null,
            [ExcelArgument("Fourth parameter to the function call (optional).")] object p4 = null,
            [ExcelArgument("Fifth parameter to the function call (optional).")] object p5 = null,
            [ExcelArgument("Sixth parameter to the function call (optional).")] object p6 = null,
            [ExcelArgument("Seventh parameter to the function call (optional).")] object p7 = null,
            [ExcelArgument("Eighth parameter to the function call (optional).")] object p8 = null)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            {
                return ExcelEmpty.Value;
            }

            try
            {
                var result = _qXL.qQuery(alias, query, p1, p2, p3, p4, p5, p6, p7, p8);
                if (result == null) return query; //null gets returned only when function definition has been sent to q.
                if (result is object[,])
                {
                    return ArrayResizer.Resize(result as object[,]);
                }
                return result;
            }
            catch (IOException io)
            {
                //this normally means that the process has been terminated on the receiving site
                // so clear the connection alias.
                return "ERR: " + io.Message;
            }
            catch (Exception e)
            {
                return "ERR: " + e.Message;
            }
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     This function performs conversion of incoming value to specified Q type
        ///     and stores it in global container returning unique identifier for the data.
        ///     The following describes conversion strings:
        ///     b - boolean
        ///     x - byte
        ///     h - short
        ///     i - int
        ///     j - long
        ///     e - real
        ///     f - float
        ///     c - char
        ///     s - symbol
        ///     m - month
        ///     d - date
        ///     z - datetime
        ///     u - minute
        ///     v - second
        ///     t - time
        ///     p - timestamp
        ///     n - timespan
        ///     Function returns the following array: (conversionKey,marker,error)
        ///     conversionKey - unique id of data that got converted (used to access it from other functions)
        ///     marker - special string fabricated to prevent random recogniton of converted data structures
        ///     error - in case conversion finished with errors error message otherwise string.
        /// </summary>
        /// <param name="value">value to be converted</param>
        /// <param name="type">converstion string</param>
        /// <returns>array (conversionKey,marker,error) </returns>
        [ExcelFunction(
            Description =
                "This function performs conversion of incoming value to specified Q type and stores it in global container returning unique identifier for the data."
            , Category = "qXL",
            Name = "qAtom")]
        // ReSharper disable UnusedMember.Global
        public static object[] QAtom(object value, object type)
        // ReSharper restore UnusedMember.Global
        {
            return ExcelDnaUtil.IsInFunctionWizard() ? new object[] { ExcelEmpty.Value } : _qXL.qAtom(value, type);
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     This function performs conversion of incoming value(range) to specified Q list
        ///     and stores it in global container returning unique identifier for the data.
        ///     The following describes conversion strings:
        ///     b - boolean
        ///     x - byte
        ///     h - short
        ///     i - int
        ///     j - long
        ///     e - real
        ///     f - float
        ///     c - char
        ///     s - symbol
        ///     m - month
        ///     d - date
        ///     z - datetime
        ///     u - minute
        ///     v - second
        ///     t - time
        ///     p - timestamp
        ///     n - timespan
        ///     Function returns the following array: (conversionKey,marker,error)
        ///     conversionKey - unique id of data that got converted (used to access it from other functions)
        ///     marker - special string fabricated to prevent random recogniton of converted data structures
        ///     error - in case conversion finished with errors error message otherwise string.
        /// </summary>
        /// <param name="value">value to be converted</param>
        /// <param name="types">type specification e.g.: "sifp" </param>
        /// <returns>array: (conversionKey,marker,error)</returns>
        [ExcelFunction(
            Description =
                "This function performs conversion of incoming value(range) to specified Q list and stores it in global container returning unique identifier for the data."
            , Category = "qXL", Name = "qList")]
        // ReSharper disable UnusedMember.Global
        public static object[] QList([ExcelArgument("value to be converted")] object value,
            // ReSharper restore UnusedMember.Global
            [ExcelArgument("type specification e.g.: \"sifp\"")] object types)
        {
            return ExcelDnaUtil.IsInFunctionWizard() ? new object[] { ExcelEmpty.Value } : _qXL.qList(value, types);
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     This function creates Q dictionary from provided data.
        /// </summary>
        /// <param name="keys">keys</param>
        /// <param name="values">values</param>
        /// <param name="types">type specification e.g.: "sift"</param>
        /// <returns>array: (conversionKey,marker,error)</returns>
        [ExcelFunction(Description = "This function creates Q dictionary from provided data.",
            Category = "qXL", Name = "qDict")]
        // ReSharper disable UnusedMember.Global
        public static object[] QDict([ExcelArgument("keys")] object keys,
            // ReSharper restore UnusedMember.Global
            [ExcelArgument("values")] object values,
            [ExcelArgument("type specification e.g.: \"sift\"")] object types)
        {
            return ExcelDnaUtil.IsInFunctionWizard()
                ? new object[] { ExcelEmpty.Value }
                : _qXL.qDict(keys, values, types);
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     This function creates QTable from the provided data.
        /// </summary>
        /// <param name="columnNames">name of the columns</param>
        /// <param name="values">data to be put into the table</param>
        /// <param name="types">type specification for columns</param>
        /// <param name="keys">sublist of columns that should be considered as keys</param>
        /// <returns>array: (conversionKey,marker,error)</returns>
        [ExcelFunction(Description = "This function creates QTable from the provided data.",
            Category = "qXL", Name = "qTable")]
        // ReSharper disable UnusedMember.Global
        public static object[] QTable([ExcelArgument("name of the columns")] object columnNames,
            // ReSharper restore UnusedMember.Global
            [ExcelArgument("data to be put into the table")] object values,
            [ExcelArgument("type specification for columns")] object types,
            [ExcelArgument("sublist of columns that should be considered as keys")] object keys = null)
        {
            return ExcelDnaUtil.IsInFunctionWizard()
                ? new object[] { ExcelEmpty.Value }
                : _qXL.qTable(columnNames, values, types, keys);
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Returns the qXL (assembly) name and version number.
        /// </summary>
        [ExcelFunction(Description = "Returns the qXL full name and version number.",
            Category = "qXL", Name = "qXLAbout")]
        // ReSharper disable UnusedMember.Global
        public static string About()
        {
            return _qXL.qXLAbout();
        }

        //------------------------------------------------------------------//
        /// <summary>
        ///     Verifies whether the param has been converted or not, in case it undergone
        ///     conversion extracts the data using conversion key.
        /// </summary>
        /// <param name="value"></param>
        /// <returns>converted parameter</returns>
        [ExcelFunction(
            Description =
                "Verifies whether the param has been converted or not, in case it undergone conversion extracts the data using conversion key"
            ,
            Category = "qXL", Name = "qConvert")]
        // ReSharper disable UnusedMember.Local
        private static object QConvert([ExcelArgument("value to be converted")] object value)
        // ReSharper restore UnusedMember.Local
        {
            return ExcelDnaUtil.IsInFunctionWizard() ? ExcelEmpty.Value : _qXL.qConvert(value);
        }

        #endregion

        #region ArrayResizer

// ReSharper disable ClassNeverInstantiated.Local
        private class ArrayResizer : XlCall
// ReSharper restore ClassNeverInstantiated.Local
        {

            // This function will run in the UDF context.
            // Needs extra protection to allow multithreaded use.
            public static object Resize(object[,] array)
            {
                var caller = Excel(xlfCaller) as ExcelReference;
                if (caller == null)
                    return array;

                var rows = array.GetLength(0);
                var columns = array.GetLength(1);

                if (rows == 0 || columns == 0)
                    return array;

                if ((caller.RowLast - caller.RowFirst + 1 == rows) &&
                    (caller.ColumnLast - caller.ColumnFirst + 1 == columns))
                {
                    // Size is already OK - just return result
                    return array;
                }

                var rowLast = caller.RowFirst + rows - 1;
                var columnLast = caller.ColumnFirst + columns - 1;

                // Check for the sheet limits
                if (rowLast > ExcelDnaUtil.ExcelLimits.MaxRows - 1 ||
                    columnLast > ExcelDnaUtil.ExcelLimits.MaxColumns - 1)
                {
                    // Can't resize - goes beyond the end of the sheet - just return #VALUE
                    // (Can't give message here, or change cells)
                    return ExcelError.ExcelErrorValue;
                }

                // TODO: Add some kind of guard for ever-changing result?
                if (columns > 1)
                {
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        var target = new ExcelReference(caller.RowFirst, caller.RowFirst, caller.ColumnFirst + 1, columnLast);
                        var firstRow = new object[columns - 1];
                        for (var i = 1; i < columns; i++)
                        {
                            firstRow[i - 1] = array[0, i];
                        }
                        target.SetValue(firstRow);
                    });
                }
                if (rows > 1)
                {
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        var target = new ExcelReference(caller.RowFirst + 1, rowLast, caller.ColumnFirst, columnLast);
                        var data = new object[rows - 1, columns];
                        for (var i = 1; i < rows; i++)
                        {
                            for (var j = 0; j < columns; j++)
                            {
                                data[i - 1, j] = array[i, j];
                            }
                        }
                        target.SetValue(data);
                    });
                }
                // Return what we have - to prevent flashing #N/A
                return array;
            }

            // This function will run in the UDF context.
            // Needs extra protection to allow multithreaded use.
            public static object ResizeCached(object[,] array, String key)
            {
                var caller = Excel(xlfCaller) as ExcelReference;
                if (caller == null)
                    return array;

                var rows = array.GetLength(0);
                var columns = array.GetLength(1);

                if (rows == 0 || columns == 0)
                    return array;

                if ((caller.RowLast - caller.RowFirst + 1 == rows) &&
                    (caller.ColumnLast - caller.ColumnFirst + 1 == columns))
                {
                    MemoryCache.Default.Remove(key);
                    // Size is already OK - just return result
                    return array;
                }

                var rowLast = caller.RowFirst + rows - 1;
                var columnLast = caller.ColumnFirst + columns - 1;

                // Check for the sheet limits
                if (rowLast > ExcelDnaUtil.ExcelLimits.MaxRows - 1 ||
                    columnLast > ExcelDnaUtil.ExcelLimits.MaxColumns - 1)
                {
                    // Can't resize - goes beyond the end of the sheet - just return #VALUE
                    // (Can't give message here, or change cells)
                    return ExcelError.ExcelErrorValue;
                }

                // TODO: Add some kind of guard for ever-changing result?
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    // Create a reference of the right size
                    var target = new ExcelReference(caller.RowFirst, rowLast, caller.ColumnFirst, columnLast,
                        caller.SheetId);
                    DoResize(target, key); // Will trigger a recalc by writing formula
                });
                // Return what we have - to prevent flashing #N/A
                return array;
            }

            private static void DoResize(ExcelReference target, String key)
            {
                // Get the current state for reset later
                using (new ExcelEchoOffHelper())
                using (new ExcelCalculationManualHelper())
                {
                    var firstCell = new ExcelReference(target.RowFirst, target.RowFirst, target.ColumnFirst,
                        target.ColumnFirst, target.SheetId);

                    // Get the formula in the first cell of the target
                    var formula = (string)Excel(xlfGetCell, 41, firstCell);
                    var isFormulaArray = (bool)Excel(xlfGetCell, 49, firstCell);
                    if (isFormulaArray)
                    {
                        // Select the sheet and firstCell - needed because we want to use SelectSpecial.
                        using (new ExcelSelectionHelper(firstCell))
                        {
                            // Extend the selection to the whole array and clear
                            Excel(xlcSelectSpecial, 6);
                            var oldArray = (ExcelReference)Excel(xlfSelection);

                            oldArray.SetValue(ExcelEmpty.Value);
                        }
                    }
                    // Get the formula and convert to R1C1 mode
                    var isR1C1Mode = (bool)Excel(xlfGetWorkspace, 4);
                    var formulaR1C1 = formula;
                    if (!isR1C1Mode)
                    {
                        object formulaR1C1Obj;
                        var formulaR1C1Return = TryExcel(xlfFormulaConvert, out formulaR1C1Obj, formula, true, false,
                            ExcelMissing.Value, firstCell);
                        if (formulaR1C1Return != XlReturn.XlReturnSuccess || formulaR1C1Obj is ExcelError)
                        {
                            MemoryCache.Default.Remove(key);
                            var firstCellAddress = (string)Excel(xlfReftext, firstCell, true);
                            Excel(xlcAlert,
                                "Cannot resize array formula at " + firstCellAddress +
                                " - formula might be too long when converted to R1C1 format.");
                            firstCell.SetValue("'" + formula);
                            return;
                        }
                        formulaR1C1 = (string)formulaR1C1Obj;
                    }
                    // Must be R1C1-style references
                    object ignoredResult;
                    var result = MemoryCache.Default[key] as object[,];
                    //Debug.Print("Resizing START: " + target.RowLast);
                    var formulaArrayReturn = TryExcel(xlcFormulaArray, out ignoredResult, formulaR1C1, target);
                    //Debug.Print("Resizing FINISH");

                    // TODO: Find some dummy macro to clear the undo stack

                    if (formulaArrayReturn == XlReturn.XlReturnSuccess)
                    {
                        if (result != null) MemoryCache.Default.Add(key, result, DateTimeOffset.Now.AddSeconds(30));
                        return;
                    }
                    MemoryCache.Default.Remove(key);
                    var firstCellAddress1 = (string)Excel(xlfReftext, firstCell, true);
                    Excel(xlcAlert,
                        "Cannot resize array formula at " + firstCellAddress1 +
                        " - result might overlap another array.");
                    // Might have failed due to array in the way.
                    firstCell.SetValue("'" + formula);
                }
            }
        }

        // RIIA-style helpers to deal with Excel selections    
        // Don't use if you agree with Eric Lippert here: http://stackoverflow.com/a/1757344/44264

        private class ExcelCalculationManualHelper : XlCall, IDisposable
        {
            private readonly object _oldCalculationMode;

            public ExcelCalculationManualHelper()
            {
                _oldCalculationMode = Excel(xlfGetDocument, 14);
                Excel(xlcOptionsCalculation, 3);
            }

            public void Dispose()
            {
                Excel(xlcOptionsCalculation, _oldCalculationMode);
            }
        }

        private class ExcelEchoOffHelper : XlCall, IDisposable
        {
            private readonly object _oldEcho;

            public ExcelEchoOffHelper()
            {
                _oldEcho = Excel(xlfGetWorkspace, 40);
                Excel(xlcEcho, false);
            }

            public void Dispose()
            {
                Excel(xlcEcho, _oldEcho);
            }
        }

        // Select an ExcelReference (perhaps on another sheet) allowing changes to be made there.
        // On clean-up, resets all the selections and the active sheet.
        // Should not be used if the work you are going to do will switch sheets, amke new sheets etc.
        private class ExcelSelectionHelper : XlCall, IDisposable
        {
            private readonly object _oldActiveCellOnActiveSheet;

            private readonly object _oldActiveCellOnRefSheet;
            private readonly object _oldSelectionOnActiveSheet;
            private readonly object _oldSelectionOnRefSheet;

            public ExcelSelectionHelper(ExcelReference refToSelect)
            {
                // Remember old selection state on the active sheet
                _oldSelectionOnActiveSheet = Excel(xlfSelection);
                _oldActiveCellOnActiveSheet = Excel(xlfActiveCell);

                // Switch to the sheet we want to select
                var refSheet = (string)Excel(xlSheetNm, refToSelect);
                Excel(xlcWorkbookSelect, new object[] { refSheet });

                // record selection and active cell on the sheet we want to select
                _oldSelectionOnRefSheet = Excel(xlfSelection);
                _oldActiveCellOnRefSheet = Excel(xlfActiveCell);

                // make the selection
                Excel(xlcFormulaGoto, refToSelect);
            }

            public void Dispose()
            {
                // Reset the selection on the target sheet
                Excel(xlcSelect, _oldSelectionOnRefSheet, _oldActiveCellOnRefSheet);

                // Reset the sheet originally selected
                var oldActiveSheet = (string)Excel(xlSheetNm, _oldSelectionOnActiveSheet);
                Excel(xlcWorkbookSelect, new object[] { oldActiveSheet });

                // Reset the selection in the active sheet (some bugs make this change sometimes too)
                Excel(xlcSelect, _oldSelectionOnActiveSheet, _oldActiveCellOnActiveSheet);
            }
        }

        #endregion
    }
}
