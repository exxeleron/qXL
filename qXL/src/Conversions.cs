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

#region

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using ExcelDna.Integration;
using qSharp;
using interop = System.Runtime.InteropServices;

#endregion

namespace qXL
{
    public static class Conversions
    {
        private const int MillisPerDay = 86400000;
        private const int ExcelDayDiff = 36526; //Difference in days between q and Excel.        

        private static readonly Dictionary<string, QType> ToQtype = new Dictionary<string, QType>
        {
            {"b", QType.Bool},
            {"g", QType.Guid},
            {"x", QType.Byte},
            {"h", QType.Short},
            {"i", QType.Int},
            {"j", QType.Long},
            {"e", QType.Float},
            {"f", QType.Double},
            {"c", QType.Char},
            {"*", QType.String},
            {"s", QType.Symbol},
            {"m", QType.Month},
            {"d", QType.Date},
            {"z", QType.Datetime},
            {"u", QType.Minute},
            {"v", QType.Second},
            {"t", QType.Time},
            {"p", QType.Timestamp},
            {"n", QType.Timespan}
        };

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Finds the precise type of given COM object.
        /// </summary>
        /// <returns>Type of provided object</returns>
        /// <summary>
        ///     Converts array to Excel-worksheet displayable 2dim array
        /// </summary>
        /// <param name="inArray">array</param>
        /// <returns>2dim array that can be displayed on the worksheet</returns>
        private static object[,] Convert2DimArray(Array inArray)
        {
            if (inArray == null)
            {
                return null;
            }

            var rowsLength = 0;
            var columnsLength = 0;
            Boolean hasNestedArray = false;
            for (var i = inArray.GetLowerBound(0); i <= inArray.GetUpperBound(0); i++)
            {
                var array1 = inArray.GetValue(i) as Array;
                if (array1 != null)
                {
                    var type = array1.GetType().Name.ToLower();
                    if (!type.Equals("char[]"))
                    {
                        hasNestedArray = true;
                        if (columnsLength < array1.GetLength(0))
                        {
                            columnsLength = array1.GetLength(0);
                        }
                        if (array1.Rank > 1)
                        {
                            rowsLength += array1.GetLength(1);
                        }
                        else
                        {
                            rowsLength += 1;
                        }
                    }
                    else
                    {
                        rowsLength += 1;
                    }
                }
                else
                {
                    rowsLength += 1;
                }                   
            }

            var result = hasNestedArray ? new object[rowsLength > 0 ? rowsLength : 1, columnsLength > 0 ? columnsLength : 1] : new object[1, rowsLength];
            var rowIdx = 0;
            var colIdx = 0;
            for (var i = inArray.GetLowerBound(0); i <= inArray.GetUpperBound(0); i++)
            {
                if (hasNestedArray)
                {
                    var array1 = inArray.GetValue(i) as Array;
                    if (array1 != null)
                    {
                        var type = array1.GetType().Name.ToLower();
                        if (!type.Equals("char[]"))
                        {
                            for (var k = array1.GetLowerBound(0); k <= array1.GetUpperBound(0); k++)
                            {
                                result[rowIdx, k] = Convert2Excel(array1.GetValue(k));

                            }
                            if (array1.GetUpperBound(0) < columnsLength)
                            {
                                for (var m = array1.GetUpperBound(0) + 1; m < columnsLength; m++)
                                {
                                    result[rowIdx, m] = "";
                                }
                            }
                            ++rowIdx;
                        }
                        else
                        {
                            result[rowIdx, 0] = Convert2Excel(inArray.GetValue(i));
                            if (1 < columnsLength)
                            {
                                for (var m = 1; m < columnsLength; m++)
                                {
                                    result[rowIdx, m] = "";
                                }
                            }
                            ++rowIdx;
                        }
                    }
                    else
                    {
                        result[rowIdx, 0] = Convert2Excel(inArray.GetValue(i));
                        if (1 < columnsLength)
                        {
                            for (var m = 1; m < columnsLength; m++)
                            {
                                result[rowIdx, m] = "";
                            }
                        }
                        ++rowIdx;
                    }
                }
                else
                {
                    result[0, colIdx] = Convert2Excel(inArray.GetValue(i));
                    ++colIdx;
                }
            }

            return result;
         }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Simply retrieves value of the COM object (workseet ranges are often passed
        ///     as COM objects).
        /// </summary>
        /// <returns>value associated with given COM object</returns>
        /// <summary>
        ///     Performs conversion of given value to given q type
        /// </summary>
        /// <param name="value">value to be converted</param>
        /// <param name="type">type to convert to</param>
        /// <returns>converted value</returns>
        public static object Convert2Q(object value, string type)
        {
            if (value == null || value is string && value.ToString() == "" || ExcelEmpty.Value == value)
            {
                return QTypes.GetQNull(ToQtype[type]);
            }

            if (type == null)
            {
                return value.ToString();
            }

            switch (type.ToLowerInvariant())
            {
                case "b":
                    return Boolean2Q(value);
                case "g":
                    return Guid2Q(value);
                case "x":
                    return Byte2Q(value);
                case "h":
                    return Short2Q(value);
                case "i":
                    return Int2Q(value);
                case "j":
                    return Long2Q(value);
                case "e":
                    return Real2Q(value);
                case "f":
                    return Float2Q(value);
                case "c":
                    return Char2Q(value);
                case "t":
                    return ExcelDate2QTime(value);
                case "d":
                    return ExcelDate2QDate(value);
                case "z":
                    return ExcelDate2QDateTime(value);
                case "p":
                    return ExcelDate2QTimestamp(value);
                case "n":
                    return ExcelDate2QTimespan(value);
                case "m":
                    return ExcelDate2QMonth(value);
                case "v":
                    return ExcelDate2QSecond(value);
                case "u":
                    return ExcelDate2QMinute(value);
                case "*":
                case "s":
                    return Sym2Q(value);
                default:
                    return value.ToString();
            }
        }

        //-------------------------------------------------------------------//
        private static object Boolean2Q(object val)
        {
            try
            {
                var array = val as Array;
                if (array != null)
                {
                    return BooleanArray2Q(array);
                }
                return Convert.ToBoolean(val);
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + val + "> to bool");
            }
        }

        //-------------------------------------------------------------------//
        private static object BooleanArray2Q(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1) //one dimensional array (from VBA)
            {
                var res = new bool[dim1];
                for (var i = 0; i < dim1; i++)
                    res[i] = Convert.ToBoolean(a.GetValue(i));
                return res;
            }
            //two dimensional array (sheet)            
            var dim2 = a.GetLength(1);
            var r = new object[dim2];
            for (var i = 0; i < dim2; i++)
            {
                var elem = new bool[dim1];
                for (var j = 0; j < dim1; j++)
                    elem[j] = Convert.ToBoolean(a.GetValue(j, i));
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object Guid2Q(object val)
        {
            try
            {
                var array = val as Array;
                return array != null
                    ? GuidArray2Q(array)
                    : TypeDescriptor.GetConverter(typeof (Guid)).ConvertFrom(val);
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + val + "> to bool");
            }
        }

        //-------------------------------------------------------------------//
        private static object GuidArray2Q(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1) //one dimensional array (from VBA)
            {
                var res = new Guid[dim1];
                for (var i = 0; i < dim1; i++)
                {
                    var val = a.GetValue(i);
                    var convertFrom = TypeDescriptor.GetConverter(typeof (Guid)).ConvertFrom(val);
                    if (convertFrom != null)
                        res[i] = val == null
                            ? (Guid) QTypes.GetQNull(QType.Guid)
                            : (Guid) convertFrom;
                }
                return res;
            }
            //two dimensional array (sheet)            
            var dim2 = a.GetLength(1);
            var r = new object[dim2];
            for (var i = 0; i < dim2; i++)
            {
                var elem = new Guid[dim1];
                for (var j = 0; j < dim1; j++)
                {
                    var val = a.GetValue(j, i);
                    var convertFrom = TypeDescriptor.GetConverter(typeof (Guid)).ConvertFrom(val);
                    if (convertFrom != null)
                        elem[j] = val == null
                            ? (Guid) QTypes.GetQNull(QType.Guid)
                            : (Guid) convertFrom;
                }
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object ByteArray2Q(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1) //one dimensional array (from VBA)
            {
                var res = new byte[dim1];
                for (var i = 0; i < dim1; i++)
                    res[i] = Convert.ToByte(a.GetValue(i));
                return res;
            }
            //two dimensional array (sheet)            
            var dim2 = a.GetLength(1);
            var r = new object[dim2];
            for (var i = 0; i < dim2; i++)
            {
                var elem = new byte[dim1];
                for (var j = 0; j < dim1; j++)
                    elem[j] = Convert.ToByte(a.GetValue(j, i));
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object Byte2Q(object val)
        {
            try
            {
                var array = val as Array;
                return array != null ? ByteArray2Q(array) : Convert.ToByte(val);
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + val + "> to byte");
            }
        }

        //-------------------------------------------------------------------//
        private static object Short2Q(object val)
        {
            try
            {
                var array = val as Array;
                return array != null ? ShortArray2Q(array) : Convert.ToInt16(val);
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + val + "> to short");
            }
        }

        //-------------------------------------------------------------------//
        private static object ShortArray2Q(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1) //one dimensional array (from VBA)
            {
                var res = new short[dim1];
                for (var i = 0; i < dim1; i++)
                    res[i] = Convert.ToInt16(a.GetValue(i));
                return res;
            }
            //2- dimensional array (from worksheet)            
            var dim2 = a.GetLength(1);
            var r = new object[dim2]; //two dimensional array
            for (var i = 0; i < dim2; i++)
            {
                var elem = new short[dim1];
                for (var j = 0; j < dim1; j++)
                    elem[j] = Convert.ToInt16(a.GetValue(j, i));
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object Int2Q(object val)
        {
            try
            {
                var array = val as Array;
                return array != null ? IntArray2Q(array) : Convert.ToInt32(val);
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + val + "> to int");
            }
        }

        //-------------------------------------------------------------------//
        private static object IntArray2Q(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1) //one dimensional array (from VBA)
            {
                var res = new int[dim1];
                for (var i = 0; i < dim1; i++)
                    res[i] = Convert.ToInt32(a.GetValue(i));
                return res;
            }
            //2-dimensional array            
            var dim2 = a.GetLength(1);
            var r = new object[dim2]; //two dimensional array
            for (var i = 0; i < dim2; i++)
            {
                var elem = new int[dim1];
                for (var j = 0; j < dim1; j++)
                    elem[j] = Convert.ToInt32(a.GetValue(j, i));
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object Long2Q(object val)
        {
            try
            {
                var array = val as Array;
                return array != null ? LongArray2Q(array) : Convert.ToInt64(val);
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + val + "> to long");
            }
        }

        //-------------------------------------------------------------------//
        private static object LongArray2Q(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1) //one dimensional array
            {
                var res = new long[dim1];
                for (var i = 0; i < dim1; i++)
                    res[i] = Convert.ToInt64(a.GetValue(i));
                return res;
            }
            var dim2 = a.GetLength(1);
            var r = new object[dim2]; //two dimensional array
            for (var i = 0; i < dim2; i++)
            {
                var elem = new long[dim1];
                for (var j = 0; j < dim1; j++)
                    elem[j] = Convert.ToInt64(a.GetValue(j, i));
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object Real2Q(object val)
        {
            try
            {
                var array = val as Array;
                return array != null ? RealArray2Q(array) : Convert.ToSingle(val);
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + val + "> to real");
            }
        }

        //-------------------------------------------------------------------//
        private static object RealArray2Q(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1) //one dimensional array
            {
                var res = new float[dim1];
                for (var i = 0; i < dim1; i++)
                    res[i] = Convert.ToSingle(a.GetValue(i));
                return res;
            }
            var dim2 = a.GetLength(1);
            var r = new object[dim2]; //two dimensional array
            for (var i = 0; i < dim2; i++)
            {
                var elem = new float[dim1];
                for (var j = 0; j < dim1; j++)
                    elem[j] = Convert.ToSingle(a.GetValue(j, i));
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object Float2Q(object val)
        {
            try
            {
                var array = val as Array;
                return array != null ? FloatArray2Q(array) : Convert.ToDouble(val);
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + val + "> to float");
            }
        }

        //-------------------------------------------------------------------//
        private static object FloatArray2Q(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1)
            {
                var res = new double[dim1];
                for (var i = 0; i < dim1; i++)
                    res[i] = Convert.ToDouble(a.GetValue(i));
                return res;
            }
            var dim2 = a.GetLength(1);
            var r = new object[dim2]; //two dimensional array
            for (var i = 0; i < dim2; i++)
            {
                var elem = new double[dim1];
                for (var j = 0; j < dim1; j++)
                    elem[j] = Convert.ToDouble(a.GetValue(j, i));
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object Char2Q(object val)
        {
            try
            {
                var array = val as Array;
                if (array != null)
                {
                    return CharArray2Q(array);
                }
                return Convert.ToChar(val);
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + val + "> to char");
            }
        }

        //-------------------------------------------------------------------//
        private static object CharArray2Q(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1)
            {
                var res = new char[dim1];
                for (var i = 0; i < dim1; i++)
                    res[i] = Convert.ToChar(a.GetValue(i).ToString());
                return res;
            }
            var dim2 = a.GetLength(1);
            var r = new object[dim2]; //two dimensional array
            for (var i = 0; i < dim2; i++)
            {
                var elem = new char[dim1];
                for (var j = 0; j < dim1; j++)
                    elem[j] = Convert.ToChar(a.GetValue(j, i).ToString());
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object ExcelDate2QTime(object date)
        {
            try
            {
                var array = date as Array;
                if (array != null)
                {
                    return ExcelDateArray2QTime(array);
                }
                if (date is DateTime)
                {
                    return new QTime((DateTime) date);
                }
                if (date is string)
                {
                    return new QTime(AdjToExcelDate(DateTime.Parse(date as string)));
                }
                var d = Convert.ToDouble(date);
                return new QTime(Convert.ToInt32(MillisPerDay*(d - (int) d)));
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + date + "> to QTime");
            }
        }

        //-------------------------------------------------------------------//
        private static object ExcelDateArray2QTime(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1)
            {
                var res = new QTime[dim1];
                for (var i = 0; i < dim1; i++)
                    res[i] = (QTime) ExcelDate2QTime(a.GetValue(i));
                return res;
            }
            var dim2 = a.GetLength(1);
            var r = new object[dim2]; //two dimensional array
            for (var i = 0; i < dim2; i++)
            {
                var elem = new QTime[dim1];
                for (var j = 0; j < dim1; j++)
                    elem[j] = (QTime) ExcelDate2QTime(a.GetValue(j, i));
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object ExcelDate2QDate(object date)
        {
            try
            {
                var array = date as Array;
                if (array != null)
                {
                    return ExcelDateArray2QDate(array);
                }
                if (date is DateTime)
                {
                    return new QDate(AdjToExcelDate((DateTime) date));
                }
                if (date is string)
                {
                    return new QDate(AdjToExcelDate(DateTime.Parse(date as string)));
                }
                if (date is int || date is short)
                {
                    return new QDate(Convert.ToInt32(date));
                }

                var v = Convert.ToDouble(date);
                var d = (int) v;
                return new QDate(d > 1 ? d - ExcelDayDiff : d + 1 - ExcelDayDiff);
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + date + "> to QDate");
            }
        }

        //-------------------------------------------------------------------//
        private static object ExcelDateArray2QDate(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1)
            {
                var res = new QDate[dim1];
                for (var i = 0; i < dim1; i++)
                    res[i] = (QDate) ExcelDate2QDate(a.GetValue(i));
                return res;
            }
            var dim2 = a.GetLength(1);
            var r = new object[dim2]; //two dimensional array
            for (var i = 0; i < dim2; i++)
            {
                var elem = new QDate[dim1];
                for (var j = 0; j < dim1; j++)
                    elem[j] = (QDate) ExcelDate2QDate(a.GetValue(j, i));
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object ExcelDateArray2QDateTime(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1)
            {
                var res = new QDateTime[dim1];
                for (var i = 0; i < dim1; i++)
                    res[i] = (QDateTime) ExcelDate2QDateTime(a.GetValue(i));
                return res;
            }
            var dim2 = a.GetLength(1);
            var r = new object[dim2]; //two dimensional array
            for (var i = 0; i < dim2; i++)
            {
                var elem = new QDateTime[dim1];
                for (var j = 0; j < dim1; j++)
                    elem[j] = (QDateTime) ExcelDate2QDateTime(a.GetValue(j, i));
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object ExcelDate2QDateTime(object date)
        {
            try
            {
                var array = date as Array;
                if (array != null)
                {
                    return ExcelDateArray2QDateTime(array);
                }
                if (date is DateTime)
                {
                    return new QDateTime(AdjToExcelDate((DateTime) date));
                }
                if (date is string)
                {
                    return new QDateTime(AdjToExcelDate(DateTime.Parse(date as string)));
                }
                if (date is int || date is short)
                {
                    return new QDateTime(Convert.ToDouble(date));
                }
                var v = Convert.ToDouble(date);
                var d = (int) v;
                return new QDateTime((d > 1 ? d - ExcelDayDiff : d + 1 - ExcelDayDiff) + (v - d));
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + date + "> to QDateTime");
            }
        }

        //-------------------------------------------------------------------//
        private static object ExcelDate2QTimestamp(object date)
        {
            try
            {
                var array = date as Array;
                if (array != null)
                {
                    return ExcelDateArray2QTimestamp(array);
                }
                if (date is DateTime)
                {
                    return new QTimestamp(AdjToExcelDate((DateTime) date));
                }
                if (date is string)
                {
                    return new QTimestamp(AdjToExcelDate(DateTime.Parse(date as string)));
                }
                var v = Convert.ToDouble(date);
                var d = (int) v;
                d = d > 1 ? d - ExcelDayDiff : d + 1 - ExcelDayDiff;
                v = MillisPerDay*((v - (int) v) + d);
                return new QTimestamp(Convert.ToInt64(1e6*v));
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + date + "> to QTimestamp");
            }
        }

        //-------------------------------------------------------------------//
        private static object ExcelDateArray2QTimestamp(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1)
            {
                var res = new QTimestamp[dim1];
                for (var i = 0; i < dim1; i++)
                    res[i] = (QTimestamp) ExcelDate2QTimestamp(a.GetValue(i));
                return res;
            }
            var dim2 = a.GetLength(1);
            var r = new object[dim2]; //two dimensional array
            for (var i = 0; i < dim2; i++)
            {
                var elem = new QTimestamp[dim1];
                for (var j = 0; j < dim1; j++)
                    elem[j] = (QTimestamp) ExcelDate2QTimestamp(a.GetValue(j, i));
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object ExcelDate2QTimespan(object date)
        {
            try
            {
                var array = date as Array;
                if (array != null)
                {
                    return ExcelDateArray2QTimespan(array);
                }
                if (date is DateTime)
                {
                    return new QTimespan((DateTime) date);
                }
                if (date is string)
                {
                    return new QTimespan(AdjToExcelDate(DateTime.Parse(date as string)));
                }
                var v = Convert.ToDouble(date);
                v = MillisPerDay*(v - (int) v);
                return new QTimespan(Convert.ToInt64(1e6*v));
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + date + "> to QTimespan");
            }
        }

        //-------------------------------------------------------------------//
        private static object ExcelDateArray2QTimespan(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1)
            {
                var res = new QTimespan[dim1];
                for (var i = 0; i < dim1; i++)
                    res[i] = (QTimespan) ExcelDate2QTimespan(a.GetValue(i));
                return res;
            }
            var dim2 = a.GetLength(1);
            var r = new object[dim2]; //two dimensional array
            for (var i = 0; i < dim2; i++)
            {
                var elem = new QTimespan[dim1];
                for (var j = 0; j < dim1; j++)
                    elem[j] = (QTimespan) ExcelDate2QTimespan(a.GetValue(j, i));
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object ExcelDate2QSecond(object date)
        {
            try
            {
                var array = date as Array;
                if (array != null)
                {
                    return ExcelDateArray2QSecond(array);
                }
                if (date is DateTime)
                {
                    return new QSecond((DateTime) date);
                }
                if (date is string)
                {
                    return new QSecond(AdjToExcelDate(DateTime.Parse(date as string)));
                }
                var v = Convert.ToDouble(date);
                return new QSecond(Convert.ToInt32(MillisPerDay/1000*(v - (int) v)));
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + date + "> to QSecond");
            }
        }

        //-------------------------------------------------------------------//
        private static object ExcelDateArray2QSecond(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1)
            {
                var res = new QSecond[dim1];
                for (var i = 0; i < dim1; i++)
                    res[i] = (QSecond) ExcelDate2QSecond(a.GetValue(i));
                return res;
            }
            var dim2 = a.GetLength(1);
            var r = new object[dim2]; //two dimensional array
            for (var i = 0; i < dim2; i++)
            {
                var elem = new QSecond[dim1];
                for (var j = 0; j < dim1; j++)
                {
                    elem[j] = (QSecond) ExcelDate2QSecond(a.GetValue(j, i));
                }
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object ExcelDate2QMinute(object date)
        {
            try
            {
                var array = date as Array;
                if (array != null)
                {
                    return ExcelDateArray2QMinute(array);
                }
                if (date is DateTime)
                {
                    return new QMinute((DateTime) date);
                }
                if (date is string)
                {
                    return new QMinute(AdjToExcelDate(DateTime.Parse(date as string)));
                }

                var tm = DateTime.FromOADate(Convert.ToDouble(date));
                return new QMinute(new DateTime(tm.Year, tm.Month, tm.Day, tm.Hour, tm.Minute, 0, 0));
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + date + "> to QMinute");
            }
        }

        //-------------------------------------------------------------------//
        private static object ExcelDateArray2QMinute(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1)
            {
                var res = new QMinute[dim1];
                for (var i = 0; i < dim1; i++)
                    res[i] = (QMinute) ExcelDate2QMinute(a.GetValue(i));
                return res;
            }
            var dim2 = a.GetLength(1);
            var r = new object[dim2]; //two dimensional array
            for (var i = 0; i < dim2; i++)
            {
                var elem = new QMinute[dim1];
                for (var j = 0; j < dim1; j++)
                {
                    elem[j] = (QMinute) ExcelDate2QMinute(a.GetValue(j, i));
                }
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object ExcelDate2QMonth(object date)
        {
            try
            {
                var array = date as Array;
                if (array != null)
                {
                    return ExcelDateArray2QMonth(array);
                }

                if (date is DateTime || date is string)
                {
                    var d = date is string
                        ? DateTime.Parse(date as string)
                        : ((DateTime) date).AddDays(-1*((DateTime) date).Day);
                    return new QMonth(AdjToExcelDate(d));
                    //return new QMonth(AdjToExcelDate((DateTime)date));
                }
                if (date is int || date is short)
                {
                    return new QMonth(Convert.ToInt32(date));
                }

                return new QMonth(AdjToExcelDate(DateTime.FromOADate(Convert.ToDouble(date))));
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + date + "> to QMonth");
            }
        }

        //-------------------------------------------------------------------//
        private static object ExcelDateArray2QMonth(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1)
            {
                var res = new QMonth[dim1];
                for (var i = 0; i < dim1; i++)
                {
                    res[i] = (QMonth) ExcelDate2QMonth(a.GetValue(i));
                }
                return res;
            }
            var dim2 = a.GetLength(1);
            var r = new object[dim2]; //two dimensional array
            for (var i = 0; i < dim2; i++)
            {
                var elem = new QMonth[dim1];
                for (var j = 0; j < dim1; j++)
                {
                    elem[j] = (QMonth) ExcelDate2QMonth(a.GetValue(j, i));
                }
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        private static object Sym2Q(object val)
        {
            try
            {
                var array = val as Array;
                return array != null ? SymArray2Q(array) : val.ToString();
            }
            catch (Exception)
            {
                throw new ConversionException("Cannot convert: <" + val + "> to symbol");
            }
        }

        //-------------------------------------------------------------------//
        private static object SymArray2Q(Array a)
        {
            var dim1 = a.GetLength(0);
            if (a.Rank == 1)
            {
                var res = new string[dim1];
                for (var i = 0; i < dim1; i++)
                {
                    res[i] = a.GetValue(i).ToString();
                }
                return res;
            }
            var dim2 = a.GetLength(1);
            var r = new object[dim2]; //two dimensional array
            for (var i = 0; i < dim2; i++)
            {
                var elem = new string[dim1];
                for (var j = 0; j < dim1; j++)
                {
                    elem[j] = a.GetValue(j, i).ToString();
                }
                r[i] = elem;
            }
            return r;
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     This function is just an adjustment function to handle 1900 year bug in spreadsheet software.
        /// </summary>
        /// <param name="dt"></param>
        /// <returns>date adjusted by 1 day if necessary</returns>
        private static DateTime AdjToExcelDate(DateTime dt)
        {
            return dt.Year == 1899 ? dt.AddDays(1) : dt;
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Converts data received from q process to data types that are displayable on
        ///     the worksheet.
        /// </summary>
        /// <param name="result">value to be converted</param>
        /// <returns>value displayable on the worksheet</returns>
        internal static object Convert2Excel(object result)
        {
            if (result == null)
            {
                return "";
            }
            var type = result.GetType().Name.ToLower();
            switch (type)
            {
                case "int16":
                    return (((short)result) == (short)QTypes.GetQNull(QType.Short)) ? "" : result;
                case "int32":
                    return (((int)result) == (int)QTypes.GetQNull(QType.Int)) ? "" : result;
                case "int64":
                case "long":
                    if ((((long) result) == (long) QTypes.GetQNull(QType.Long)))
                    {
                        return "";
                    }
                    return Environment.Is64BitProcess ? result : Convert.ToDouble(result);                 
                case "double":
                    return Double.IsNaN((double)result) ? "" : result;
                case "single":
                    return Single.IsNaN(((float)result)) ? "" : result;
                case "boolean":
                case "byte":
                    return result;
                case "guid":
                    return (((Guid) result) == (Guid) QTypes.GetQNull(QType.Guid)) ? "" : result.ToString();
                case "char":
                    return (((char) result) == (char) QTypes.GetQNull(QType.Char)) ? "" : result.ToString();
                case "string":
                    return result.ToString() == QTypes.GetQNull(QType.Symbol).ToString()
                        ? ""
                        : result;
                case "char[]":
                    return ((char[]) result).Length > 0 ? new string((char[]) result) : "";
                case "qlambda":
                    return ((QLambda) result).Expression;
                case "qtimestamp":
                {
                    if (!((QTimestamp) result).Equals(((QTimestamp) QTypes.GetQNull(QType.Timestamp))))
                    {
                        return ((QTimestamp) result).ToDateTime();
                    }
                    return "";
                }
                case "qdatetime":
                {
                    if (!((QDateTime) result).Equals(QTypes.GetQNull(QType.Datetime)))
                    {
                        return ((QDateTime) result).ToDateTime();
                    }
                    return "";
                }
                case "qtime":
                {
                    if (!((QTime) result).Equals(QTypes.GetQNull(QType.Time)))
                    {
                        return GetTime((QTime) result);
                    }
                    return "";
                }
                case "qdate":
                {
                    if (!((QDate) result).Equals(QTypes.GetQNull(QType.Date)))
                    {
                        return ((QDate) result).ToDateTime();
                    }
                    return "";
                }
                case "qtimespan":
                {
                    if (!((QTimespan) result).Equals(QTypes.GetQNull(QType.Timespan)))
                    {
                        return ((QTimespan) result).ToDateTime();
                    }
                    return "";
                }
                case "qsecond":
                {
                    if (!((QSecond) result).Equals(QTypes.GetQNull(QType.Second)))
                    {
                        return GetTime((QSecond) result);
                    }

                    return "";
                }
                case "qmonth":
                {
                    if (!((QMonth) result).Equals(QTypes.GetQNull(QType.Month)))
                    {
                        return ((QMonth) result).ToDateTime();
                    }
                    return "";
                }
                case "qminute":
                {
                    if (!((QMinute) result).Equals(QTypes.GetQNull(QType.Minute)))
                    {
                        return GetTime((QMinute) result);
                    }

                    return "";
                }
                case "qdictionary":
                    return QDict2Excel((QDictionary) result);
                case "qtable":
                    return QTable2Excel((QTable) result);
                case "qkeyedtable":
                    return QKeyedTable2Excel((QKeyedTable) result);
                default:
                    if (type.Contains("[]"))
                    {
                        return Convert2DimArray(result as Array);
                    }
                    return result.ToString();
            }
        }

        private static DateTime GetTime(IDateTime time)
        {
            var t = time.ToDateTime();
            return new DateTime(1, 1, 1, t.Hour, t.Minute, t.Second, t.Millisecond, t.Kind);
        }

        //-------------------------------------------------------------------//
        private static object[,] QDict2Excel(QDictionary dict)
        {
            try
            {
                var len = Utils.GetMaxDictSize(dict);
                var res = new object[len + 1, dict.Keys.Length];
                var keyCounter = 0;
                foreach (QDictionary.KeyValuePair kv in dict)
                {
                    var k = Convert2Excel(kv.Key);
                    if (k is Array)
                    {
                        k = (from object x in ((Array) k) where x != null select x).Aggregate("", (current, x) => current + (" " + x)).Trim();
                    }
                    res[0, keyCounter] = k;//kv.Key.ToString();
                    var array = kv.Value as Array;
                    var type = kv.Value.GetType().Name.ToLower();
                    if (array != null && !type.Equals("char[]"))
                    {
                        for (var i = 0; i < array.Length; i++)
                        {
                            res[i + 1, keyCounter] = Convert2Excel(array.GetValue(i));
                        }
                        if (array.Length + 1 < res.GetLength(0))
                        {
                            for (var i = array.Length; i < res.GetLength(0) - 1; i++)
                            {
                                res[i + 1, keyCounter] = ""; // remove null elements
                            }
                        }
                    }
                    else
                    {
                        res[1, keyCounter] = Convert2Excel(kv.Value);
                        if (1 < res.GetLength(0))
                        {
                            for (var i = 1; i < res.GetLength(0) - 1; i++)
                            {
                                res[i + 1, keyCounter] = ""; // remove null elements
                            }
                        }
                    }
                    keyCounter++;
                }
                return res;
            }
            catch (ConversionException e)
            {
                return new object[,] {{e.Message}};
            }
        }

        //-------------------------------------------------------------------//
        private static object[,] QTable2Excel(QTable table)
        {
            var res = new object[table.RowsCount + 1, table.ColumnsCount];
            for (var i = 0; i < table.ColumnsCount; i++)
                for (var j = 0; j < table.RowsCount + 1; j++)
                {
                    if (j == 0) //write header
                    {
                        res[j, i] = table.Columns.GetValue(i).ToString();
                    }
                    else //write data
                    {
                        res[j, i] = Convert2Excel(((Array) table.Data.GetValue(i)).GetValue(j - 1));
                    }
                }
            return res;
        }

        //-------------------------------------------------------------------//
        private static object[,] QKeyedTable2Excel(QKeyedTable table)
        {
            var res = new object[table.Values.RowsCount + 1, table.Values.ColumnsCount + table.Keys.ColumnsCount];
            for (var i = 0; i < table.Keys.ColumnsCount + table.Values.ColumnsCount; i++)
                for (var j = 0; j < table.Values.RowsCount + 1; j++)
                {
                    if (j == 0) //write colnames
                    {
                        res[j, i] = (i < table.Keys.ColumnsCount)
                            ? table.Keys.Columns.GetValue(i).ToString()
                            : table.Values.Columns.GetValue(i - table.Keys.ColumnsCount).ToString();
                    }
                    else //write data
                    {
                        res[j, i] = (i < table.Keys.ColumnsCount)
                            ? Convert2Excel(((Array) table.Keys.Data.GetValue(i)).GetValue(j - 1))
                            : Convert2Excel(
                                ((Array) table.Values.Data.GetValue(i - table.Keys.ColumnsCount)).GetValue(
                                    j - 1));
                    }
                }
            return res;
        }
    }
}
