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
using System.Collections.Concurrent;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using ExcelDna.Integration;
using qSharp;

#endregion

namespace qXL
{
    [ComVisible(true)]
    // ReSharper disable InconsistentNaming
    public class qXLShared
        // ReSharper restore InconsistentNaming
    {
        private const string ConversionMarker = "7097B31CF26749BA9839C996B19598FC";

        private const string ErrLengthMismatch =
            "Provided format specification has different length than the range to be converted.";

        private const string ErrCol2ValMismatch = "Incompatible lengths between column names and data.";

        private static readonly ConcurrentDictionary<string, QConnection> Connections =
            new ConcurrentDictionary<string, QConnection>();

        private static readonly ConcurrentDictionary<string, object> Conversions =
            new ConcurrentDictionary<string, object>();

        // ReSharper disable InconsistentNaming
        public object qOpen(string alias, string hostname, object port, string username = null, string password = null)
            // ReSharper restore InconsistentNaming
        {
            try
            {
                if (String.IsNullOrEmpty(alias))
                {
                    return "Invalid alias";
                }
                var c = GetConnection(alias);
                if (c != null)
                {
                    return alias;
                }

                if (String.IsNullOrEmpty(hostname))
                {
                    return "Invalid hostname";
                }

                int prt;
                try
                {
                    prt = Int32.Parse(port.ToString());
                }
                catch
                {
                    return "Invalid port";
                }

                try
                {
                    c = new QBasicConnection(hostname, prt, username, password);
                    c.Open();
                    Connections[alias] = c;
                }
                catch (QException e)
                {
                    return "ERR: " + e.Message;
                }
            }
            catch (Exception e)
            {
                return "ERR: " + e.Message;
            }
            return alias;
        }

        // ReSharper disable InconsistentNaming
        public object qClose(string alias)
            // ReSharper restore InconsistentNaming
        {
            try
            {
                if (Connections.ContainsKey(alias))
                {
                    QConnection con;
                    if (Connections.TryRemove(alias, out con))
                    {
                        con.Close();
                    }
                    return "Closed";
                }
            }
            catch (Exception e)
            {
                return "ERR: " + e.Message;
            }
            return String.Format("Unknown alias '{0}'", alias);
        }

        // ReSharper disable InconsistentNaming
        public void qCloseAll()
            // ReSharper restore InconsistentNaming
        {
            var keys = Connections.Keys.ToArray();
            foreach (var key in keys)
            {
                qClose(key);
            }
        }

        // ReSharper disable InconsistentNaming
        public object qQuery(string alias, object query,
            // ReSharper restore InconsistentNaming
            object p1 = null, object p2 = null, object p3 = null, object p4 = null,
            object p5 = null, object p6 = null, object p7 = null, object p8 = null)
        {
            try
            {
                object[] parameters = {p1, p2, p3, p4, p5, p6, p7, p8};
                //take only provided parameters (skip the ones that are not set)
                parameters =
                    parameters.Where(
                        x => x != null && x.GetType() != typeof (ExcelMissing) && x.GetType() != typeof (ExcelError)
                             && x.GetType() != typeof (ExcelEmpty)).ToArray();


                //to translate possible ranges to their values
                for (var i = 0; i < parameters.Length; i++)
                {
                    parameters[i] = qConvert(parameters[i]);
                }

                var c = GetConnection(alias);
                if (c == null)
                {
                    return String.Format(@"Alias '{0}' not defined", alias);
                }

                var result = (parameters.Length > 0)
                    ? c.Sync(query.ToString(), parameters)
                    : c.Sync(query.ToString());

                if (result == null)
                {
                    return null; //null gets returned only when function definition has been sent to q.
                }
                var array = qXL.Conversions.Convert2Excel(result);
                return array ?? result;
            }
            catch (IOException io)
            {
                //this normally means that the process has been terminated on the receiving site
                //so clear the connection alias.
                //Connections.Remove(alias);
                return "ERR: " + io.Message;
            }
            catch (Exception e)
            {
                return "ERR: " + e.Message;
            }
        }

        public object[] qAtom(object value, object type)
        {
            var convKey = GenerateRandomString();
            try
            {
                Conversions.TryAdd(convKey, qXL.Conversions.Convert2Q(value, GetTypeString(type)));
                return new object[] {convKey, ConversionMarker, null};
            }
            catch (Exception e)
            {
                return new object[] {convKey, ConversionMarker, e.Message};
            }
        }

        // ReSharper disable InconsistentNaming
        public object[] qList(object value, object types, bool forceListOfLists = false)
            // ReSharper restore InconsistentNaming
        {
            var convKey = GenerateRandomString();

            var v = value as Array ?? new[] {value};

            try
            {
                var ts = GetTypeString(types);
                object[] res = null;
                var format = ts.ToCharArray();
                if (v.Rank == 1)
                {
                    if (format.Length != v.GetLength(0) && format.Length != 1)
                    {
                        return new object[] {convKey, ConversionMarker, ErrLengthMismatch};
                    }

                    res = new object[v.GetLength(0)];

                    if (format.Length == v.GetLength(0))
                    {
                        for (var i = v.GetLowerBound(0); i <= v.GetUpperBound(0); i++)
                            res[i - v.GetLowerBound(0)] = qXL.Conversions.Convert2Q(v.GetValue(i),
                                format[i - v.GetLowerBound(0)].ToString(CultureInfo.InvariantCulture));
                    }
                    else if (format.Length == 1)
                    {
                        for (var i = v.GetLowerBound(0); i <= v.GetUpperBound(0); i++)
                            res[i - v.GetLowerBound(0)] = qXL.Conversions.Convert2Q(v.GetValue(i),
                                format[0].ToString(CultureInfo.InvariantCulture));
                    }
                }
                if (v.Rank == 2)
                {
                    // list of lists
                    if (v.GetLength(0) > 1 && v.GetLength(1) > 1)
                    {
                        res = new object[v.GetLength(1)];
                        for (var i = v.GetLowerBound(1); i <= v.GetUpperBound(1); i++)
                        {
                            var vector = new object[v.GetLength(0)];
                            for (var j = v.GetLowerBound(0); j <= v.GetUpperBound(0); j++)
                            {
                                if (format.Length != 1)
                                {
                                    vector[j - v.GetLowerBound(0)] = qXL.Conversions.Convert2Q(v.GetValue(j, i),
                                        format[i - v.GetLowerBound(1)].ToString(CultureInfo.InvariantCulture));
                                }
                                else
                                {
                                    vector[j - v.GetLowerBound(0)] = qXL.Conversions.Convert2Q(v.GetValue(j, i),
                                        format[0].ToString(CultureInfo.InvariantCulture));
                                }
                            }
                            res[i - v.GetLowerBound(1)] = vector;
                        }
                    }
                    else
                    {
                        var idx = v.GetLength(0) == 1 ? 1 : 0;
                        if (format.Length != v.GetLength(idx) && format.Length != 1)
                        {
                            return new object[] {convKey, ConversionMarker, ErrLengthMismatch};
                        }
                        res = new object[v.GetLength(idx)];

                        for (var i = v.GetLowerBound(idx); i <= v.GetUpperBound(idx); i++)
                        {
                            var converted =
                                qXL.Conversions.Convert2Q(
                                    v.GetValue(idx == 0 ? i : v.GetLowerBound(1), idx == 0 ? v.GetLowerBound(1) : i),
                                    format.Length != 1
                                        ? format[i - v.GetLowerBound(idx)].ToString(CultureInfo.InvariantCulture)
                                        : format[0].ToString(CultureInfo.InvariantCulture));
                            res[i - v.GetLowerBound(idx)] = forceListOfLists ? new[] {converted} : converted;
                        }
                    }
                }
                Conversions.TryAdd(convKey, res);
                return new object[] {convKey, ConversionMarker, null};
            }
            catch (Exception e)
            {
                return new object[] {convKey, ConversionMarker, e.Message};
            }
        }

        // ReSharper disable InconsistentNaming
        public object[] qDict(object keys, object values, object types)
            // ReSharper restore InconsistentNaming
        {
            if (!(values is Array))
                values = new[] {values};

            if ((values as Array).Rank != 2 && !(keys is Array))
                values = new[] {values};

            if (!(keys is Array) && keys != null)
                keys = new[] {keys};

            var array = keys as Array;
            if (array != null && array.Rank == 2) //thats how it comes from worksheet.
            {
                var k = keys as Array;
                if (k.GetLength(0) == 1 || k.GetLength(1) == 1)
                {
                    keys = Utils.Com2DArray2Array(k);
                }
            }

            var res = qList(values, types, true);

            if (res[2] != null)
                return res;

            var dict = new QDictionary(keys as Array, Conversions[res[0].ToString()] as Array);
            Conversions[res[0].ToString()] = dict;
            return res;
        }

        // ReSharper disable InconsistentNaming
        public object[] qTable(object columnNames, object values, object types, object keys)
            // ReSharper restore InconsistentNaming
        {
            var array = values as Array;
            if (array != null && (array.Rank != 2 && !(columnNames is Array)))
                values = new[] {values};

            if (!(columnNames is Array) && columnNames != null)
                columnNames = new[] {columnNames};

            if (keys != null && keys.GetType() == typeof (ExcelMissing))
            {
                keys = null;
            }

            var array1 = columnNames as Array;
            if (array1 != null && array1.Rank == 2) //thats how it comes from worksheet.
            {
                var k = columnNames as Array;
                if (k.GetLength(0) == 1 || k.GetLength(1) == 1)
                {
                    columnNames = Utils.Com2DArray2Array(k);
                }
            }
            var res = qList(values, types, true);
            if (res[2] != null) //there was an exception converting values,so no point to attempt table creation.
            {
                return res;
            }

            var array2 = columnNames as Array;
            var array3 = Conversions[res[0].ToString()] as Array;
            if (array3 != null && (array2 != null && array2.Length != array3.Length))
            {
                res[2] = ErrCol2ValMismatch;
                return res;
            }

            if (keys != null)
            {
                if (!(keys is Array))
                    keys = new[] {keys};

                if (columnNames == null) return res;
                var tab = new QKeyedTable((columnNames as Array).OfType<string>().ToArray(),
                    (keys as Array).OfType<string>().ToArray(), array3);

                Conversions[res[0].ToString()] = tab;
            }
            else
            {
                if (columnNames == null) return res;
                var tab = new QTable((columnNames as Array).OfType<string>().ToArray(), array3);

                Conversions[res[0].ToString()] = tab;
            }
            return res;
        }

        public object qConvert(object value)
        {
            var array = value as Array;
            if (array == null || array.Length != 3)
                return value;

            switch (array.Rank)
            {
                case 1:
                    if (array.GetValue(array.GetLowerBound(0) + 1).ToString() == ConversionMarker)
                    {
                        if (array.GetValue(array.GetLowerBound(0) + 2) != null &&
                            array.GetValue(array.GetLowerBound(0), array.GetLowerBound(1) + 2) != ExcelEmpty.Value)
                        {
                            throw new ConversionException(
                                array.GetValue(array.GetLowerBound(0), array.GetLowerBound(1) + 2).ToString());
                        }
                        var res = Conversions[array.GetValue(array.GetLowerBound(0)).ToString()];
                        object val;
                        Conversions.TryRemove(array.GetValue(array.GetLowerBound(0)).ToString(), out val);
                        //clean conversions dictionary

                        return res;
                    }
                    break;
                case 2:
                    if (array.GetValue(array.GetLowerBound(0), array.GetLowerBound(1) + 1).ToString() ==
                        ConversionMarker)
                    {
                        if (array.GetValue(array.GetLowerBound(0), array.GetLowerBound(1) + 2) != null &&
                            array.GetValue(array.GetLowerBound(0), array.GetLowerBound(1) + 2) != ExcelEmpty.Value)
                        {
                            throw new ConversionException(
                                array.GetValue(array.GetLowerBound(0), array.GetLowerBound(1) + 2).ToString());
                        }
                        var res = Conversions[array.GetValue(array.GetLowerBound(0), array.GetLowerBound(1)).ToString()];
                        object val;
                        Conversions.TryRemove(
                            array.GetValue(array.GetLowerBound(0), array.GetLowerBound(1)).ToString(), out val);
                        //clean conversions dictionary
                        return res;
                    }
                    break;
            }
            return value;
        }

        private static string GetTypeString(object type)
        {
            string ts;
            var v = type as Array;
            if (v != null)
            {
                var sb = new StringBuilder();
                switch (v.Rank)
                {
                    case 1:
                        for (var i = v.GetLowerBound(0); i <= v.GetUpperBound(0); i++)
                        {
                            sb.Append(v.GetValue(i));
                        }
                        break;
                    case 2:
                        for (var i = v.GetLowerBound(0); i <= v.GetUpperBound(0); i++)
                        {
                            for (var j = v.GetLowerBound(1); j <= v.GetUpperBound(1); j++)
                            {
                                sb.Append(v.GetValue(i, j));
                            }
                        }
                        break;
                }
                ts = sb.ToString().ToLower();
            }
            else
            {
                ts = type != null ? type.ToString().ToLower() : null;
            }
            return ts;
        }

        // ReSharper disable InconsistentNaming
        public string qXLAbout()
            // ReSharper restore InconsistentNaming
        {
            var attributes = Assembly.GetExecutingAssembly()
                .GetCustomAttributes(typeof (AssemblyProductAttribute), false);

            AssemblyProductAttribute attribute = null;
            if (attributes.Length > 0)
            {
                attribute = attributes[0] as AssemblyProductAttribute;
            }

            return (attribute == null ? Assembly.GetExecutingAssembly().GetName().Name : attribute.Product) + " " +
                   Assembly.GetExecutingAssembly().GetName().Version;
        }

        //------------------------------------------------------------------//
        /// <summary>
        ///     Generates new GUID which will be used as a uniqe
        ///     key for signing the converted data structures.
        /// </summary>
        /// <returns>random string </returns>
        private static string GenerateRandomString()
        {
            return Guid.NewGuid().ToString();
        }

        private static QConnection GetConnection(string alias)
        {
            if (!Connections.ContainsKey(alias)) return null;
            var c = Connections[alias];
            if (c == null)
            {
                return null;
            }
            if (c.IsConnected())
            {
                return c;
            }
            // reconnect
            c.Close();
            c.Open();
            return c;
        }
    }
}