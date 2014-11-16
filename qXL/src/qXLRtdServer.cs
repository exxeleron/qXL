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
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using qSharp;

#endregion

namespace qXL
{
    [Guid("7CC378E1-8E12-4FB3-B9E4-556A27950F83"), ProgId("qXLRTD"), ComVisible(true)]
    // ReSharper disable InconsistentNaming
    public class qXLRtdServer : ExcelRtdServer
        // ReSharper restore InconsistentNaming

    {
        //mapping between aliases and actual Q connections
        private const int AliasPos = 0;
        private const int TabPos = 1;
        private const int SymPos = 2;
        private const int ColPos = 3;
        private const int HisPos = 4;
        private const int SymbolIndetifier = 5;

        private const string BackTick = "`";
        private static readonly object LockThis = new object();

        private static readonly ConcurrentDictionary<string, QCallbackConnection> Connections =
            new ConcurrentDictionary<string, QCallbackConnection>();

        private static readonly ConcurrentDictionary<string, TopicMap> WildCardMapping =
            new ConcurrentDictionary<string, TopicMap>();

        //mapping between connection alias and the TopicMap
        private static readonly ConcurrentDictionary<string, TopicMap> Mapping =
            new ConcurrentDictionary<string, TopicMap>();

        private static readonly ConcurrentDictionary<string, Dictionary<string, SortedDictionary<string, string>>>
            IdSymMap = new ConcurrentDictionary<string, Dictionary<string, SortedDictionary<string, string>>>();

        private static readonly ConcurrentDictionary<string, Dictionary<string, Dictionary<string, string>>> SymIdMap =
            new ConcurrentDictionary<string, Dictionary<string, Dictionary<string, string>>>();

        private static readonly DataCache Cache = new DataCache(1); //by default we keep only previous value.

        private static readonly ConcurrentDictionary<TopicInfo, object> DataOut =
            new ConcurrentDictionary<TopicInfo, object>();

        private static readonly ConcurrentDictionary<string, Dictionary<string, HashSet<string>>> AllSymbols =
            new ConcurrentDictionary<string, Dictionary<string, HashSet<string>>>();

        private static string _symColName = "sym";

        private static string _funcAdd = ".u.add";
        private static string _funcDel = ".u.del";
        private static string _funcSub = ".u.sub";

        #region ExcelRtdServer

        //-------------------------------------------------------------------//        
        /// <summary>
        ///     Adds id to mapping with null value.
        /// </summary>
        /// <param name="alias">alias of the connection producing the data</param>
        /// <param name="table">name of the table</param>
        /// <param name="id">subscription id to add</param>
        private static void AddEmptyId(string alias, string table, string id)
        {
            if (!IdSymMap.ContainsKey(alias))
            {
                IdSymMap[alias] = new Dictionary<string, SortedDictionary<string, string>>();
                SymIdMap[alias] = new Dictionary<string, Dictionary<string, string>>();
                AllSymbols[alias] = new Dictionary<string, HashSet<string>>();
            }

            if (!IdSymMap[alias].ContainsKey(table))
            {
                IdSymMap[alias][table] = new SortedDictionary<string, string>(new StringIntComparer());
                SymIdMap[alias][table] = new Dictionary<string, string>();
                AllSymbols[alias][table] = new HashSet<string>();
            }


            if (!IdSymMap[alias][table].ContainsKey(id))
            {
                IdSymMap[alias][table].Add(id, null);
            }
        }

        /*
        private static void RemoveSymbolBasedOnName(string alias, string table, string symName)
        {
            if (!SymIdMap[alias][table].ContainsKey(symName)) return;
            var id = SymIdMap[alias][table][symName];
            IdSymMap[alias][table].Remove(id);
            SymIdMap[alias][table].Remove(symName);
        }
        */

        private static void RemoveSymbolBasedOnId(string alias, string table, string symId)
        {
            if (!IdSymMap[alias][table].ContainsKey(symId)) return;
            var symName = IdSymMap[alias][table][symId];
            IdSymMap[alias][table].Remove(symId);
            if (symName != null && SymIdMap[alias][table].ContainsKey(symName))
                SymIdMap[alias][table].Remove(symName);
        }

        //-------------------------------------------------------------------//        
        /// <summary>
        ///     Gets the subscription id based on symbol. If symbol is not assigned to anything it tries to assign it and returns
        ///     newly assigned id.
        /// </summary>
        /// <param name="alias">alias of the connection producing the data</param>
        /// <param name="table">name of the table</param>
        /// <param name="symName">symbol name to look for/assign</param>
        private static string GetSymbolId(string alias, string table, string symName)
        {
            if (!SymIdMap.ContainsKey(alias))
                return null;

            if (!SymIdMap[alias].ContainsKey(table))
                return null;

            if (SymIdMap[alias][table].ContainsKey(symName))
            {
                return SymIdMap[alias][table][symName];
            }
            foreach (var keyId in IdSymMap[alias][table].Keys.Where(keyId => IdSymMap[alias][table][keyId] == null))
            {
                IdSymMap[alias][table][keyId] = symName;
                SymIdMap[alias][table].Add(symName, keyId);
                return keyId;
            }

            return null;
        }

        //-------------------------------------------------------------------//
        protected override void ServerTerminate()
        {
            var keys = Connections.Keys.ToArray();
            foreach (var key in keys)
            {
                RtdClose(key);
            }
            Mapping.Clear(); // = null;
            WildCardMapping.Clear();
            Connections.Clear();
            IdSymMap.Clear();
            SymIdMap.Clear();
        }

        //-------------------------------------------------------------------//        
        /// <summary>
        ///     Tries to fill cells with data from cache.
        /// </summary>
        /// <param name="alias">alias of the connection producing the data</param>
        /// <param name="tab">name of the table</param>
        /// <param name="subscriptionId">subscriptionId mapping</param>
        /// <param name="col">column</param>
        private static void TryToFillFromCache(string alias, string tab, int subscriptionId, string col)
        {
            //check if this subscription id is aleady mapped to any symbol
            if (IdSymMap[alias][tab].ContainsKey(subscriptionId.ToString(CultureInfo.InvariantCulture)) &&
                IdSymMap[alias][tab][subscriptionId.ToString(CultureInfo.InvariantCulture)] != null)
            {
                //if yes, when update it from cache
                var dataToDisplay = Cache.GetData(alias, tab,
                    IdSymMap[alias][tab][subscriptionId.ToString(CultureInfo.InvariantCulture)], col, null);

                if (dataToDisplay == null) return;
                foreach (
                    var ti in
                        WildCardMapping[alias].GetTopics(tab, subscriptionId.ToString(CultureInfo.InvariantCulture),
                            col))
                {
                    DataOut.AddOrUpdate(ti, "", (k, v) => "");
                    ti.Topic.UpdateValue(dataToDisplay);
                }
            }
            else //no subscription id with symbol
            {
                //get potential symbols
                ICollection<String> symbolsBeingDisplayed = SymIdMap[alias][tab].Keys;
                var notDisplayed = new HashSet<string>(AllSymbols[alias][tab]);
                foreach (var symbol in symbolsBeingDisplayed)
                {
                    notDisplayed.Remove(symbol);
                }

                object dataToDisplay = null;
                object symbolToAdd = null;
                //check which symbol  is in cache
                foreach (var symbol in notDisplayed)
                {
                    dataToDisplay = Cache.GetData(alias, tab, symbol, col, null);
                    if (dataToDisplay != null)
                    {
                        symbolToAdd = symbol;
                        break;
                    }
                }

                //if anything is in cache, use it.
                if (dataToDisplay != null)
                {
                    IdSymMap[alias][tab][subscriptionId.ToString(CultureInfo.InvariantCulture)] = symbolToAdd.ToString();
                    SymIdMap[alias][tab].Add(symbolToAdd.ToString(),
                        subscriptionId.ToString(CultureInfo.InvariantCulture));

                    foreach (
                        var ti in
                            WildCardMapping[alias].GetTopics(tab, subscriptionId.ToString(CultureInfo.InvariantCulture),
                                col))
                    {
                        DataOut.AddOrUpdate(ti, "", (k, v) => "");
                        ti.Topic.UpdateValue(dataToDisplay);
                    }
                }
            }
        }

        //--------------------------called by excel just after entering formula RTD(xxxxx)-----------------------------------------//
        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            var alias = topicInfo[AliasPos];
            var tab = topicInfo[TabPos];
            var sym = topicInfo[SymPos];
            var col = topicInfo[ColPos];
            var his = topicInfo.Count > 4 ? topicInfo[HisPos] : null;
            var subscriptionId = 0;
            if (topicInfo.Count != 6 && sym.Equals(BackTick))
            {
                return "Missing required parameters if using with back tick.";
            }
            if (topicInfo.Count == 6)
            {
                try
                {
                    subscriptionId = Convert.ToInt32(topicInfo[SymbolIndetifier]);
                }
                catch (FormatException)
                {
                    return "Input string is not a sequence of digits (last parameter).";
                }
            }

            if (string.IsNullOrEmpty(tab))
            {
                return "Table is missing.";
            }
            if (string.IsNullOrEmpty(sym))
            {
                return "Symbol is missing.";
            }
            if (string.IsNullOrEmpty(col))
            {
                return "Column is missing.";
            }

            if (sym.Equals(BackTick))
            {
                if (Connections.ContainsKey(alias))
                {
                    if (!WildCardMapping.ContainsKey(alias) ||
                        (WildCardMapping.ContainsKey(alias) && !WildCardMapping[alias].GetTables().Contains(tab)))
                    {
                        RtdSubscribeTable(alias, tab);
                    }
                }

                //add id to symbol's map
                AddEmptyId(alias, tab, subscriptionId.ToString(CultureInfo.InvariantCulture));
                if (!WildCardMapping.ContainsKey(alias))
                {
                    var tm = new TopicMap();
                    WildCardMapping[alias] = tm;
                }

                WildCardMapping[alias].AddTopic(new TopicInfo(topic, alias, tab,
                    subscriptionId.ToString(CultureInfo.InvariantCulture), col, his));

                TryToFillFromCache(alias, tab, subscriptionId, col);
            }
            else
            {
                if (!Mapping.ContainsKey(alias) || !Mapping[alias].ContainsSymbol(tab, sym))
                {
                    if (Connections.ContainsKey(alias) &&
                        (!WildCardMapping.ContainsKey(alias) ||
                         (WildCardMapping.ContainsKey(alias) && !WildCardMapping[alias].GetTables().Contains(tab))))
                    {
                        RtdSubscribe(alias, tab, sym);
                    }
                }

                if (Mapping.ContainsKey(alias))
                {
                    Mapping[alias].AddTopic(new TopicInfo(topic, alias, tab, sym, col, his));
                }
                else
                {
                    var tm = new TopicMap();
                    tm.AddTopic(new TopicInfo(topic, alias, tab, sym, col, his));
                    Mapping[alias] = tm;
                }
            }
            newValues = true;
            return ExcelEmpty.Value;
        }


        //-------------------------------------------------------------------//
        protected override void DisconnectData(Topic topic)
        {
            //we get only the topicID so we need to check on which connection it is defined.
            foreach (var kv in Mapping)
            {
                if (!kv.Value.ContainsTopic(topic)) continue;
                //tab,sym,col
                var t = kv.Value.GetMapKeys(topic);
                kv.Value.RemoveTopic(topic);
                if (kv.Value.GetTopicCount(t.Item1, t.Item2) == 0)
                {
                    //all columns for given table and sym have been removed -> unsubscribe it totally
                    RtdUnsubscribe(kv.Key);
                }
                break;
            }

            foreach (var wildCardKv in WildCardMapping)
            {
                if (!wildCardKv.Value.ContainsTopic(topic)) continue;
                //tab,sym,col
                var t = wildCardKv.Value.GetMapKeys(topic);
                wildCardKv.Value.RemoveTopic(topic);
                //alias wildCardKV.key 
                var table = t.Item1;
                var sym = t.Item2;
                if (!wildCardKv.Value.ContainsSymbol(table, sym))
                {
                    RemoveSymbolBasedOnId(wildCardKv.Key, table, sym);
                }
                break;
            }
        }

        #endregion

        #region AddIn RTD

        //-------------------------------------------------------------------//        
        /// <summary>
        ///     Event triggered by the q process sending data
        /// </summary>
        /// <param name="sender">QConnection object delivering the data</param>
        /// <param name="args">data</param>
        private static void OnUpdate(object sender, QMessageEvent args)
        {
            var alias = GetAliasForConnection(sender as QConnection);
            {
                if (args == null || args.Message.Data == null) return;
                if (args.Message.Data is Array)
                {
                    var a = args.Message.Data as Array;

                    if (a.Length != 3) return;
                    if (!(a.GetValue(2) is QTable)) return;
                    try
                    {
                        var tab = a.GetValue(1).ToString();
                        var data = a.GetValue(2) as QTable;
                        lock (LockThis)
                        {
                            //this lock is mandatory here, since we need to preserve order filling topics.
                            UpdateCache(alias, tab, data);
                        }
                    }
// ReSharper disable EmptyGeneralCatchClause
                    catch (Exception)
// ReSharper restore EmptyGeneralCatchClause
                    {
                        // Console.WriteLine(e.Message);
                    }
                }
                else
                {
                    if (args.Message.Data.GetType().ToString() == "qSharp.QConnectionException" &&
                        ((QConnectionException) args.Message.Data).Message.Equals("Cannot read data from stream"))
                    {
                        //connections[alias].connectionCell.Formula = CloseConnectionRTD(alias);
                        RtdClose(alias);
                    }
                }
            }
        }

        //-------------------------------------------------------------------//        
        /// <summary>
        ///     This function applies data received from Q process to in-memory cache.
        ///     All operations are structured in a way that only data that has been subscribed is kept,
        ///     i.e.: data from columns that were not requested is simply discarded.
        /// </summary>
        /// <param name="alias">alias of the connection producing the data</param>
        /// <param name="table">name of the table</param>
        /// <param name="data">table data</param>
        private static void UpdateCache(string alias, string table, QTable data)
        {
            var symIdx = data.GetColumnIndex(_symColName);
            var cols = data.Columns;
            foreach (QTable.Row row in data)
            {
                var ra = row.ToArray();
                var symName = ra[symIdx].ToString();

                var symId = GetSymbolId(alias, table, symName);
                if (WildCardMapping.ContainsKey(alias) /*&& symId != null*/)
                {
                    for (var i = 0; i < cols.Length; i++)
                    {
                        Cache.UpdateData(alias, table, symName, cols[i], Conversions.Convert2Excel(ra[i]));

                        AllSymbols[alias][table].Add(symName);
                        if (symId == null) continue;
                        if (!WildCardMapping[alias].ContainsColumn(table, symId, cols[i])) continue;

                        foreach (var ti in WildCardMapping[alias].GetTopics(table, symId, cols[i]))
                        {
                            DataOut.AddOrUpdate(ti, "", (k, v) => "");
                            var val = string.IsNullOrEmpty(ti.History)
                                ? Conversions.Convert2Excel(ra[i])
                                : Cache.GetData(alias, table, symName, cols[i], ti.History);
                            ti.Topic.UpdateValue(val ?? ExcelEmpty.Value);
                        }
                    }
                }

                if (!Mapping.ContainsKey(alias) || !Mapping[alias].ContainsSymbol(table, symName))
                {
                    continue;
                }
                for (var i = 0; i < cols.Length; i++)
                {
                    if (!Mapping[alias].ContainsColumn(table, symName, cols[i])) continue;
                    Cache.UpdateData(alias, table, symName, cols[i], Conversions.Convert2Excel(ra[i]));
                    foreach (var ti in Mapping[alias].GetTopics(table, symName, cols[i]))
                    {
                        DataOut.AddOrUpdate(ti, "", (k, v) => "");
                        var val = string.IsNullOrEmpty(ti.History)
                            ? Conversions.Convert2Excel(ra[i])
                            : Cache.GetData(alias, table, symName, cols[i], ti.History);
                        ti.Topic.UpdateValue(val ?? ExcelEmpty.Value);
                    }
                }
            }
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Re-maps connection object to its alias.
        /// </summary>
        /// <param name="qConnection">QConnection object</param>
        /// <returns>connection alias bound with given connection or null in case alias cannot be found</returns>
        private static string GetAliasForConnection(QConnection qConnection)
        {
            return null != Connections
                ? (from kv in Connections where kv.Value.Equals(qConnection) select kv.Key)
                    .FirstOrDefault()
                : null;
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Extends current subscription bound with connection alias by given sym for
        ///     given tab.
        /// </summary>
        /// <param name="alias">connection alias</param>
        /// <param name="tab">table</param>
        /// <param name="sym">symbol</param>
// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable once UnusedMethodReturnValue.Global
// ReSharper disable UnusedMethodReturnValue.Local
        private static object RtdSubscribe(string alias,
// ReSharper restore UnusedMethodReturnValue.Local
// ReSharper restore MemberCanBePrivate.Global
            string tab,
            string sym)
        {
            var c = GetConnection(alias);
            if (c == null) return ExcelError.ExcelErrorNull;
            var syms = new[] {sym};
            c.Async(_funcAdd, new object[] {tab, syms});

            return ExcelEmpty.Value;
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Removes given symbol for given table from subscription on the connection
        ///     bound with given alias.
        /// </summary>
        /// <param name="alias">connection alias</param>
// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable once UnusedMethodReturnValue.Global
// ReSharper disable UnusedMethodReturnValue.Local
        private static object RtdUnsubscribe(string alias)
// ReSharper restore UnusedMethodReturnValue.Local
// ReSharper restore MemberCanBePrivate.Global
        {
            var c = GetConnection(alias);
            if (c == null) return ExcelError.ExcelErrorNull;
            return ExcelEmpty.Value;
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Subscribe entire content of the mapping for given alias.
        /// </summary>
        /// <param name="alias">connection alias</param>
// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable once UnusedMethodReturnValue.Global
// ReSharper disable UnusedMethodReturnValue.Local
        private static object RtdSubscribeAllTables(string alias)
// ReSharper restore UnusedMethodReturnValue.Local
// ReSharper restore MemberCanBePrivate.Global
        {
            var c = GetConnection(alias);
            if (c == null)
            {
                return ExcelError.ExcelErrorNull;
            }
            if (!WildCardMapping.ContainsKey(alias))
            {
                return ExcelError.ExcelErrorNA;
            }

            var tables = WildCardMapping[alias].GetTables();
            foreach (var t in tables)
            {
                c.Async(_funcSub, new object[] {t, ""});
            }

            return ExcelEmpty.Value;
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Subscribe entire content of the mapping for given alias.
        /// </summary>
        /// <param name="alias">connection alias</param>
// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable once UnusedMethodReturnValue.Global
// ReSharper disable UnusedMethodReturnValue.Local
        private static object RtdSubscribeAll(string alias)
// ReSharper restore UnusedMethodReturnValue.Local
// ReSharper restore MemberCanBePrivate.Global
        {
            var c = GetConnection(alias);
            if (c == null)
            {
                return ExcelError.ExcelErrorNull;
            }
            if (!Mapping.ContainsKey(alias))
            {
                return ExcelError.ExcelErrorNA;
            }

            var tables = Mapping[alias].GetTables();
            foreach (var t in tables)
            {
                var syms = Mapping[alias].GetSymbols(t);
                c.Async(_funcSub, new object[] {t, syms});
            }

            return ExcelEmpty.Value;
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Subscribe for all symbols from given table
        /// </summary>
        /// <param name="alias">connection alias</param>
        /// <param name="tableName">table name</param>
        // ReSharper disable MemberCanBePrivate.Global
// ReSharper disable once UnusedMethodReturnValue.Global
// ReSharper disable UnusedMethodReturnValue.Local
        private static object RtdSubscribeTable(string alias, string tableName)
// ReSharper restore UnusedMethodReturnValue.Local
            // ReSharper restore MemberCanBePrivate.Global
        {
            var c = GetConnection(alias);
            if (c == null)
            {
                return ExcelError.ExcelErrorNull;
            }

            c.Async(_funcSub, new object[] {tableName, ""});

            return ExcelEmpty.Value;
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Opens connection to q process serving as RTD Server
        /// </summary>
        /// <param name="alias">connection alias</param>
        /// <param name="hostname">hostname</param>
        /// <param name="port">port</param>
        /// <param name="username">user</param>
        /// <param name="password">password</param>
        /// <returns>alias in case of successfull connection, error message otherwise</returns>
        [ExcelFunction(Description = "Opens connection to q process serving as RTD Server.",
            Category = "qXL (RTD)", Name = "qRtdOpen")]
// ReSharper disable UnusedMember.Global
        public static object RtdOpen([ExcelArgument("Logical identifier for the connection.")] object alias,
// ReSharper restore UnusedMember.Global
            [ExcelArgument(
                "Hostname, fqdn or ip of the machine running q process to which connection should be established."
                )] object hostname,
            [ExcelArgument("Port number of the q process.")] object port,
            [ExcelArgument("Username")] object username = null,
            [ExcelArgument("Password")] object password = null)
        {
            try
            {
                if (alias == null || !alias.ToString().Trim().Any())
                {
                    return "Invalid alias";
                }

                if (hostname == null || !hostname.ToString().Trim().Any())
                {
                    return "Invalid hostname";
                }

                try
                {
                    port = Int32.Parse(port.ToString());
                }
                catch
                {
                    return "Invalid port";
                }

                var c = GetConnection(alias.ToString());
                if (c != null)
                {
                    return alias;
                }

                var u = username != null && username.GetType() == ExcelMissing.Value.GetType()
                    ? null
                    : username as string;
                var p = password != null && password.GetType() == ExcelMissing.Value.GetType()
                    ? null
                    : password as string;
                try
                {
                    c = new QCallbackConnection(hostname.ToString(), (int) port, u, p);
                    c.Open();
                    c.DataReceived += OnUpdate; //assign handler function.
                    c.StartListener();

                    Connections[alias.ToString()] = c;

                    if (WildCardMapping.ContainsKey(alias.ToString()))
                    {
                        RtdSubscribeAllTables(alias.ToString());
                    }
                    if (Mapping.ContainsKey(alias.ToString()) && !WildCardMapping.ContainsKey(alias.ToString()))
                    {
                        RtdSubscribeAll(alias.ToString());
                    }
                }
                catch (QException e)
                {
                    return e.Message;
                }
            }
            catch (Exception e)
            {
                return e.Message;
            }
            return alias;
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Closes the underlying connection to q process.
        /// </summary>
        /// <param name="alias">alias bound with q connection</param>
        /// <returns>"Closed" string in case of successfull close, error message otherwise</returns>
        [ExcelFunction(Description = "Closes the connection associated with given alias.",
            Category = "qXL (RTD)", Name = "qRtdClose")]
// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable once UnusedMethodReturnValue.Global
        public static object RtdClose([ExcelArgument("Logical identifier for the connection.")] object alias)
// ReSharper restore MemberCanBePrivate.Global
        {
            if (alias == null)
            {
                return "Alias is null.";
            }

            try
            {
                var a = alias.ToString();
                QConnection c = null;
                if (Connections.ContainsKey(a))
                {
                    c = Connections[a];
                }
                if (Mapping.ContainsKey(a)) //remove topics connected with this connection
                {
                    TopicMap val;
                    Mapping.TryRemove(a, out val);
                }

                if (WildCardMapping.ContainsKey(a)) //remove topics connected with this connection
                {
                    TopicMap val;
                    WildCardMapping.TryRemove(a, out val);
                    if (val != null)
                    {
                        IdSymMap[a].Clear();
                        SymIdMap[a].Clear();
                    }
                }

                if (c != null)
                {
                    Connections[a].StopListener();
                    Connections[a].Close();
                    QCallbackConnection con;
                    Connections.TryRemove(a, out con);
                    return String.Format("Disconnected from '{0}'", a);
                }
            }
            catch (Exception e)
            {
                return e.Message;
            }
            return String.Format("Unknown alias '{0}'", alias);
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Allows setting of various parameters for RTD server.
        ///     Valid parameter names are:
        ///     sym.column.name - [string] column name in which instrument identifier is kept (default: sym)
        ///     function.add - [string] name of function that should be used for adding symbols for subscription (default: .u.add)
        ///     function.sub - [string] name of function that should be used for subscribing symbols (default: .u.sub)
        ///     function.del - [string] name of function that should be used for removing symbols from subscription (default:
        ///     .u.del)
        ///     data.history.length - [int] amount of historical updates kept in memory for reference
        /// </summary>
        /// <param name="paramName">parameter name</param>
        /// <param name="paramValue">parameter value</param>
        /// <returns>value of the parameter or error message</returns>
        [ExcelFunction(Description = "Allows setting of various parameters for RTD server.",
            Category = "qXL (RTD)", Name = "qRtdConfigure")]
// ReSharper disable UnusedMember.Global
        public static object RtdConfigure([ExcelArgument("Parameter name.")] string paramName,
// ReSharper restore UnusedMember.Global
            [ExcelArgument("Parameter value.")] object paramValue)
        {
            try
            {
                switch (paramName)
                {
                    case "sym.column.name":
                        _symColName = paramValue.ToString();
                        return _symColName;
                    case "function.add":
                        _funcAdd = paramValue.ToString();
                        return _funcAdd;
                    case "function.sub":
                        _funcSub = paramValue.ToString();
                        return _funcSub;
                    case "function.del":
                        _funcDel = paramValue.ToString();
                        return _funcDel;
                    case "data.history.length":
                        int h;
                        Int32.TryParse(paramValue.ToString(), out h);
                        Cache.ChangeHistoryLength(h);
                        return Cache.GetHistoryLength();
                    default:
                        return "Unknown param: " + paramName;
                }
            }
            catch (Exception)
            {
                return "Invalid operation";
            }
        }

        #endregion

        #region Helper

        private static QCallbackConnection GetConnection(string alias)
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
            c.Open();
            c.DataReceived += OnUpdate; //assign handler function.
            c.StartListener();
            return c;
        }

        #endregion
    }
}