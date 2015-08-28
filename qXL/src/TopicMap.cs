// 
// Copyright (c) 2011-2015 Exxeleron GmbH
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
using System.Linq;
using ExcelDna.Integration.Rtd;

#endregion

namespace qXL
{
    internal class TopicMap
    {
        //reversed map to speed up topic removal
        private readonly Dictionary<ExcelRtdServer.Topic, Tuple<string, string, string>> _revMap =
            new Dictionary<ExcelRtdServer.Topic, Tuple<string, string, string>>();

        private readonly Dictionary<string, Dictionary<string, Dictionary<string, List<TopicInfo>>>> _topicMap =
            new Dictionary<string, Dictionary<string, Dictionary<string, List<TopicInfo>>>>();

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Allows adding of the TopicInfo structure to the map.
        /// </summary>
        /// <param name="ti">object to be stored in map</param>
        public void AddTopic(TopicInfo ti)
        {
            if (string.IsNullOrEmpty(ti.Table) || string.IsNullOrEmpty(ti.Symbol) || string.IsNullOrEmpty(ti.Column))
            {
                return;
            }

            if (!_revMap.ContainsKey(ti.Topic))
                _revMap[ti.Topic] = new Tuple<string, string, string>(ti.Table, ti.Symbol, ti.Column);

            if (!_topicMap.ContainsKey(ti.Table))
                _topicMap[ti.Table] = new Dictionary<string, Dictionary<string, List<TopicInfo>>>();

            if (!_topicMap[ti.Table].ContainsKey(ti.Symbol))
                _topicMap[ti.Table][ti.Symbol] = new Dictionary<string, List<TopicInfo>>();

            if (!_topicMap[ti.Table][ti.Symbol].ContainsKey(ti.Column))
            {
                var li = new List<TopicInfo> {ti};
                _topicMap[ti.Table][ti.Symbol].Add(ti.Column, li);
            }
            else
            {
                _topicMap[ti.Table][ti.Symbol][ti.Column].Add(ti);
            }
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Provides mapping between tab-sym-col triple and list of Excel topic ids for RTD formula. Topics connected
        ///     with historical values for given symbol are kept together with the Topic for the symbol itself.
        /// </summary>
        /// <param name="tab">table</param>
        /// <param name="sym">symbol</param>
        /// <param name="col">column</param>
        /// <returns>List of TopicInfo objects connected with symbol or its historical values</returns>
        public IEnumerable<TopicInfo> GetTopics(string tab, string sym, string col)
        {
            if (!_topicMap.ContainsKey(tab))
                return null;
            if (!_topicMap[tab].ContainsKey(sym))
                return null;
            return !_topicMap[tab][sym].ContainsKey(col) ? Enumerable.Empty<TopicInfo>() : _topicMap[tab][sym][col];
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Removes given topic from internal mpapings
        /// </summary>
        /// <param name="topic">topic id to be removed</param>
        public void RemoveTopic(ExcelRtdServer.Topic topic)
        {
            if (!_revMap.ContainsKey(topic)) return;
            var t = _revMap[topic]; //get table,sym,col                

            var list = _topicMap[t.Item1][t.Item2][t.Item3];
            foreach (var ti in list.Where(ti => ti.Topic == topic))
            {
                list.Remove(ti);
                break;
            }
            _revMap.Remove(topic);
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Verifies whether given topic exists in internal mappings
        /// </summary>
        /// <param name="topic">topic id</param>
        /// <returns>ture in case topic is present, false otherwise</returns>
        public bool ContainsTopic(ExcelRtdServer.Topic topic)
        {
            return _revMap.ContainsKey(topic);
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Verifies whether symbol for given table is already present in internal mappings.
        /// </summary>
        /// <param name="tab">table</param>
        /// <param name="sym">symbol</param>
        /// <returns>true if symbol exists in internal mappings,false otherwise</returns>
        public bool ContainsSymbol(string tab, string sym)
        {
            if (!(_topicMap.ContainsKey(tab) && _topicMap[tab].ContainsKey(sym)))
            {
                return false;
            }

            return _topicMap[tab][sym].Any(col => col.Value.Count > 0);
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Verifies whether given column for given symbol in given table is mapped to some topic already
        /// </summary>
        /// <param name="tab">table</param>
        /// <param name="sym">symbol</param>
        /// <param name="col">column</param>
        /// <returns>true in case column is mapped to some topic already, false otherwise</returns>
        public bool ContainsColumn(string tab, string sym, string col)
        {
            return _topicMap.ContainsKey(tab) && _topicMap[tab].ContainsKey(sym) && _topicMap[tab][sym].ContainsKey(col) &&
                   _topicMap[tab][sym][col].Count > 0;
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Provides means of getting all symbols for given table that are mapped
        ///     to some topic already.
        /// </summary>
        /// <param name="tab">table</param>
        /// <returns>list of symbols</returns>
        public string[] GetSymbols(string tab)
        {
            return _topicMap.ContainsKey(tab) ? _topicMap[tab].Keys.ToArray() : null;
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Lists all the tables that contain mapping for Excel topic ids for RTD.
        /// </summary>
        /// <returns>list of tables in mapping</returns>
        public IEnumerable<string> GetTables()
        {
            return _topicMap.Keys.ToArray();
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Provides mapping between given topic id from Excel RTD and table,symbol,column
        ///     from kdb+
        /// </summary>
        /// <param name="topic">RTD formula topic id</param>
        /// <returns>Triplet: table,symbol,column  or null in case topic is not mapped</returns>
        public Tuple<string, string, string> GetMapKeys(ExcelRtdServer.Topic topic)
        {
            return _revMap.ContainsKey(topic) ? _revMap[topic] : null;
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Counts topics connected with given table and symbol
        /// </summary>
        /// <param name="tab">table</param>
        /// <param name="sym">symbol</param>
        /// <returns>count of topics bound with given table and symbol(including history)</returns>
        public int GetTopicCount(string tab, string sym)
        {
            if (!_topicMap.ContainsKey(tab))
                return 0;
            return !_topicMap[tab].ContainsKey(sym) ? 0 : _topicMap[tab][sym].Sum(kv => kv.Value.Count);
        }
    }
}