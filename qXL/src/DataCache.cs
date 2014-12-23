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
using System.Collections.Generic;
using System.Linq;

#endregion

namespace qXL
{
    internal class DataCache
    {
        private const string Ad = "?";

        private readonly
            ConcurrentDictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, object>>>> _data =
                new ConcurrentDictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, object>>>>();

        private int _hLen;
        //artificial delimiter


        /// <summary>
        ///     Simple constructor allowing initialization of history length.
        /// </summary>
        /// <param name="historyLength">amout of history positions that should be kept per symbol</param>
        public DataCache(int historyLength)
        {
            _hLen = -1*historyLength;
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Function for updating data in memory. Automatically rewinds data for historical values.
        /// </summary>
        /// <param name="alias">connection alias</param>
        /// <param name="table">table name</param>
        /// <param name="symbol">symbol name</param>
        /// <param name="column">column name</param>
        /// <param name="data">data to be stored on most recent("current") position</param>
        public void UpdateData(string alias, string table, string symbol, string column, object data)
        {
            if (!_data.ContainsKey(alias))
            {
                _data[alias] = new Dictionary<string, Dictionary<string, Dictionary<string, object>>>();
            }

            if (!_data[alias].ContainsKey(table))
                _data[alias][table] = new Dictionary<string, Dictionary<string, object>>();

            if (!_data[alias][table].ContainsKey(symbol))
                _data[alias][table][symbol] = new Dictionary<string, object>();

            if (!_data[alias][table][symbol].ContainsKey(column))
            {
                _data[alias][table][symbol].Add(column, data);
            }
            else
            {
                RewindData(alias, table, symbol, column);
                _data[alias][table][symbol][column] = data;
            }
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Function for shifting data in history. In case we have a history set to 5 this function will:
        ///     1. re-write 4th value to 5th position
        ///     2. re-write 3rd value to 4th position
        ///     3. re-write 2nd value to 3rd position
        ///     4. re-write 1st value to 2nd position
        /// </summary>
        /// <param name="alias">connection alias</param>
        /// <param name="table">name of the table</param>
        /// <param name="symbol">symbol</param>
        /// <param name="column">column name</param>
        private void RewindData(string alias, string table, string symbol, string column)
        {
            for (var i = _hLen + 1; i <= 0; i++)
            {
                var currSym = i != 0 ? symbol + Ad + i : symbol;
                var prevSym = symbol + Ad + (i - 1);

                var currVal = GetData(alias, table, currSym, column);

                if (_data[alias][table].ContainsKey(prevSym))
                {
                    if (_data[alias][table][prevSym].ContainsKey(column))
                        _data[alias][table][prevSym][column] = currVal;
                    else
                        _data[alias][table][prevSym].Add(column, currVal);
                }
                else
                {
                    var tmp = new Dictionary<string, object> {{column, currVal}};
                    _data[alias][table].Add(prevSym, tmp);
                }
            }
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Simple data getter. Does not take history into account, so for history fetching overloaded version of this function
        ///     should be used.
        ///     In case data is missing will return null.
        /// </summary>
        /// <param name="alias">connection alias</param>
        /// <param name="table">table name</param>
        /// <param name="symbol">symbol</param>
        /// <param name="column">column</param>
        /// <returns>current value from data cache</returns>
        private object GetData(string alias, string table, string symbol, string column)
        {
            if (!_data.ContainsKey(alias))
                return null;
            if (!_data[alias].ContainsKey(table))
                return null;
            if (!_data[alias][table].ContainsKey(symbol))
                return null;
            return !_data[alias][table][symbol].ContainsKey(column) ? null : _data[alias][table][symbol][column];
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Simple data getter. Allows to retrieve data from given history level, e.g.: "-2"
        ///     In case data is missing will return null.
        /// </summary>
        /// <param name="alias">connection alias</param>
        /// <param name="table">table name </param>
        /// <param name="symbol">symbol </param>
        /// <param name="column">column</param>
        /// <param name="historyLvl">
        ///     "position" in history relative to "current" value, e.g.: "-1","-2" would mean previous and
        ///     previous previous value respectively
        /// </param>
        /// <returns>value at given history level or null in case not found</returns>
        public object GetData(string alias, string table, string symbol, string column, string historyLvl)
        {
            if (historyLvl != null && historyLvl != "0")
            {
                symbol = symbol + Ad + historyLvl;
            }
            return GetData(alias, table, symbol, column);
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Changes history size to be kept in memory. In case the history is made shorter data that is now found behind the
        ///     history window will be deleted.
        /// </summary>
        /// <param name="newLength">new length of the history vector</param>
        public void ChangeHistoryLength(int newLength)
        {
            var nl = -1*newLength;
            if (nl > _hLen) // in case length is made smaller, we need to clean "old trash" from memory
            {
                //one cannot modify collection while iterating over it, so elements to be removed will be stored in toRem variable and removed after iteration.
                var toRem = new List<Tuple<string, string, string>>();

                for (var i = _hLen; i < nl; i++)
                {
                    toRem.AddRange(from kv in _data
                        from kv1 in kv.Value
                        from kv2 in kv1.Value
                        where kv2.Key.Contains(Ad + i)
                        select new Tuple<string, string, string>(kv.Key, kv1.Key, kv2.Key));
                }

                foreach (var t in toRem)
                {
                    _data[t.Item1][t.Item2].Remove(t.Item3);
                }
            }
            _hLen = nl;
        }

        //-------------------------------------------------------------------//
        /// <summary>
        /// </summary>
        /// <returns> current amount of history positions</returns>
        public int GetHistoryLength()
        {
            return -1*_hLen;
        }
    }
}