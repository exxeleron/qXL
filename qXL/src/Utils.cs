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
using System.Linq;
using qSharp;

#endregion

namespace qXL
{
    internal static class Utils
    {
        //-------------------------------------------------------------------//
        /// <summary>
        ///     Searches for the longest array in a dictionary.
        /// </summary>
        /// <param name="dict">dictionary</param>
        /// <returns>length of longest array within the dictionary</returns>
        public static int GetMaxDictSize(QDictionary dict)
        {
            var maxLen = -1;
            foreach (var array in from QDictionary.KeyValuePair kv in dict select kv.Value as Array)
            {
                if (array != null)
                {
                    var type = array.GetType().Name.ToLower();
                    if (!type.Equals("char[]"))
                    {
                        if (array.GetValue(array.GetLowerBound(0)) is Array)
                            throw new ConversionException(
                                "Cannot handle nested multidimensional arrays as a dictionary value");

                        maxLen = (maxLen < array.Length) ? (array.Length) : maxLen;
                    }
                    else
                    {
                        maxLen = (maxLen < 1) ? 1 : maxLen;
                    }
                }
                else
                {
                    maxLen = (maxLen < 1) ? 1 : maxLen;
                }
            }
            return maxLen;
        }

        //-------------------------------------------------------------------//
        /// <summary>
        ///     Converts COM like 2dim array where one dimension is of length 1 to regular array.
        /// </summary>
        /// <param name="a">COM array</param>
        /// <returns>regular array</returns>
        public static object[] Com2DArray2Array(Array a)
        {
            if (a == null)
                return null;

            object[] converted = null;
            switch (a.Rank)
            {
                case 1:
                    converted = new object[a.GetLength(0)];
                    for (var i = a.GetLowerBound(0); i <= a.GetUpperBound(0); i++)
                    {
                        converted[i] = a.GetValue(i);
                    }
                    break;
                case 2:
                {
                    var d1 = a.GetLength(0);
                    var d2 = a.GetLength(1);
                    var len = (d1 > d2) ? d1 : d2;
                    converted = new object[len];
                    var dim = (d1 > d2) ? 0 : 1;
                    for (var i = a.GetLowerBound(dim); i <= a.GetUpperBound(dim); i++)
                    {
                        converted[i - a.GetLowerBound(dim)] = a.GetValue((d1 == 1 ? a.GetLowerBound(0) : i),
                            (d2 == 1 ? a.GetLowerBound(1) : i));
                    }
                }
                    break;
            }

            return converted;
        }
    }
}