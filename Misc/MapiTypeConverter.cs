/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;

    using MapiTypeConverterMap = System.Collections.Generic.Dictionary<MapiPropertyType, MapiTypeConverterMapEntry>;

    /// <summary>
    /// Utility class to convert between MAPI Property type values and strings.
    /// </summary>
    internal class MapiTypeConverter
    {
        /// <summary>
        /// Assume DateTime values are in UTC.
        /// </summary>
        private const DateTimeStyles UtcDataTimeStyles = DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal;

        /// <summary>
        /// Map from MAPI property type to converter entry.
        /// </summary>
        private static LazyMember<MapiTypeConverterMap> mapiTypeConverterMap = new LazyMember<MapiTypeConverterMap>(
            delegate()
            {
                MapiTypeConverterMap map = new MapiTypeConverterMap();

                map.Add(
                    MapiPropertyType.ApplicationTime,
                    new MapiTypeConverterMapEntry(typeof(double)));

                map.Add(
                    MapiPropertyType.ApplicationTimeArray,
                    new MapiTypeConverterMapEntry(typeof(double)) { IsArray = true });

                var byteConverter = new MapiTypeConverterMapEntry(typeof(byte[]))
                    {
                        Parse = (s) => string.IsNullOrEmpty(s) ? null : Convert.FromBase64String(s),
                        ConvertToString = (o) => Convert.ToBase64String((byte[])o),
                    };

                map.Add(
                    MapiPropertyType.Binary,
                    byteConverter);

                var byteArrayConverter = new MapiTypeConverterMapEntry(typeof(byte[]))
                    {
                        Parse = (s) => string.IsNullOrEmpty(s) ? null : Convert.FromBase64String(s),
                        ConvertToString = (o) => Convert.ToBase64String((byte[])o),
                        IsArray = true
                    };

                map.Add(
                    MapiPropertyType.BinaryArray,
                    byteArrayConverter);

                var boolConverter = new MapiTypeConverterMapEntry(typeof(bool))
                    {
                        Parse = (s) => Convert.ChangeType(s, typeof(bool), CultureInfo.InvariantCulture),
                        ConvertToString = (o) => ((bool)o).ToString(CultureInfo.InvariantCulture).ToLower(),
                    };

                map.Add(
                    MapiPropertyType.Boolean,
                    boolConverter);

                var clsidConverter = new MapiTypeConverterMapEntry(typeof(Guid))
                    {
                        Parse = (s) => new Guid(s),
                        ConvertToString = (o) => ((Guid)o).ToString(),
                    };

                map.Add(
                    MapiPropertyType.CLSID,
                    clsidConverter);

                var clsidArrayConverter = new MapiTypeConverterMapEntry(typeof(Guid))
                    {
                        Parse = (s) => new Guid(s),
                        ConvertToString = (o) => ((Guid)o).ToString(),
                        IsArray = true
                    };

                map.Add(
                    MapiPropertyType.CLSIDArray,
                    clsidArrayConverter);

                map.Add(
                    MapiPropertyType.Currency,
                    new MapiTypeConverterMapEntry(typeof(Int64)));

                map.Add(
                    MapiPropertyType.CurrencyArray,
                    new MapiTypeConverterMapEntry(typeof(Int64)) { IsArray = true });

                map.Add(
                    MapiPropertyType.Double,
                    new MapiTypeConverterMapEntry(typeof(double)));

                map.Add(
                    MapiPropertyType.DoubleArray,
                    new MapiTypeConverterMapEntry(typeof(double)) { IsArray = true });

                map.Add(
                    MapiPropertyType.Error,
                    new MapiTypeConverterMapEntry(typeof(Int32)));

                map.Add(
                    MapiPropertyType.Float,
                    new MapiTypeConverterMapEntry(typeof(float)));

                map.Add(
                    MapiPropertyType.FloatArray,
                    new MapiTypeConverterMapEntry(typeof(float)) { IsArray = true });

                map.Add(
                    MapiPropertyType.Integer,
                    new MapiTypeConverterMapEntry(typeof(Int32))
                    {
                        Parse = (s) => MapiTypeConverter.ParseMapiIntegerValue(s)
                    });

                map.Add(
                    MapiPropertyType.IntegerArray,
                    new MapiTypeConverterMapEntry(typeof(Int32)) { IsArray = true });

                map.Add(
                    MapiPropertyType.Long,
                    new MapiTypeConverterMapEntry(typeof(Int64)));

                map.Add(
                    MapiPropertyType.LongArray,
                    new MapiTypeConverterMapEntry(typeof(Int64)) { IsArray = true });

                var objectConverter = new MapiTypeConverterMapEntry(typeof(string))
                    {
                        Parse = (s) => s
                    };

                map.Add(
                    MapiPropertyType.Object,
                    objectConverter);

                var objectArrayConverter = new MapiTypeConverterMapEntry(typeof(string))
                    {
                        Parse = (s) => s,
                        IsArray = true
                    };

                map.Add(
                    MapiPropertyType.ObjectArray,
                    objectArrayConverter);

                map.Add(
                    MapiPropertyType.Short,
                    new MapiTypeConverterMapEntry(typeof(Int16)));

                map.Add(
                    MapiPropertyType.ShortArray,
                    new MapiTypeConverterMapEntry(typeof(Int16)) { IsArray = true });

                var stringConverter = new MapiTypeConverterMapEntry(typeof(string))
                    {
                        Parse = (s) => s
                    };

                map.Add(
                    MapiPropertyType.String,
                    stringConverter);

                var stringArrayConverter = new MapiTypeConverterMapEntry(typeof(string))
                    {
                        Parse = (s) => s,
                        IsArray = true
                    };

                map.Add(
                    MapiPropertyType.StringArray,
                    stringArrayConverter);

                var sysTimeConverter = new MapiTypeConverterMapEntry(typeof(DateTime))
                    {
                        Parse = (s) => DateTime.Parse(s, CultureInfo.InvariantCulture, UtcDataTimeStyles),
                        ConvertToString = (o) => EwsUtilities.DateTimeToXSDateTime((DateTime)o) // Can't use DataTime.ToString()
                    };

                map.Add(
                    MapiPropertyType.SystemTime,
                    sysTimeConverter);

                var sysTimeArrayConverter = new MapiTypeConverterMapEntry(typeof(DateTime))
                    {
                        IsArray = true,
                        Parse = (s) => DateTime.Parse(s, CultureInfo.InvariantCulture, UtcDataTimeStyles),
                        ConvertToString = (o) => EwsUtilities.DateTimeToXSDateTime((DateTime)o) // Can't use DataTime.ToString()
                    };

                map.Add(
                    MapiPropertyType.SystemTimeArray,
                    sysTimeArrayConverter);

                return map;
            });

        /// <summary>
        /// Converts the string list to array.
        /// </summary>
        /// <param name="mapiPropType">Type of the MAPI property.</param>
        /// <param name="strings">Strings.</param>
        /// <returns>Array of objects.</returns>
        internal static Array ConvertToValue(MapiPropertyType mapiPropType, IEnumerable<string> strings)
        {
            EwsUtilities.ValidateParam(strings, "strings");

            MapiTypeConverterMapEntry typeConverter = MapiTypeConverterMap[mapiPropType];
            Array array = Array.CreateInstance(typeConverter.Type, strings.Count<string>());

            int index = 0;
            foreach (string stringValue in strings)
            {
                object value = typeConverter.ConvertToValueOrDefault(stringValue);
                array.SetValue(value, index++);
            }

            return array;
        }

        /// <summary>
        /// Converts a string to value consistent with MAPI type.
        /// </summary>
        /// <param name="mapiPropType">Type of the MAPI property.</param>
        /// <param name="stringValue">String to convert to a value.</param>
        /// <returns></returns>
        internal static object ConvertToValue(MapiPropertyType mapiPropType, string stringValue)
        {
            return MapiTypeConverterMap[mapiPropType].ConvertToValue(stringValue);
        }

        /// <summary>
        /// Converts a value to a string.
        /// </summary>
        /// <param name="mapiPropType">Type of the MAPI property.</param>
        /// <param name="value">Value to convert to string.</param>
        /// <returns>String value.</returns>
        internal static string ConvertToString(MapiPropertyType mapiPropType, object value)
        {
            return (value == null)
                    ? string.Empty
                    : MapiTypeConverterMap[mapiPropType].ConvertToString(value);
        }

        /// <summary>
        /// Change value to a value of compatible type.
        /// </summary>
        /// <param name="mapiType">Type of the mapi property.</param>
        /// <param name="value">The value.</param>
        /// <returns>Compatible value.</returns>
        internal static object ChangeType(MapiPropertyType mapiType, object value)
        {
            EwsUtilities.ValidateParam(value, "value");

            return MapiTypeConverterMap[mapiType].ChangeType(value);
        }

        /// <summary>
        /// Converts a MAPI Integer value.
        /// </summary>
        /// <remarks>
        /// Usually the value is an integer but there are cases where the value has been "schematized" to an 
        /// Enumeration value (e.g. NoData) which we have no choice but to fallback and represent as a string.
        /// </remarks>
        /// <param name="s">The string value.</param>
        /// <returns>Integer value or the original string if the value could not be parsed as such.</returns>
        internal static object ParseMapiIntegerValue(string s)
        {
            int intValue;
            if (Int32.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out intValue))
            {
                return intValue;
            }
            else
            {
                return s;
            }
        }

        /// <summary>
        /// Determines whether MapiPropertyType is an array type.
        /// </summary>
        /// <param name="mapiType">Type of the mapi.</param>
        /// <returns>True if this is an array type.</returns>
        internal static bool IsArrayType(MapiPropertyType mapiType)
        {
            return MapiTypeConverterMap[mapiType].IsArray;
        }

        /// <summary>
        /// Gets the MAPI type converter map.
        /// </summary>
        /// <value>The MAPI type converter map.</value>
        internal static MapiTypeConverterMap MapiTypeConverterMap
        {
            get { return mapiTypeConverterMap.Member; }
        }
    }
}