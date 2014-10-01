#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the MapiTypeConverterMapEntry class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Globalization;

    using TypeToDefaultValueMap = System.Collections.Generic.Dictionary<System.Type, object>;

    /// <summary>
    /// Represents an entry in the MapiTypeConverter map.
    /// </summary>
    internal class MapiTypeConverterMapEntry
    {
        /// <summary>
        /// Map CLR types used for MAPI properties to matching default values.
        /// </summary>
        private static LazyMember<TypeToDefaultValueMap> defaultValueMap = new LazyMember<TypeToDefaultValueMap>(
            () =>
            {
                TypeToDefaultValueMap map = new TypeToDefaultValueMap();

                map.Add(typeof(bool), false);
                map.Add(typeof(byte[]), null);
                map.Add(typeof(Int16), (Int16)0);
                map.Add(typeof(Int32), (Int32)0);
                map.Add(typeof(Int64), (Int64)0);
                map.Add(typeof(float), (float)0.0);
                map.Add(typeof(double), (double)0.0);
                map.Add(typeof(DateTime), DateTime.MinValue);
                map.Add(typeof(Guid), Guid.Empty);
                map.Add(typeof(string), null);

                return map;
            });

        /// <summary>
        /// Initializes a new instance of the <see cref="MapiTypeConverterMapEntry"/> class.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <remarks>
        /// By default, converting a type to string is done by calling value.ToString. Instances
        /// can override this behavior.
        /// By default, converting a string to the appropriate value type is done by calling Convert.ChangeType
        /// Instances may override this behavior.
        /// </remarks>
        internal MapiTypeConverterMapEntry(Type type)
        {
            EwsUtilities.Assert(
                defaultValueMap.Member.ContainsKey(type),
                "MapiTypeConverterMapEntry ctor",
                string.Format("No default value entry for type {0}", type.Name));

            this.Type = type;
            this.ConvertToString = (o) => (string)Convert.ChangeType(o, typeof(string), CultureInfo.InvariantCulture);
            this.Parse = (s) => Convert.ChangeType(s, type, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Change value to a value of compatible type.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>New value.</returns>
        /// <remarks>
        /// The type of a simple value should match exactly or be convertible to the appropriate type. An 
        /// array value has to be a single dimension (rank), contain at least one value and contain 
        /// elements that exactly match the expected type. (We could relax this last requirement so that,
        /// for example, you could pass an array of Int32 that could be converted to an array of Double
        /// but that seems like overkill).
        /// </remarks>
        internal object ChangeType(object value)
        {
            if (this.IsArray)
            {
                this.ValidateValueAsArray(value);
                return value;
            }
            else if (value.GetType() == this.Type)
            {
                return value;
            }
            else
            {
                try
                {
                    return Convert.ChangeType(value, this.Type, CultureInfo.InvariantCulture);
                }
                catch (InvalidCastException ex)
                {
                    throw new ArgumentException(
                        string.Format(
                            Strings.ValueOfTypeCannotBeConverted,
                            value,
                            value.GetType(),
                            this.Type),
                        ex);
                }
            }
        }

        /// <summary>
        /// Converts a string to value consistent with type.
        /// </summary>
        /// <param name="stringValue">String to convert to a value.</param>
        /// <returns>Value.</returns>
        internal object ConvertToValue(string stringValue)
        {
            try
            {
                return this.Parse(stringValue);
            }
            catch (FormatException ex)
            {
                throw new ServiceXmlDeserializationException(
                    string.Format(
                        Strings.ValueCannotBeConverted,
                        stringValue,
                        this.Type),
                    ex);
            }
            catch (InvalidCastException ex)
            {
                throw new ServiceXmlDeserializationException(
                    string.Format(
                        Strings.ValueCannotBeConverted,
                        stringValue,
                        this.Type),
                    ex);
            }
            catch (OverflowException ex)
            {
                throw new ServiceXmlDeserializationException(
                    string.Format(
                        Strings.ValueCannotBeConverted,
                        stringValue,
                        this.Type),
                    ex);
            }
        }

        /// <summary>
        /// Converts a string to value consistent with type (or uses the default value if the string is null or empty).
        /// </summary>
        /// <param name="stringValue">String to convert to a value.</param>
        /// <returns>Value.</returns>
        /// <remarks>For array types, this method is called for each array element.</remarks>
        internal object ConvertToValueOrDefault(string stringValue)
        {
            return string.IsNullOrEmpty(stringValue) ? this.DefaultValue : this.ConvertToValue(stringValue);
        }

        /// <summary>
        /// Validates array value.
        /// </summary>
        /// <param name="value">The value.</param>
        private void ValidateValueAsArray(object value)
        {
            Array array = value as Array;
            if (array == null)
            {
                throw new ArgumentException(
                    string.Format(
                        Strings.IncompatibleTypeForArray,
                        value.GetType(),
                        this.Type));
            }
            else if (array.Rank != 1)
            {
                throw new ArgumentException(Strings.ArrayMustHaveSingleDimension);
            }
            else if (array.Length == 0)
            {
                throw new ArgumentException(Strings.ArrayMustHaveAtLeastOneElement);
            }
            else if (array.GetType().GetElementType() != this.Type)
            {
                throw new ArgumentException(
                    string.Format(
                        Strings.IncompatibleTypeForArray,
                        value.GetType(),
                        this.Type));
            }
        }

        #region Properties

        /// <summary>
        /// Gets or sets the string parser.
        /// </summary>
        /// <remarks>For array types, this method is called for each array element.</remarks>
        internal Func<string, object> Parse
        { 
            get; set; 
        }

        /// <summary>
        /// Gets or sets the string to object converter.
        /// </summary>
        /// <remarks>For array types, this method is called for each array element.</remarks>
        internal Func<object, string> ConvertToString
        {
            get; set;
        }

        /// <summary>
        /// Gets or sets the type.
        /// </summary>
        /// <remarks>For array types, this is the type of an element.</remarks>
        internal Type Type
        {
            get; set;
        }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is array.
        /// </summary>
        /// <value><c>true</c> if this instance is array; otherwise, <c>false</c>.</value>
        internal bool IsArray
        {
            get; set;
        }

        /// <summary>
        /// Gets the default value for the type.
        /// </summary>
        internal object DefaultValue
        {
            get { return defaultValueMap.Member[this.Type]; }
        }

        #endregion
    }
}
