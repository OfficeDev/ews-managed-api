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

    /// <summary>
    /// Represents a collection of DayOfTheWeek values.
    /// </summary>
    public sealed class DayOfTheWeekCollection : ComplexProperty, IEnumerable<DayOfTheWeek>
    {
        private List<DayOfTheWeek> items = new List<DayOfTheWeek>();

        /// <summary>
        /// Initializes a new instance of the <see cref="DayOfTheWeekCollection"/> class.
        /// </summary>
        internal DayOfTheWeekCollection()
        {
        }

        /// <summary>
        /// Convert to string.
        /// </summary>
        /// <param name="separator">The separator.</param>
        /// <returns>String representation of collection.</returns>
        private string ToString(string separator)
        {
            if (this.Count == 0)
            {
                return string.Empty;
            }
            else
            {
                string[] daysOfTheWeekArray = new string[this.Count];

                for (int i = 0; i < this.Count; i++)
                {
                    daysOfTheWeekArray[i] = this[i].ToString();
                }

                return string.Join(separator, daysOfTheWeekArray);
            }
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal override void LoadFromXml(EwsServiceXmlReader reader, string xmlElementName)
        {
            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, xmlElementName);

            EwsUtilities.ParseEnumValueList<DayOfTheWeek>(
                this.items,
                reader.ReadElementValue(),
                ' ');
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service"></param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            // DayOfTheWeekCollection is implemented as a Json primitive, rather than as a Json object.
            throw new InvalidOperationException();
        }

        /// <summary>
        /// Loads from json value.
        /// </summary>
        /// <param name="jsonValue">The json value.</param>
        internal void LoadFromJsonValue(string jsonValue)
        {
            EwsUtilities.ParseEnumValueList<DayOfTheWeek>(
                this.items,
                jsonValue,
                ' ');
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal override void WriteToXml(EwsServiceXmlWriter writer, string xmlElementName)
        {
            string daysOfWeekAsString = this.ToString(" ");

            if (!string.IsNullOrEmpty(daysOfWeekAsString))
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.DaysOfWeek,
                    daysOfWeekAsString);
            }
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            return this.ToString(" ");
        }

        /// <summary>
        /// Builds string representation of the collection.
        /// </summary>
        /// <returns>A comma-delimited string representing the collection.</returns>
        public override string ToString()
        {
            return this.ToString(",");
        }

        /// <summary>
        /// Adds a day to the collection if it is not already present.
        /// </summary>
        /// <param name="dayOfTheWeek">The day to add.</param>
        public void Add(DayOfTheWeek dayOfTheWeek)
        {
            if (!this.items.Contains(dayOfTheWeek))
            {
                this.items.Add(dayOfTheWeek);
                this.Changed();
            }
        }

        /// <summary>
        /// Adds multiple days to the collection if they are not already present.
        /// </summary>
        /// <param name="daysOfTheWeek">The days to add.</param>
        public void AddRange(IEnumerable<DayOfTheWeek> daysOfTheWeek)
        {
            foreach (DayOfTheWeek dayOfTheWeek in daysOfTheWeek)
            {
                this.Add(dayOfTheWeek);
            }
        }

        /// <summary>
        /// Clears the collection.
        /// </summary>
        public void Clear()
        {
            if (this.Count > 0)
            {
                this.items.Clear();
                this.Changed();
            }
        }

        /// <summary>
        /// Remove a specific day from the collection.
        /// </summary>
        /// <param name="dayOfTheWeek">The day to remove.</param>
        /// <returns>True if the day was removed from the collection, false otherwise.</returns>
        public bool Remove(DayOfTheWeek dayOfTheWeek)
        {
            bool result = this.items.Remove(dayOfTheWeek);

            if (result)
            {
                this.Changed();
            }

            return result;
        }

        /// <summary>
        /// Removes the day at a specific index.
        /// </summary>
        /// <param name="index">The index of the day to remove.</param>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= this.Count)
            {
                throw new ArgumentOutOfRangeException("index", Strings.IndexIsOutOfRange);
            }

            this.items.RemoveAt(index);
            this.Changed();
        }

        /// <summary>
        /// Gets the DayOfTheWeek at a specific index in the collection.
        /// </summary>
        /// <param name="index">Index</param>
        /// <returns>DayOfTheWeek at index</returns>
        public DayOfTheWeek this[int index]
        {
            get
            {
                return this.items[index];
            }
        }

        /// <summary>
        /// Gets the number of days in the collection.
        /// </summary>
        public int Count
        {
            get { return this.items.Count; }
        }

        #region IEnumerable<DayOfTheWeek> Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<DayOfTheWeek> GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        #endregion
    }
}