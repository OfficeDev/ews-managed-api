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

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents a JSON construction of an object.
    /// Used for serialization and deserialization.
    /// </summary>
    internal class JsonObject : Dictionary<string, object>
    {
        /// <summary>
        /// Special property name used by EWS JSON endpoint to indicate the object type.
        /// </summary>
        private const string TypeAttribute = "__type";

        /// <summary>
        /// Namespace for Exchange JSON types.
        /// </summary>
        private const string JsonTypeNamespace = "Exchange";

        /// <remarks> 
        /// Used for existing XmlElements that have attributes and a text value.
        /// Eg., "<![CDATA[<Body BodyType="HTML">Hello, World!</Body>]]>"
        /// This property is the key for the value of the text element in such an XML Element.
        /// </remarks>
        internal const string JsonValueString = "Value";

        /// <summary>
        /// Validates the object.
        /// </summary>
        /// <param name="entry">The entry.</param>
        private static void ValidateObject(object entry)
        {
            if (!(entry == null ||
                entry is JsonObject ||
                entry is Enum ||
                entry is bool ||
                entry is string ||
                entry is int ||
                entry is float ||
                entry is double ||
                entry is long ||
                entry is DateTime ||
                entry.GetType().IsArray))
            {
                throw new InvalidOperationException(String.Format("Object [{0}] in the array is not serializable to JSON", entry));
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="JsonObject"/> class.
        /// </summary>
        internal JsonObject()
        {
        }

               /// <summary>
        /// Adds the specified name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        internal void Add(string name, JsonObject value)
        {
            this.InternalAdd(name, value);
        }

        /// <summary>
        /// Adds the specified name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        internal void Add(string name, string value)
        {
            if (value != null)
            {
                this.InternalAdd(name, value);
            }
        }

        /// <summary>
        /// Adds the specified name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        internal void Add(string name, Enum value)
        {
            if (value != null)
            {
                this.InternalAdd(name, value.ToString());
            }
        }

        /// <summary>
        /// Adds the specified name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        internal void Add(string name, bool value)
        {
            this.InternalAdd(name, value);
        }

        /// <summary>
        /// Adds the specified name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        internal void Add(string name, int value)
        {
            this.InternalAdd(name, value);
        }

        /// <summary>
        /// Adds the specified name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        internal void Add(string name, float value)
        {
            this.InternalAdd(name, value);
        }

        /// <summary>
        /// Adds the specified name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        internal void Add(string name, double value)
        {
            this.InternalAdd(name, value);
        }

        /// <summary>
        /// Adds the specified name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        internal void Add(string name, long value)
        {
            this.InternalAdd(name, value);
        }

        /// <summary>
        /// Adds the specified name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        public new void Add(string name, object value)
        {
            ValidateObject(value);
            this.InternalAdd(name, value);
        }

        private void InternalAdd(string name, object value)
        {
            base.Add(name, value);
        }

        /// <summary>
        /// Adds the type parameter.
        /// </summary>
        /// <param name="typeName">Name of the type.</param>
        internal void AddTypeParameter(string typeName)
        {
            this.InternalAdd(
                JsonObject.TypeAttribute,
                String.Format("{0}:#{1}", typeName, JsonObject.JsonTypeNamespace));
        }

        /// <summary>
        /// Adds the specified name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        internal void Add(string name, IEnumerable<object> value)
        {
            object[] valueArray = value.ToArray();

            foreach (object entry in valueArray)
            {
                ValidateObject(entry);              
            }

            this.InternalAdd(name, valueArray);
        }

        /// <summary>
        /// Adds the base64.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="stream">The stream.</param>
        internal void AddBase64(string key, Stream stream)
        {
            // We use a memory stream here because we don't know that a generic Stream can tell us how long it is.
            using (MemoryStream buffer = new MemoryStream())
            {
                byte[] byteBuffer = new byte[4096];
                int bytesRead = 0;

                while ((bytesRead = stream.Read(byteBuffer, 0, byteBuffer.Length)) != 0)
                {
                    buffer.Write(byteBuffer, 0, bytesRead);
                }

                this.AddBase64(key, buffer.GetBuffer(), 0, (int)buffer.Length);
            }
        }

        /// <summary>
        /// Adds the base64.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="buffer">The buffer.</param>
        internal void AddBase64(string key, byte[] buffer)
        {
            this.AddBase64(key, buffer, 0, buffer.Length);
        }

        /// <summary>
        /// Adds the base64.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="buffer">The buffer.</param>
        /// <param name="offset">The offset.</param>
        /// <param name="count">The count.</param>
        internal void AddBase64(string key, byte[] buffer, int offset, int count)
        {
            this.InternalAdd(key, Convert.ToBase64String(buffer, offset, count));
        }

        /// <summary>
        /// Serializes to JSON.
        /// </summary>
        /// <param name="stream">The stream.</param>
        internal void SerializeToJson(Stream stream)
        {
            this.SerializeToJson(stream, false);
        }

        /// <summary>
        /// Serializes to JSON.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="prettyPrint">if true, pretty-print the results.</param>
        internal void SerializeToJson(Stream stream, bool prettyPrint)
        {
            using (JsonWriter writer = new JsonWriter(stream, prettyPrint))
            {
                this.WriteValue(writer, this);

                writer.Flush();
            }
        }

        /// <summary>
        /// Writes key value pair.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="key">The key.</param>
        /// <param name="value">The value.</param>
        private void WriteKeyValuePair(JsonWriter writer, string key, object value)
        {
            writer.WriteKey(key);
            this.WriteValue(writer, value);
        }

        /// <summary>
        /// Writes the value.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="value">The value.</param>
        private void WriteValue(JsonWriter writer, object value)
        {
            if (value == null)
            {
                writer.WriteNullValue();
            }
            else if (value is string)
            {
                writer.WriteValue((string)value);
            } 
            else if (value.GetType().IsEnum)
            {
                writer.WriteValue((Enum)value);
            }
            else if (value is int) 
            {
                writer.WriteValue((int)value);
            }
            else if (value is long)
            {
                writer.WriteValue((long)value);
            }
            else if (value is float)
            {
                writer.WriteValue((float)value);
            }
            else if (value is double)
            {
                writer.WriteValue((double)value);
            }
            else if (value is bool)
            {
                writer.WriteValue((bool)value);
            }
            else if (value is DateTime)
            {
                writer.WriteValue((DateTime)value);
            }
            else if (value is JsonObject)
            {
                writer.PushObjectClosure();

                JsonObject jsObject = value as JsonObject;

                // Wcf JSON deserializer requires the __type attribute to be first.
                //
                if (jsObject.ContainsKey(JsonObject.TypeAttribute))
                {
                    this.WriteKeyValuePair(writer, JsonObject.TypeAttribute, jsObject[JsonObject.TypeAttribute]);
                }

                foreach (KeyValuePair<string, object> kvPair in jsObject)
                {
                    if (kvPair.Key != JsonObject.TypeAttribute)
                    {
                        this.WriteKeyValuePair(writer, kvPair.Key, kvPair.Value);
                    }
                }

                writer.PopClosure();
            }
            else if (value.GetType().IsArray)
            {
                writer.PushArrayClosure();

                foreach (var arrayEntry in (Array)value)
                {
                    this.WriteValue(writer, arrayEntry);
                }

                writer.PopClosure();
            }
            else
            {
                throw new InvalidOperationException(String.Format("Object [{0}] is not JSON serializable", value));
            }
        }

        /// <summary>
        /// Reads the value for the selected key as an int.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns></returns>
        internal int ReadAsInt(string key)
        {
            if (!this.ContainsKey(key))
            {
                throw new ServiceJsonDeserializationException();
            }

            object value = this[key];

            if (!(value is long))
            {
                throw new ServiceJsonDeserializationException();
            }

            return (int)(long)value;
        }

        /// <summary>
        /// Reads the value for the selected key as an double.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns></returns>
        internal double ReadAsDouble(string key)
        {
            if (!this.ContainsKey(key))
            {
                throw new ServiceJsonDeserializationException();
            }

            object value = this[key];

            if (!(value is long))
            {
                throw new ServiceJsonDeserializationException();
            }

            return (double)(long)value;
        }

        /// <summary>
        /// Reads the value for the selected key as a string.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns></returns>
        internal string ReadAsString(string key)
        {
            if (!this.ContainsKey(key))
            {
                throw new ServiceJsonDeserializationException();
            }

            object value = this[key];

            if (value != null &&
                !(value is string))
            {
                throw new ServiceJsonDeserializationException();
            }

            return value as string;
        }

        /// <summary>
        /// Reads the value for the selected key as a JSON object.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns></returns>
        internal JsonObject ReadAsJsonObject(string key)
        {
            if (!this.ContainsKey(key))
            {
                throw new ServiceJsonDeserializationException();
            }

            object value = this[key];

            if (value != null &&
                !(value is JsonObject))
            {
                throw new ServiceJsonDeserializationException();
            }

            return value as JsonObject;
        }

        /// <summary>
        /// Reads the value for the selected key as a JSON object.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns></returns>
        internal object[] ReadAsArray(string key)
        {
            if (!this.ContainsKey(key))
            {
                return new object[0];
            }

            object value = this[key];

            if (value != null &&
                !(value is object[]))
            {
                throw new ServiceJsonDeserializationException();
            }

            return value as object[];
        }

        /// <summary>
        /// Determines whether object has type property.
        /// </summary>
        /// <returns>Returns true if JsonObject has a type property.</returns>
        internal bool HasTypeProperty()
        {
            return this.ContainsKey(JsonObject.TypeAttribute);
        }

        /// <summary>
        /// Reads the type string.
        /// </summary>
        /// <returns></returns>
        internal string ReadTypeString()
        {
            if (!this.HasTypeProperty())
            {
                throw new ServiceJsonDeserializationException();
            }

            string typeAttribute = this.ReadAsString(JsonObject.TypeAttribute);

            return typeAttribute.Substring(
                0, 
                typeAttribute.IndexOf(String.Format(":#{0}", JsonObject.JsonTypeNamespace)));
        }

        /// <summary>
        /// Reads the enum value.
        /// </summary>
        /// <typeparam name="T">An enum type</typeparam>
        /// <param name="key">The key.</param>
        /// <returns></returns>
        internal T ReadEnumValue<T>(string key)
        {
            return EwsUtilities.Parse<T>(this.ReadAsString(key));
        }

        /// <summary>
        /// Reads as bool.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns></returns>
        internal bool ReadAsBool(string key)
        {
            if (!this.ContainsKey(key))
            {
                throw new ServiceJsonDeserializationException();
            }

            object value = this[key];

            if (!(value is bool))
            {
                throw new ServiceJsonDeserializationException();
            }

            return (bool)value;
        }

        /// <summary>
        /// Reads the content as base64.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="stream">The stream.</param>
        internal void ReadAsBase64Content(string key, System.IO.Stream stream)
        {
            byte[] buffer = ReadAsBase64Content(key);

            stream.Write(buffer, 0, buffer.Length);
        }

        /// <summary>
        /// Reads the content of as base64.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns></returns>
        internal byte[] ReadAsBase64Content(string key)
        {
            string base64Content = this.ReadAsString(key);

            byte[] buffer = Convert.FromBase64String(base64Content);
            return buffer;
        }
    }
}