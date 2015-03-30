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
using System.Text.RegularExpressions;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Class to parse a JSON stream into an instance of a JSON object.
    /// </summary>
    /// <remarks>See http://www.ietf.org/rfc/rfc4627.txt</remarks>
    internal class JsonParser
    {
        private JsonTokenizer tokenizer;

        /// <summary>
        /// Initializes a new instance of the <see cref="JsonParser"/> class.
        /// </summary>
        /// <param name="inputStream">The input stream.</param>
        internal JsonParser(Stream inputStream)
        {
            this.tokenizer = new JsonTokenizer(inputStream);
        }

        internal JsonObject Parse()
        {
            return this.ParseObject();
        }

        /// <summary>
        /// Parses the object.
        /// </summary>
        /// <returns></returns>
        private JsonObject ParseObject()
        {
            JsonObject jsonObject = new JsonObject();

            string token;

            this.ReadAndValidateToken(out token, JsonTokenType.ObjectOpen);

            while (this.tokenizer.Peek() != JsonTokenType.ObjectClose)
            {
                this.ParseKeyValuePair(jsonObject);

                if (this.tokenizer.Peek() != JsonTokenType.Comma)
                {
                    break;
                }
                else
                {
                    this.ReadAndValidateToken(out token, JsonTokenType.Comma);
                }
            }

            this.ReadAndValidateToken(out token, JsonTokenType.ObjectClose);

            return jsonObject;
        }

        /// <summary>
        /// Parses the key value pair.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        private void ParseKeyValuePair(JsonObject jsonObject)
        {
            string name;
            string colon;

            this.ReadAndValidateToken(out name, JsonTokenType.String);
            this.ReadAndValidateToken(out colon, JsonTokenType.Colon);

            name = this.UnescapeString(name);

            jsonObject.Add(name, this.ParseValue());
        }

        /// <summary>
        /// Parses the value.
        /// </summary>
        /// <returns></returns>
        private object ParseValue()
        {
            string valueToken;

            switch (this.tokenizer.Peek())
            {
                case JsonTokenType.ArrayOpen:
                    object[] jsonArray = this.ParseArray();
                    return jsonArray;

                case JsonTokenType.ObjectOpen:
                    JsonObject jsonChildObject = this.ParseObject();
                    return jsonChildObject;

                case JsonTokenType.String:
                    this.ReadAndValidateToken(out valueToken, JsonTokenType.String);
                    return this.UnescapeString(valueToken);

                case JsonTokenType.Null:
                    this.ReadAndValidateToken(out valueToken, JsonTokenType.Null);
                    return null;

                case JsonTokenType.Boolean:
                    this.ReadAndValidateToken(out valueToken, JsonTokenType.Boolean);
                    return bool.Parse(valueToken);

                case JsonTokenType.Number:
                    this.ReadAndValidateToken(out valueToken, JsonTokenType.Number);
                    return this.ParseNumber(valueToken);

                default:
                    // TODO: Add a message to better locate the error?
                    throw new ServiceJsonDeserializationException();
            }
        }

        /// <summary>
        /// Parses the number.
        /// </summary>
        /// <param name="valueToken">The value token.</param>
        /// <returns></returns>
        private object ParseNumber(string valueToken)
        {
            if (Regex.IsMatch(valueToken, @"^-?\d+$"))
            {
                return long.Parse(valueToken, CultureInfo.InvariantCulture);
            }
            else
            {
                return double.Parse(valueToken, CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        /// Parses the array.
        /// </summary>
        /// <returns></returns>
        private object[] ParseArray()
        {
            List<object> objects = new List<object>();
            string token;

            this.ReadAndValidateToken(out token, JsonTokenType.ArrayOpen);

            while (this.tokenizer.Peek() != JsonTokenType.ArrayClose)
            {
                objects.Add(this.ParseValue());

                if (this.tokenizer.Peek() != JsonTokenType.Comma)
                {
                    break;
                }
                else
                {
                    this.ReadAndValidateToken(out token, JsonTokenType.Comma);
                }
            }

            this.ReadAndValidateToken(out token, JsonTokenType.ArrayClose);

            return objects.ToArray();
        }

        /// <summary>
        /// Unescapes the string.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        private string UnescapeString(string value)
        {
            // Trim the quotes off the ends
            string result = value.Substring(1, value.Length - 2);

            // See if we nee to bother decoding the string
            if (result.Contains('\\'))
            {
                if (result.Contains(@"\\"))
                {
                    result = result.Replace(@"\\", @"\");
                }
                if (result.Contains(@"\"""))
                {
                    result = result.Replace(@"\""", @"""");
                }
                if (result.Contains(@"\/"))
                {
                    result = result.Replace(@"\/", @"/");
                }
                if (result.Contains(@"\b"))
                {
                    result = result.Replace(@"\b", "\b");
                }
                if (result.Contains(@"\f"))
                {
                    result = result.Replace(@"\f", "\f");
                }
                if (result.Contains(@"\n"))
                {
                    result = result.Replace(@"\n", "\n");
                }
                if (result.Contains(@"\r"))
                {
                    result = result.Replace(@"\r", "\r");
                }
                if (result.Contains(@"\t"))
                {
                    result = result.Replace(@"\t", "\t");
                }
                if (result.Contains(@"\u"))
                {
                    MatchCollection unicodeMatches = Regex.Matches(result, @"\\u([\da-fA-F]{4})");
                    
                    foreach (Match currentMatch in unicodeMatches)
                    {
                        if (currentMatch.Success)
                        {
                            int hexCode = int.Parse(currentMatch.Value.Substring(2), NumberStyles.HexNumber);
                            string charToReplace = char.ConvertFromUtf32(hexCode);

                            result = result.Replace(currentMatch.Value, charToReplace);
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Reads the and validate token.
        /// </summary>
        /// <param name="token">The token.</param>
        /// <param name="expectedTokenTypes">The expected token types.</param>
        /// <returns></returns>
        private JsonTokenType ReadAndValidateToken(out string token, params JsonTokenType[] expectedTokenTypes)
        {
            JsonTokenType tokenType = this.tokenizer.NextToken(out token);

            foreach (JsonTokenType expectedToken in expectedTokenTypes)
            {
                if (tokenType == expectedToken)
                {
                    return tokenType;
                }
            }
            
            // TODO: Add a message to better locate the error?
            throw new ServiceJsonDeserializationException();
        }
    }
}