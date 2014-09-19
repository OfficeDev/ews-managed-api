// ---------------------------------------------------------------------------
// <copyright file="JsonTokenizer.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// The various tokens this tokenizer recognizes
    /// </summary>
    internal enum JsonTokenType
    {
        /// <summary>
        /// "chars" or ""
        /// </summary>
        String,

        /// <summary>
        /// digits with optional negative sign, fractional component, and/or exponent
        /// </summary>
        Number,

        /// <summary>
        /// true or false
        /// </summary>
        Boolean,

        /// <summary>
        /// null
        /// </summary>
        Null,

        /// <summary>
        /// {
        /// </summary>
        ObjectOpen,

        /// <summary>
        /// }
        /// </summary>
        ObjectClose,

        /// <summary>
        /// [
        /// </summary>
        ArrayOpen,

        /// <summary>
        /// ]
        /// </summary>
        ArrayClose,

        /// <summary>
        /// :
        /// </summary>
        Colon,

        /// <summary>
        /// ,
        /// </summary>
        Comma,

        /// <summary>
        /// EOF
        /// </summary>
        EndOfFile,
    }

    /// <summary>
    /// Class to break a JSON stream into its component tokens to be consumed by a JSON parser.
    /// </summary>
    internal class JsonTokenizer
    {
        /// <summary>
        /// Matches:
        ///     ""
        /// or
        ///     "chars"
        /// where 'chars' includes any unicode character except \ or ", plus the escaped characters below.
        /// </summary>
        private const string JsonStringRegExCode = @"""([^\\""]|\\""|\\\\|\\/|\\b|\\f|\\n|\\r|\\t|\\u[\da-fA-F]{4})*""";

        /// <summary>
        /// Matches numbers with an optional leading negative, optional decimal, and optional exponent.
        /// </summary>
        private const string JsonNumberRegExCode = @"-?\d+(.\d+)?([eE][+-]?\d+)?";

        /// <summary>
        /// Matches true or false;
        /// </summary>
        private const string JsonBooleanRegExCode = @"(true|false)";

        /// <summary>
        /// Matches null
        /// </summary>
        private const string JsonNullRegExCode = @"null";

        /// <summary>
        /// Matches {
        /// </summary>
        private const string JsonOpenObjectRegExCode = @"\{";

        /// <summary>
        /// Matches }
        /// </summary>
        private const string JsonCloseObjectRegExCode = @"\}";

        /// <summary>
        /// Matches [
        /// </summary>
        private const string JsonOpenArrayRegExCode = @"\[";

        /// <summary>
        /// Matches ]
        /// </summary>
        private const string JsonCloseArrayRegExCode = @"\]";

        /// <summary>
        /// Matches :
        /// </summary>
        private const string JsonColonRegExCode = @"\:";

        /// <summary>
        /// Matches ,
        /// </summary>
        private const string JsonCommaRegExCode = @"\,";

        private static Regex jsonStringRegEx;
        private static Regex jsonNumberRegEx;
        private static Regex jsonBooleanRegEx;
        private static Regex jsonNullRegEx;
        private static Regex jsonOpenObjectRegEx;
        private static Regex jsonCloseObjectRegEx;
        private static Regex jsonOpenArrayRegEx;
        private static Regex jsonCloseArrayRegEx;
        private static Regex jsonColonRegEx;
        private static Regex jsonCommaRegEx;
        private static Regex whitespaceRegEx;

        private static Dictionary<JsonTokenType, Regex> tokenDictionary;
        private static Regex fullTokenizerRegex;

        static JsonTokenizer()
        {
            jsonStringRegEx = new Regex(JsonStringRegExCode, RegexOptions.Compiled);
            jsonNumberRegEx = new Regex(JsonNumberRegExCode, RegexOptions.Compiled);
            jsonBooleanRegEx = new Regex(JsonBooleanRegExCode, RegexOptions.Compiled);
            jsonNullRegEx = new Regex(JsonNullRegExCode, RegexOptions.Compiled);
            jsonOpenObjectRegEx = new Regex(JsonOpenObjectRegExCode, RegexOptions.Compiled);
            jsonCloseObjectRegEx = new Regex(JsonCloseObjectRegExCode, RegexOptions.Compiled);
            jsonOpenArrayRegEx = new Regex(JsonOpenArrayRegExCode, RegexOptions.Compiled);
            jsonCloseArrayRegEx = new Regex(JsonCloseArrayRegExCode, RegexOptions.Compiled);
            jsonColonRegEx = new Regex(JsonColonRegExCode, RegexOptions.Compiled);
            jsonCommaRegEx = new Regex(JsonCommaRegExCode, RegexOptions.Compiled);

            whitespaceRegEx = new Regex("\\s");

            tokenDictionary = new Dictionary<JsonTokenType, Regex>();
            tokenDictionary.Add(JsonTokenType.Number, jsonNumberRegEx);
            tokenDictionary.Add(JsonTokenType.Boolean, jsonBooleanRegEx);
            tokenDictionary.Add(JsonTokenType.Null, jsonNullRegEx);
            tokenDictionary.Add(JsonTokenType.ObjectOpen, jsonOpenObjectRegEx);
            tokenDictionary.Add(JsonTokenType.ObjectClose, jsonCloseObjectRegEx);
            tokenDictionary.Add(JsonTokenType.ArrayOpen, jsonOpenArrayRegEx);
            tokenDictionary.Add(JsonTokenType.ArrayClose, jsonCloseArrayRegEx);
            tokenDictionary.Add(JsonTokenType.Colon, jsonColonRegEx);
            tokenDictionary.Add(JsonTokenType.Comma, jsonCommaRegEx);
            tokenDictionary.Add(JsonTokenType.String, jsonStringRegEx);

            StringBuilder tokenizerRegExCode = new StringBuilder();
            bool firstEntry = true;

            foreach (Regex regEx in tokenDictionary.Values)
            {
                if (firstEntry)
                {
                    firstEntry = false;
                }
                else
                {
                    tokenizerRegExCode.Append("|");
                }

                tokenizerRegExCode.Append("(");
                tokenizerRegExCode.Append(regEx.ToString());
                tokenizerRegExCode.Append(")");
            }

            fullTokenizerRegex = new Regex(tokenizerRegExCode.ToString(), RegexOptions.Compiled);
        }

        private Match currentMatch;

        private JsonTokenType nextTokenType;
        private string nextToken;
        private bool nextTokenPeeked = false;

        internal JsonTokenizer(Stream input)
        {
            StreamReader reader = new StreamReader(input);
            string response = reader.ReadToEnd();
            currentMatch = fullTokenizerRegex.Match(response);            
        }

        internal JsonTokenType Peek()
        {
            if (!this.nextTokenPeeked)
            {
                this.nextTokenType = this.NextToken(out this.nextToken);
                this.nextTokenPeeked = true;
            }

            return this.nextTokenType;
        }

        internal JsonTokenType NextToken(out string token)
        {
            if (this.nextTokenPeeked)
            {
                token = this.nextToken;
                this.nextTokenPeeked = false;

                return this.nextTokenType;
            }
            else
            {
                token = this.currentMatch.Value;

                while (token.Trim().Length == 0)
                {
                    // skip through whitespace and null tokens
                    this.AdvanceRegExMatch();
                    token = this.currentMatch.Value;
                }

                foreach (KeyValuePair<JsonTokenType, Regex> tokenRegexPair in tokenDictionary)
                {
                    Match regexMatch = tokenRegexPair.Value.Match(token);
                    if (regexMatch.Success 
                        && regexMatch.Index == 0
                        && regexMatch.Length == token.Length)
                    {
                        this.AdvanceRegExMatch();
                        return tokenRegexPair.Key;
                    }
                }

                throw new ServiceJsonDeserializationException();
            }
        }

        private void AdvanceRegExMatch()
        {
            this.currentMatch = this.currentMatch.NextMatch();
        }
    }
}
