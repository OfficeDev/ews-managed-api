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
    using System.Collections;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Reflection;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Xml;

    /// <summary>
    /// EWS utilities
    /// </summary>
    internal static class EwsUtilities
    {
        #region Private members

        /// <summary>
        /// Map from XML element names to ServiceObject type and constructors. 
        /// </summary>
        private static LazyMember<ServiceObjectInfo> serviceObjectInfo = new LazyMember<ServiceObjectInfo>(
            delegate()
            {
                return new ServiceObjectInfo();
            });

        /// <summary>
        /// Version of API binary.
        /// </summary>
        private static LazyMember<string> buildVersion = new LazyMember<string>(
            delegate()
            {
                try
                {
                    FileVersionInfo fileInfo = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
                    return fileInfo.FileVersion;
                }
                catch
                {
                    // OM:2026839 When run in an environment with partial trust, fetching the build version blows up.
                    // Just return a hardcoded value on failure.
                    return "0.0";
                }
            });

        /// <summary>
        /// Dictionary of enum type to ExchangeVersion maps. 
        /// </summary>
        private static LazyMember<Dictionary<Type, Dictionary<Enum, ExchangeVersion>>> enumVersionDictionaries = new LazyMember<Dictionary<Type, Dictionary<Enum, ExchangeVersion>>>(
            () => new Dictionary<Type, Dictionary<Enum, ExchangeVersion>>()
            {
                { typeof(WellKnownFolderName), BuildEnumDict(typeof(WellKnownFolderName)) },
                { typeof(ItemTraversal), BuildEnumDict(typeof(ItemTraversal)) },
                { typeof(ConversationQueryTraversal), BuildEnumDict(typeof(ConversationQueryTraversal)) },
                { typeof(FileAsMapping), BuildEnumDict(typeof(FileAsMapping)) },
                { typeof(EventType), BuildEnumDict(typeof(EventType)) },
                { typeof(MeetingRequestsDeliveryScope), BuildEnumDict(typeof(MeetingRequestsDeliveryScope)) },
                { typeof(ViewFilter), BuildEnumDict(typeof(ViewFilter)) },
            });

        /// <summary>
        /// Dictionary of enum type to schema-name-to-enum-value maps.
        /// </summary>
        private static LazyMember<Dictionary<Type, Dictionary<string, Enum>>> schemaToEnumDictionaries = new LazyMember<Dictionary<Type, Dictionary<string, Enum>>>(
            () => new Dictionary<Type, Dictionary<string, Enum>>
            {
                { typeof(EventType), BuildSchemaToEnumDict(typeof(EventType)) },
                { typeof(MailboxType), BuildSchemaToEnumDict(typeof(MailboxType)) },
                { typeof(FileAsMapping), BuildSchemaToEnumDict(typeof(FileAsMapping)) },
                { typeof(RuleProperty), BuildSchemaToEnumDict(typeof(RuleProperty)) },
                { typeof(WellKnownFolderName), BuildSchemaToEnumDict(typeof(WellKnownFolderName)) },
            });

        /// <summary>
        /// Dictionary of enum type to enum-value-to-schema-name maps.
        /// </summary>
        private static LazyMember<Dictionary<Type, Dictionary<Enum, string>>> enumToSchemaDictionaries = new LazyMember<Dictionary<Type, Dictionary<Enum, string>>>(
            () => new Dictionary<Type, Dictionary<Enum, string>>
            {
                { typeof(EventType), BuildEnumToSchemaDict(typeof(EventType)) },
                { typeof(MailboxType), BuildEnumToSchemaDict(typeof(MailboxType)) },
                { typeof(FileAsMapping), BuildEnumToSchemaDict(typeof(FileAsMapping)) },
                { typeof(RuleProperty), BuildEnumToSchemaDict(typeof(RuleProperty)) },
                { typeof(WellKnownFolderName), BuildEnumToSchemaDict(typeof(WellKnownFolderName)) },
            });

        /// <summary>
        /// Dictionary to map from special CLR type names to their "short" names.
        /// </summary>
        private static LazyMember<Dictionary<string, string>> typeNameToShortNameMap = new LazyMember<Dictionary<string, string>>(
            () => new Dictionary<string, string>
            {
                { "Boolean", "bool" },
                { "Int16", "short" },
                { "Int32", "int" },
                { "String", "string" }
            });
        #endregion

        #region Constants

        internal const string XSFalse = "false";
        internal const string XSTrue = "true";

        internal const string EwsTypesNamespacePrefix = "t";
        internal const string EwsMessagesNamespacePrefix = "m";
        internal const string EwsErrorsNamespacePrefix = "e";
        internal const string EwsSoapNamespacePrefix = "soap";
        internal const string EwsXmlSchemaInstanceNamespacePrefix = "xsi";
        internal const string PassportSoapFaultNamespacePrefix = "psf";
        internal const string WSTrustFebruary2005NamespacePrefix = "wst";
        internal const string WSAddressingNamespacePrefix = "wsa";
        internal const string AutodiscoverSoapNamespacePrefix = "a";
        internal const string WSSecurityUtilityNamespacePrefix = "wsu";
        internal const string WSSecuritySecExtNamespacePrefix = "wsse";

        internal const string EwsTypesNamespace = "http://schemas.microsoft.com/exchange/services/2006/types";
        internal const string EwsMessagesNamespace = "http://schemas.microsoft.com/exchange/services/2006/messages";
        internal const string EwsErrorsNamespace = "http://schemas.microsoft.com/exchange/services/2006/errors";
        internal const string EwsSoapNamespace = "http://schemas.xmlsoap.org/soap/envelope/";
        internal const string EwsSoap12Namespace = "http://www.w3.org/2003/05/soap-envelope";
        internal const string EwsXmlSchemaInstanceNamespace = "http://www.w3.org/2001/XMLSchema-instance";
        internal const string PassportSoapFaultNamespace = "http://schemas.microsoft.com/Passport/SoapServices/SOAPFault";
        internal const string WSTrustFebruary2005Namespace = "http://schemas.xmlsoap.org/ws/2005/02/trust";
        internal const string WSAddressingNamespace = "http://www.w3.org/2005/08/addressing"; // "http://schemas.xmlsoap.org/ws/2004/08/addressing";
        internal const string AutodiscoverSoapNamespace = "http://schemas.microsoft.com/exchange/2010/Autodiscover";
        internal const string WSSecurityUtilityNamespace = "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd";
        internal const string WSSecuritySecExtNamespace = "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd";

        /// <summary>
        /// Regular expression for legal domain names.
        /// </summary>
        internal const string DomainRegex = "^[-a-zA-Z0-9_.]+$";
        #endregion

        /// <summary>
        /// Asserts that the specified condition if true.
        /// </summary>
        /// <param name="condition">Assertion.</param>
        /// <param name="caller">The caller.</param>
        /// <param name="message">The message to use if assertion fails.</param>
        internal static void Assert(
            bool condition,
            string caller,
            string message)
        {
            Debug.Assert(
                condition,
                string.Format("[{0}] {1}", caller, message));
        }

        /// <summary>
        /// Gets the namespace prefix from an XmlNamespace enum value.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <returns>Namespace prefix string.</returns>
        internal static string GetNamespacePrefix(XmlNamespace xmlNamespace)
        {
            switch (xmlNamespace)
            {
                case XmlNamespace.Types:
                    return EwsTypesNamespacePrefix;
                case XmlNamespace.Messages:
                    return EwsMessagesNamespacePrefix;
                case XmlNamespace.Errors:
                    return EwsErrorsNamespacePrefix;
                case XmlNamespace.Soap:
                case XmlNamespace.Soap12:
                    return EwsSoapNamespacePrefix;
                case XmlNamespace.XmlSchemaInstance:
                    return EwsXmlSchemaInstanceNamespacePrefix;
                case XmlNamespace.PassportSoapFault:
                    return PassportSoapFaultNamespacePrefix;
                case XmlNamespace.WSTrustFebruary2005:
                    return WSTrustFebruary2005NamespacePrefix;
                case XmlNamespace.WSAddressing:
                    return WSAddressingNamespacePrefix;
                case XmlNamespace.Autodiscover:
                    return AutodiscoverSoapNamespacePrefix;
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// Gets the namespace URI from an XmlNamespace enum value.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <returns>Uri as string</returns>
        internal static string GetNamespaceUri(XmlNamespace xmlNamespace)
        {
            switch (xmlNamespace)
            {
                case XmlNamespace.Types:
                    return EwsTypesNamespace;
                case XmlNamespace.Messages:
                    return EwsMessagesNamespace;
                case XmlNamespace.Errors:
                    return EwsErrorsNamespace;
                case XmlNamespace.Soap:
                    return EwsSoapNamespace;
                case XmlNamespace.Soap12:
                    return EwsSoap12Namespace;
                case XmlNamespace.XmlSchemaInstance:
                    return EwsXmlSchemaInstanceNamespace;
                case XmlNamespace.PassportSoapFault:
                    return PassportSoapFaultNamespace;
                case XmlNamespace.WSTrustFebruary2005:
                    return WSTrustFebruary2005Namespace;
                case XmlNamespace.WSAddressing:
                    return WSAddressingNamespace;
                case XmlNamespace.Autodiscover:
                    return AutodiscoverSoapNamespace;
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// Gets the XmlNamespace enum value from a namespace Uri.
        /// </summary>
        /// <param name="namespaceUri">XML namespace Uri.</param>
        /// <returns>XmlNamespace enum value.</returns>
        internal static XmlNamespace GetNamespaceFromUri(string namespaceUri)
        {
            switch (namespaceUri)
            {
                case EwsErrorsNamespace:
                    return XmlNamespace.Errors;
                case EwsTypesNamespace:
                    return XmlNamespace.Types;
                case EwsMessagesNamespace:
                    return XmlNamespace.Messages;
                case EwsSoapNamespace:
                    return XmlNamespace.Soap;
                case EwsSoap12Namespace:
                    return XmlNamespace.Soap12;
                case EwsXmlSchemaInstanceNamespace:
                    return XmlNamespace.XmlSchemaInstance;
                case PassportSoapFaultNamespace:
                    return XmlNamespace.PassportSoapFault;
                case WSTrustFebruary2005Namespace:
                    return XmlNamespace.WSTrustFebruary2005;
                case WSAddressingNamespace:
                    return XmlNamespace.WSAddressing;
                default:
                    return XmlNamespace.NotSpecified;
            }
        }

        /// <summary>
        /// Creates EWS object based on XML element name.
        /// </summary>
        /// <typeparam name="TServiceObject">The type of the service object.</typeparam>
        /// <param name="service">The service.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>Service object.</returns>
        internal static TServiceObject CreateEwsObjectFromXmlElementName<TServiceObject>(ExchangeService service, string xmlElementName)
            where TServiceObject : ServiceObject
        {
            Type itemClass;

            if (EwsUtilities.serviceObjectInfo.Member.XmlElementNameToServiceObjectClassMap.TryGetValue(xmlElementName, out itemClass))
            {
                CreateServiceObjectWithServiceParam creationDelegate;

                if (EwsUtilities.serviceObjectInfo.Member.ServiceObjectConstructorsWithServiceParam.TryGetValue(itemClass, out creationDelegate))
                {
                    return (TServiceObject)creationDelegate(service);
                }
                else
                {
                    throw new ArgumentException(Strings.NoAppropriateConstructorForItemClass);
                }
            }
            else
            {
                return default(TServiceObject);
            }
        }

        /// <summary>
        /// Creates Item from Item class.
        /// </summary>
        /// <param name="itemAttachment">The item attachment.</param>
        /// <param name="itemClass">The item class.</param>
        /// <param name="isNew">If true, item attachment is new.</param>
        /// <returns>New Item.</returns>
        internal static Item CreateItemFromItemClass(
            ItemAttachment itemAttachment,
            Type itemClass,
            bool isNew)
        {
            CreateServiceObjectWithAttachmentParam creationDelegate;

            if (EwsUtilities.serviceObjectInfo.Member.ServiceObjectConstructorsWithAttachmentParam.TryGetValue(itemClass, out creationDelegate))
            {
                return (Item)creationDelegate(itemAttachment, isNew);
            }
            else
            {
                throw new ArgumentException(Strings.NoAppropriateConstructorForItemClass);
            }
        }

        /// <summary>
        /// Creates Item based on XML element name.
        /// </summary>
        /// <param name="itemAttachment">The item attachment.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>New Item.</returns>
        internal static Item CreateItemFromXmlElementName(ItemAttachment itemAttachment, string xmlElementName)
        {
            Type itemClass;

            if (EwsUtilities.serviceObjectInfo.Member.XmlElementNameToServiceObjectClassMap.TryGetValue(xmlElementName, out itemClass))
            {
                return CreateItemFromItemClass(itemAttachment, itemClass, false);
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Gets the expected item type based on the local name.
        /// </summary>
        /// <param name="xmlElementName"></param>
        /// <returns></returns>
        internal static Type GetItemTypeFromXmlElementName(string xmlElementName)
        {
            Type itemClass = null;
            EwsUtilities.serviceObjectInfo.Member.XmlElementNameToServiceObjectClassMap.TryGetValue(xmlElementName, out itemClass);
            return itemClass;
        }

        /// <summary>
        /// Finds the first item of type TItem (not a descendant type) in the specified collection.
        /// </summary>
        /// <typeparam name="TItem">The type of the item to find.</typeparam>
        /// <param name="items">The collection.</param>
        /// <returns>A TItem instance or null if no instance of TItem could be found.</returns>
        internal static TItem FindFirstItemOfType<TItem>(IEnumerable<Item> items)
            where TItem : Item
        {
            Type itemType = typeof(TItem);

            foreach (Item item in items)
            {
                // We're looking for an exact class match here.
                if (item.GetType() == itemType)
                {
                    return (TItem)item;
                }
            }

            return null;
        }

        #region Tracing routines

        /// <summary>
        /// Write trace start element.
        /// </summary>
        /// <param name="writer">The writer to write the start element to.</param>
        /// <param name="traceTag">The trace tag.</param>
        /// <param name="includeVersion">If true, include build version attribute.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Exchange.Usage", "EX0009:DoNotUseDateTimeNowOrFromFileTime", Justification = "Client API")]
        private static void WriteTraceStartElement(
            XmlWriter writer,
            string traceTag,
            bool includeVersion)
        {
            writer.WriteStartElement("Trace");
            writer.WriteAttributeString("Tag", traceTag);
            writer.WriteAttributeString("Tid", Thread.CurrentThread.ManagedThreadId.ToString());
            writer.WriteAttributeString("Time", DateTime.UtcNow.ToString("u", DateTimeFormatInfo.InvariantInfo));

            if (includeVersion)
            {
                writer.WriteAttributeString("Version", EwsUtilities.BuildVersion);
            }
        }

        /// <summary>
        /// Format log message.
        /// </summary>
        /// <param name="entryKind">Kind of the entry.</param>
        /// <param name="logEntry">The log entry.</param>
        /// <returns>XML log entry as a string.</returns>
        internal static string FormatLogMessage(string entryKind, string logEntry)
        {
            StringBuilder sb = new StringBuilder();
            using (StringWriter writer = new StringWriter(sb))
            {
                using (XmlTextWriter xmlWriter = new XmlTextWriter(writer))
                {
                    xmlWriter.Formatting = Formatting.Indented;

                    EwsUtilities.WriteTraceStartElement(xmlWriter, entryKind, false);

                    xmlWriter.WriteWhitespace(Environment.NewLine);
                    xmlWriter.WriteValue(logEntry);
                    xmlWriter.WriteWhitespace(Environment.NewLine);

                    xmlWriter.WriteEndElement(); // Trace
                    xmlWriter.WriteWhitespace(Environment.NewLine);
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Format the HTTP headers.
        /// </summary>
        /// <param name="sb">StringBuilder.</param>
        /// <param name="headers">The HTTP headers.</param>
        private static void FormatHttpHeaders(StringBuilder sb, WebHeaderCollection headers)
        {
            foreach (string key in headers.Keys)
            {
                sb.Append(
                    string.Format(
                        "{0}: {1}\n",
                        key,
                        headers[key]));
            }
        }

        /// <summary>
        /// Format request HTTP headers.
        /// </summary>
        /// <param name="request">The HTTP request.</param>
        internal static string FormatHttpRequestHeaders(IEwsHttpWebRequest request)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(string.Format("{0} {1} HTTP/1.1\n", request.Method, request.RequestUri.AbsolutePath));
            EwsUtilities.FormatHttpHeaders(sb, request.Headers);
            sb.Append("\n");

            return sb.ToString();
        }

        /// <summary>
        /// Format response HTTP headers.
        /// </summary>
        /// <param name="response">The HTTP response.</param>
        internal static string FormatHttpResponseHeaders(IEwsHttpWebResponse response)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(
                string.Format(
                    "HTTP/{0} {1} {2}\n",
                    response.ProtocolVersion,
                    (int)response.StatusCode,
                    response.StatusDescription));

            sb.Append(EwsUtilities.FormatHttpHeaders(response.Headers));
            sb.Append("\n");
            return sb.ToString();
        }

        /// <summary>
        /// Format request HTTP headers.
        /// </summary>
        /// <param name="request">The HTTP request.</param>
        internal static string FormatHttpRequestHeaders(HttpWebRequest request)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(
                string.Format(
                    "{0} {1} HTTP/{2}\n",
                    request.Method.ToUpperInvariant(),
                    request.RequestUri.AbsolutePath,
                    request.ProtocolVersion));

            sb.Append(EwsUtilities.FormatHttpHeaders(request.Headers));
            sb.Append("\n");
            return sb.ToString();
        }

        /// <summary>
        /// Formats HTTP headers.
        /// </summary>
        /// <param name="headers">The headers.</param>
        /// <returns>Headers as a string</returns>
        private static string FormatHttpHeaders(WebHeaderCollection headers)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string key in headers.Keys)
            {
                sb.Append(
                    string.Format(
                        "{0}: {1}\n",
                        key,
                        headers[key]));
            }
            return sb.ToString();
        }

        /// <summary>
        /// Format XML content in a MemoryStream for message.
        /// </summary>
        /// <param name="entryKind">Kind of the entry.</param>
        /// <param name="memoryStream">The memory stream.</param>
        /// <returns>XML log entry as a string.</returns>
        internal static string FormatLogMessageWithXmlContent(string entryKind, MemoryStream memoryStream)
        {
            StringBuilder sb = new StringBuilder();
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.ConformanceLevel = ConformanceLevel.Fragment;
            settings.IgnoreComments = true;
            settings.IgnoreWhitespace = true;
            settings.CloseInput = false;

            // Remember the current location in the MemoryStream.
            long lastPosition = memoryStream.Position;

            // Rewind the position since we want to format the entire contents.
            memoryStream.Position = 0;

            try
            {
                using (XmlReader reader = XmlReader.Create(memoryStream, settings))
                {
                    using (StringWriter writer = new StringWriter(sb))
                    {
                        using (XmlTextWriter xmlWriter = new XmlTextWriter(writer))
                        {
                            xmlWriter.Formatting = Formatting.Indented;

                            EwsUtilities.WriteTraceStartElement(xmlWriter, entryKind, true);

                            while (!reader.EOF)
                            {
                                xmlWriter.WriteNode(reader, true);
                            }

                            xmlWriter.WriteEndElement(); // Trace
                            xmlWriter.WriteWhitespace(Environment.NewLine);
                        }
                    }
                }
            }
            catch (XmlException)
            {
                // We tried to format the content as "pretty" XML. Apparently the content is
                // not well-formed XML or isn't XML at all. Fallback and treat it as plain text.
                sb.Length = 0;
                memoryStream.Position = 0;
                sb.Append(Encoding.UTF8.GetString(memoryStream.GetBuffer(), 0, (int)memoryStream.Length));
            }

            // Restore Position in the stream.
            memoryStream.Position = lastPosition;

            return sb.ToString();
        }

        #endregion

        #region Stream routines

        /// <summary>
        /// Copies source stream to target.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="target">The target.</param>
        internal static void CopyStream(Stream source, Stream target)
        {
            // See if this is a MemoryStream -- we can use WriteTo.
            MemoryStream memContentStream = source as MemoryStream;
            if (memContentStream != null)
            {
                memContentStream.WriteTo(target);
            }
            else
            {
                // Otherwise, copy data through a buffer
                byte[] buffer = new byte[4096];
                int bufferSize = buffer.Length;
                int bytesRead = source.Read(buffer, 0, bufferSize);
                while (bytesRead > 0)
                {
                    target.Write(buffer, 0, bytesRead);
                    bytesRead = source.Read(buffer, 0, bufferSize);
                }
            }
        }

        #endregion

        /// <summary>
        /// Gets the build version.
        /// </summary>
        /// <value>The build version.</value>
        internal static string BuildVersion
        {
            get { return EwsUtilities.buildVersion.Member; }
        }

        #region Conversion routines

        /// <summary>
        /// Convert bool to XML Schema bool.
        /// </summary>
        /// <param name="value">Bool value.</param>
        /// <returns>String representing bool value in XML Schema.</returns>
        internal static string BoolToXSBool(bool value)
        {
            return value ? EwsUtilities.XSTrue : EwsUtilities.XSFalse;
        }

        /// <summary>
        /// Parses an enum value list.
        /// </summary>
        /// <typeparam name="T">Type of value.</typeparam>
        /// <param name="list">The list.</param>
        /// <param name="value">The value.</param>
        /// <param name="separators">The separators.</param>
        internal static void ParseEnumValueList<T>(
            IList<T> list,
            string value,
            params char[] separators)
            where T : struct
        {
            EwsUtilities.Assert(
                typeof(T).IsEnum,
                "EwsUtilities.ParseEnumValueList",
                "T is not an enum type.");

            if (string.IsNullOrEmpty(value))
            {
                return;
            }

            string[] enumValues = value.Split(separators);

            foreach (string enumValue in enumValues)
            {
                list.Add((T)Enum.Parse(typeof(T), enumValue, false));
            }
        }

        /// <summary>
        /// Converts an enum to a string, using the mapping dictionaries if appropriate.
        /// </summary>
        /// <param name="value">The enum value to be serialized</param>
        /// <returns>String representation of enum to be used in the protocol</returns>
        internal static string SerializeEnum(Enum value)
        {
            Dictionary<Enum, string> enumToStringDict;
            string strValue;
            if (enumToSchemaDictionaries.Member.TryGetValue(value.GetType(), out enumToStringDict) &&
                enumToStringDict.TryGetValue(value, out strValue))
            {
                return strValue;
            }
            else
            {
                return value.ToString();
            }
        }

        /// <summary>
        /// Parses specified value based on type.
        /// </summary>
        /// <typeparam name="T">Type of value.</typeparam>
        /// <param name="value">The value.</param>
        /// <returns>Value of type T.</returns>
        internal static T Parse<T>(string value)
        {
            if (typeof(T).IsEnum)
            {
                Dictionary<string, Enum> stringToEnumDict;
                Enum enumValue;
                if (schemaToEnumDictionaries.Member.TryGetValue(typeof(T), out stringToEnumDict) &&
                    stringToEnumDict.TryGetValue(value, out enumValue))
                {
                    // This double-casting is ugly, but necessary. By this point, we know that T is an Enum
                    // (same as returned by the dictionary), but the compiler can't prove it. Thus, the 
                    // up-cast before we can down-cast.
                    return (T)((object)enumValue);
                }
                else
                {
                    return (T)Enum.Parse(typeof(T), value, false);
                }
            }
            else
            {
                return (T)Convert.ChangeType(value, typeof(T), CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        /// Tries to parses the specified value to the specified type.
        /// </summary>
        /// <typeparam name="T">The type into which to cast the provided value.</typeparam>
        /// <param name="value">The value to parse.</param>
        /// <param name="result">The value cast to the specified type, if TryParse succeeds. Otherwise, the value of result is indeterminate.</param>
        /// <returns>True if value could be parsed; otherwise, false.</returns>
        internal static bool TryParse<T>(string value, out T result)
        {
            try
            {
                result = EwsUtilities.Parse<T>(value);

                return true;
            }
            //// Catch all exceptions here, we're not interested in the reason why TryParse failed.
            catch (Exception)
            {
                result = default(T);

                return false;
            }
        }

        /// <summary>
        /// Converts the specified date and time from one time zone to another.
        /// </summary>
        /// <param name="dateTime">The date time to convert.</param>
        /// <param name="sourceTimeZone">The source time zone.</param>
        /// <param name="destinationTimeZone">The destination time zone.</param>
        /// <returns>A DateTime that holds the converted</returns>
        internal static DateTime ConvertTime(
            DateTime dateTime,
            TimeZoneInfo sourceTimeZone,
            TimeZoneInfo destinationTimeZone)
        {
            try
            {
                return TimeZoneInfo.ConvertTime(
                    dateTime,
                    sourceTimeZone,
                    destinationTimeZone);
            }
            catch (ArgumentException e)
            {
                throw new TimeZoneConversionException(
                    string.Format(
                        Strings.CannotConvertBetweenTimeZones,
                        EwsUtilities.DateTimeToXSDateTime(dateTime),
                        sourceTimeZone.DisplayName,
                        destinationTimeZone.DisplayName),
                    e);
            }
        }

        /// <summary>
        /// Reads the string as date time, assuming it is unbiased (e.g. 2009/01/01T08:00)
        /// and scoped to service's time zone.
        /// </summary>
        /// <param name="dateString">The date string.</param>
        /// <param name="service">The service.</param>
        /// <returns>The string's value as a DateTime object.</returns>
        internal static DateTime ParseAsUnbiasedDatetimescopedToServicetimeZone(string dateString, ExchangeService service)
        {
            // Convert the element's value to a DateTime with no adjustment.
            DateTime tempDate = DateTime.Parse(dateString, CultureInfo.InvariantCulture);

            // Set the kind according to the service's time zone
            if (service.TimeZone == TimeZoneInfo.Utc)
            {
                return new DateTime(tempDate.Ticks, DateTimeKind.Utc);
            }
            else if (EwsUtilities.IsLocalTimeZone(service.TimeZone))
            {
                return new DateTime(tempDate.Ticks, DateTimeKind.Local);
            }
            else
            {
                return new DateTime(tempDate.Ticks, DateTimeKind.Unspecified);
            }
        }

        /// <summary>
        /// Determines whether the specified time zone is the same as the system's local time zone.
        /// </summary>
        /// <param name="timeZone">The time zone to check.</param>
        /// <returns>
        ///     <c>true</c> if the specified time zone is the same as the system's local time zone; otherwise, <c>false</c>.
        /// </returns>
        internal static bool IsLocalTimeZone(TimeZoneInfo timeZone)
        {
            return (TimeZoneInfo.Local == timeZone) || (TimeZoneInfo.Local.Id == timeZone.Id && TimeZoneInfo.Local.HasSameRules(timeZone));
        }

        /// <summary>
        /// Convert DateTime to XML Schema date.
        /// </summary>
        /// <param name="date">The date to be converted.</param>
        /// <returns>String representation of DateTime.</returns>
        internal static string DateTimeToXSDate(DateTime date)
        {
            // Depending on the current culture, DateTime formatter will 
            // translate dates from one culture to another (e.g. Gregorian to Lunar).  The server
            // however, considers all dates to be in Gregorian, so using the InvariantCulture will
            // ensure this.
            string format;

            switch (date.Kind)
            {
                case DateTimeKind.Utc:
                    format = "yyyy-MM-ddZ";
                    break;
                case DateTimeKind.Unspecified:
                    format = "yyyy-MM-dd";
                    break;
                default: // DateTimeKind.Local is remaining
                    format = "yyyy-MM-ddzzz";
                    break;
            }

            return date.ToString(format, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Dates the DateTime into an XML schema date time.
        /// </summary>
        /// <param name="dateTime">The date time.</param>
        /// <returns>String representation of DateTime.</returns>
        internal static string DateTimeToXSDateTime(DateTime dateTime)
        {
            string format = "yyyy-MM-ddTHH:mm:ss.fff";

            switch (dateTime.Kind)
            {
                case DateTimeKind.Utc:
                    format += "Z";
                    break;
                case DateTimeKind.Local:
                    format += "zzz";
                    break;
                default:
                    break;
            }

            // Depending on the current culture, DateTime formatter will replace ':' with 
            // the DateTimeFormatInfo.TimeSeparator property which may not be ':'. Force the proper string
            // to be used by using the InvariantCulture.
            return dateTime.ToString(format, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Convert EWS DayOfTheWeek enum to System.DayOfWeek.
        /// </summary>
        /// <param name="dayOfTheWeek">The day of the week.</param>
        /// <returns>System.DayOfWeek value.</returns>
        internal static DayOfWeek EwsToSystemDayOfWeek(DayOfTheWeek dayOfTheWeek)
        {
            if (dayOfTheWeek == DayOfTheWeek.Day ||
                dayOfTheWeek == DayOfTheWeek.Weekday ||
                dayOfTheWeek == DayOfTheWeek.WeekendDay)
            {
                throw new ArgumentException(
                    string.Format("Cannot convert {0} to System.DayOfWeek enum value", dayOfTheWeek),
                    "dayOfTheWeek");
            }
            else
            {
                return (DayOfWeek)dayOfTheWeek;
            }
        }

        /// <summary>
        /// Convert System.DayOfWeek type to EWS DayOfTheWeek.
        /// </summary>
        /// <param name="dayOfWeek">The dayOfWeek.</param>
        /// <returns>EWS DayOfWeek value</returns>
        internal static DayOfTheWeek SystemToEwsDayOfTheWeek(DayOfWeek dayOfWeek)
        {
            return (DayOfTheWeek)dayOfWeek;
        }

        /// <summary>
        /// Takes a System.TimeSpan structure and converts it into an 
        /// xs:duration string as defined by the W3 Consortiums Recommendation
        /// "XML Schema Part 2: Datatypes Second Edition", 
        /// http://www.w3.org/TR/xmlschema-2/#duration
        /// </summary>
        /// <param name="timeSpan">TimeSpan structure to convert</param>
        /// <returns>xs:duration formatted string</returns>
        internal static string TimeSpanToXSDuration(TimeSpan timeSpan)
        {
            // Optional '-' offset
            string offsetStr = (timeSpan.TotalSeconds < 0) ? "-" : string.Empty;

            // The TimeSpan structure does not have a Year or Month 
            // property, therefore we wouldn't be able to return an xs:duration
            // string from a TimeSpan that included the nY or nM components.
            return String.Format(
                "{0}P{1}DT{2}H{3}M{4}S",
                offsetStr,
                Math.Abs(timeSpan.Days),
                Math.Abs(timeSpan.Hours),
                Math.Abs(timeSpan.Minutes),
                Math.Abs(timeSpan.Seconds) + "." + Math.Abs(timeSpan.Milliseconds));
        }

        /// <summary>
        /// Takes an xs:duration string as defined by the W3 Consortiums
        /// Recommendation "XML Schema Part 2: Datatypes Second Edition", 
        /// http://www.w3.org/TR/xmlschema-2/#duration, and converts it
        /// into a System.TimeSpan structure
        /// </summary>
        /// <remarks>
        /// This method uses the following approximations:
        ///     1 year = 365 days
        ///     1 month = 30 days
        /// Additionally, it only allows for four decimal points of
        /// seconds precision.
        /// </remarks>
        /// <param name="xsDuration">xs:duration string to convert</param>
        /// <returns>System.TimeSpan structure</returns>
        internal static TimeSpan XSDurationToTimeSpan(string xsDuration)
        {
            Regex timeSpanParser = new Regex(
                "(?<pos>-)?" +
                "P" +
                "((?<year>[0-9]+)Y)?" +
                "((?<month>[0-9]+)M)?" +
                "((?<day>[0-9]+)D)?" +
                "(T" +
                "((?<hour>[0-9]+)H)?" +
                "((?<minute>[0-9]+)M)?" +
                "((?<seconds>[0-9]+)(\\.(?<precision>[0-9]+))?S)?)?");

            Match m = timeSpanParser.Match(xsDuration);
            if (!m.Success)
            {
                throw new ArgumentException(Strings.XsDurationCouldNotBeParsed);
            }
            string token = m.Result("${pos}");
            bool negative = false;
            if (!String.IsNullOrEmpty(token))
            {
                negative = true;
            }

            // Year
            token = m.Result("${year}");
            int year = 0;
            if (!String.IsNullOrEmpty(token))
            {
                year = Int32.Parse(token);
            }

            // Month
            token = m.Result("${month}");
            int month = 0;
            if (!String.IsNullOrEmpty(token))
            {
                month = Int32.Parse(token);
            }

            // Day
            token = m.Result("${day}");
            int day = 0;
            if (!String.IsNullOrEmpty(token))
            {
                day = Int32.Parse(token);
            }

            // Hour
            token = m.Result("${hour}");
            int hour = 0;
            if (!String.IsNullOrEmpty(token))
            {
                hour = Int32.Parse(token);
            }

            // Minute
            token = m.Result("${minute}");
            int minute = 0;
            if (!String.IsNullOrEmpty(token))
            {
                minute = Int32.Parse(token);
            }

            // Seconds
            token = m.Result("${seconds}");
            int seconds = 0;
            if (!String.IsNullOrEmpty(token))
            {
                seconds = Int32.Parse(token);
            }

            int milliseconds = 0;
            token = m.Result("${precision}");

            // Only allowed 4 digits of precision
            if (token.Length > 4)
            {
                token = token.Substring(0, 4);
            }

            if (!String.IsNullOrEmpty(token))
            {
                milliseconds = Int32.Parse(token);
            }

            // Apply conversions of year and months to days.
            // Year = 365 days
            // Month = 30 days
            day = day + (year * 365) + (month * 30);
            TimeSpan retval = new TimeSpan(day, hour, minute, seconds, milliseconds);

            if (negative)
            {
                retval = -retval;
            }

            return retval;
        }

        /// <summary>
        /// Converts the specified time span to its XSD representation.
        /// </summary>
        /// <param name="timeSpan">The time span.</param>
        /// <returns>The XSD representation of the specified time span.</returns>
        public static string TimeSpanToXSTime(TimeSpan timeSpan)
        {
            return string.Format(
                "{0:00}:{1:00}:{2:00}",
                timeSpan.Hours,
                timeSpan.Minutes,
                timeSpan.Seconds);
        }

        #endregion

        #region Type Name utilities
        /// <summary>
        /// Gets the printable name of a CLR type.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <returns>Printable name.</returns>
        public static string GetPrintableTypeName(Type type)
        {
            if (type.IsGenericType)
            {
                // Convert generic type to printable form (e.g. List<Item>)
                string genericPrefix = type.Name.Substring(0, type.Name.IndexOf('`'));
                StringBuilder nameBuilder = new StringBuilder(genericPrefix);

                // Note: building array of generic parameters is done recursively. Each parameter could be any type.
                string[] genericArgs = type.GetGenericArguments().ToList<Type>().ConvertAll<string>(t => GetPrintableTypeName(t)).ToArray<string>();

                nameBuilder.Append("<");
                nameBuilder.Append(string.Join(",", genericArgs));
                nameBuilder.Append(">");
                return nameBuilder.ToString();
            }
            else if (type.IsArray)
            {
                // Convert array type to printable form.
                string arrayPrefix = type.Name.Substring(0, type.Name.IndexOf('['));
                StringBuilder nameBuilder = new StringBuilder(EwsUtilities.GetSimplifiedTypeName(arrayPrefix));
                for (int rank = 0; rank < type.GetArrayRank(); rank++)
                {
                    nameBuilder.Append("[]");
                }
                return nameBuilder.ToString();
            }
            else
            {
                return EwsUtilities.GetSimplifiedTypeName(type.Name);
            }
        }

        /// <summary>
        /// Gets the printable name of a simple CLR type.
        /// </summary>
        /// <param name="typeName">The type name.</param>
        /// <returns>Printable name.</returns>
        private static string GetSimplifiedTypeName(string typeName)
        {
            // If type has a shortname (e.g. int for Int32) map to the short name.
            string name;
            return typeNameToShortNameMap.Member.TryGetValue(typeName, out name) ? name : typeName;
        }

        #endregion

        #region EmailAddress parsing

        /// <summary>
        /// Gets the domain name from an email address.
        /// </summary>
        /// <param name="emailAddress">The email address.</param>
        /// <returns>Domain name.</returns>
        internal static string DomainFromEmailAddress(string emailAddress)
        {
            string[] emailAddressParts = emailAddress.Split('@');

            if (emailAddressParts.Length != 2 || string.IsNullOrEmpty(emailAddressParts[1]))
            {
                throw new FormatException(Strings.InvalidEmailAddress);
            }

            return emailAddressParts[1];
        }

        #endregion

        #region Method parameters validation routines

        /// <summary>
        /// Validates parameter (and allows null value).
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="paramName">Name of the param.</param>
        internal static void ValidateParamAllowNull(object param, string paramName)
        {
            ISelfValidate selfValidate = param as ISelfValidate;

            if (selfValidate != null)
            {
                try
                {
                    selfValidate.Validate();
                }
                catch (ServiceValidationException e)
                {
                    throw new ArgumentException(
                        Strings.ValidationFailed,
                        paramName,
                        e);
                }
            }

            ServiceObject ewsObject = param as ServiceObject;

            if (ewsObject != null)
            {
                if (ewsObject.IsNew)
                {
                    throw new ArgumentException(Strings.ObjectDoesNotHaveId, paramName);
                }
            }
        }

        /// <summary>
        /// Validates parameter (null value not allowed).
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="paramName">Name of the param.</param>
        internal static void ValidateParam(object param, string paramName)
        {
            bool isValid;

            string strParam = param as string;
            if (strParam != null)
            {
                isValid = !string.IsNullOrEmpty(strParam);
            }
            else
            {
                isValid = param != null;
            }

            if (!isValid)
            {
                throw new ArgumentNullException(paramName);
            }

            ValidateParamAllowNull(param, paramName);
        }

        /// <summary>
        /// Validates parameter collection.
        /// </summary>
        /// <param name="collection">The collection.</param>
        /// <param name="paramName">Name of the param.</param>
        internal static void ValidateParamCollection(IEnumerable collection, string paramName)
        {
            ValidateParam(collection, paramName);

            int count = 0;

            foreach (object obj in collection)
            {
                try
                {
                    ValidateParam(obj, string.Format("collection[{0}]", count));
                }
                catch (ArgumentException e)
                {
                    throw new ArgumentException(
                        string.Format("The element at position {0} is invalid", count),
                        paramName,
                        e);
                }

                count++;
            }

            if (count == 0)
            {
                throw new ArgumentException(Strings.CollectionIsEmpty, paramName);
            }
        }

        /// <summary>
        /// Validates string parameter to be non-empty string (null value allowed).
        /// </summary>
        /// <param name="param">The string parameter.</param>
        /// <param name="paramName">Name of the parameter.</param>
        internal static void ValidateNonBlankStringParamAllowNull(string param, string paramName)
        {
            if (param != null)
            {
                // Non-empty string has at least one character which is *not* a whitespace character
                if (param.Length == param.CountMatchingChars((c) => Char.IsWhiteSpace(c)))
                {
                    throw new ArgumentException(Strings.ArgumentIsBlankString, paramName);
                }
            }
        }

        /// <summary>
        /// Validates string parameter to be non-empty string (null value not allowed).
        /// </summary>
        /// <param name="param">The string parameter.</param>
        /// <param name="paramName">Name of the parameter.</param>
        internal static void ValidateNonBlankStringParam(string param, string paramName)
        {
            if (param == null)
            {
                throw new ArgumentNullException(paramName);
            }

            ValidateNonBlankStringParamAllowNull(param, paramName);
        }

        /// <summary>
        /// Validates the enum value against the request version.
        /// </summary>
        /// <param name="enumValue">The enum value.</param>
        /// <param name="requestVersion">The request version.</param>
        /// <exception cref="ServiceVersionException">Raised if this enum value requires a later version of Exchange.</exception>
        internal static void ValidateEnumVersionValue(Enum enumValue, ExchangeVersion requestVersion)
        {
            Type enumType = enumValue.GetType();
            Dictionary<Enum, ExchangeVersion> enumVersionDict = enumVersionDictionaries.Member[enumType];
            ExchangeVersion enumVersion = enumVersionDict[enumValue];
            if (requestVersion < enumVersion)
            {
                throw new ServiceVersionException(
                    string.Format(
                                  Strings.EnumValueIncompatibleWithRequestVersion,
                                  enumValue.ToString(),
                                  enumType.Name,
                                  enumVersion));
            }
        }

        /// <summary>
        /// Validates service object version against the request version.
        /// </summary>
        /// <param name="serviceObject">The service object.</param>
        /// <param name="requestVersion">The request version.</param>
        /// <exception cref="ServiceVersionException">Raised if this service object type requires a later version of Exchange.</exception>
        internal static void ValidateServiceObjectVersion(ServiceObject serviceObject, ExchangeVersion requestVersion)
        {
            ExchangeVersion minimumRequiredServerVersion = serviceObject.GetMinimumRequiredServerVersion();

            if (requestVersion < minimumRequiredServerVersion)
            {
                throw new ServiceVersionException(
                    string.Format(
                    Strings.ObjectTypeIncompatibleWithRequestVersion,
                    serviceObject.GetType().Name,
                    minimumRequiredServerVersion));
            }
        }

        /// <summary>
        /// Validates property version against the request version.
        /// </summary>
        /// <param name="service">The Exchange service.</param>
        /// <param name="minimumServerVersion">The minimum server version that supports the property.</param>
        /// <param name="propertyName">Name of the property.</param>
        internal static void ValidatePropertyVersion(
            ExchangeService service,
            ExchangeVersion minimumServerVersion,
            string propertyName)
        {
            if (service.RequestedServerVersion < minimumServerVersion)
            {
                throw new ServiceVersionException(
                    string.Format(
                    Strings.PropertyIncompatibleWithRequestVersion,
                    propertyName,
                    minimumServerVersion));
            }
        }

        /// <summary>
        /// Validates method version against the request version.
        /// </summary>
        /// <param name="service">The Exchange service.</param>
        /// <param name="minimumServerVersion">The minimum server version that supports the method.</param>
        /// <param name="methodName">Name of the method.</param>
        internal static void ValidateMethodVersion(
            ExchangeService service,
            ExchangeVersion minimumServerVersion,
            string methodName)
        {
            if (service.RequestedServerVersion < minimumServerVersion)
            {
                throw new ServiceVersionException(
                    string.Format(
                    Strings.MethodIncompatibleWithRequestVersion,
                    methodName,
                    minimumServerVersion));
            }
        }

        /// <summary>
        /// Validates class version against the request version.
        /// </summary>
        /// <param name="service">The Exchange service.</param>
        /// <param name="minimumServerVersion">The minimum server version that supports the method.</param>
        /// <param name="className">Name of the class.</param>
        internal static void ValidateClassVersion(
            ExchangeService service,
            ExchangeVersion minimumServerVersion,
            string className)
        {
            if (service.RequestedServerVersion < minimumServerVersion)
            {
                throw new ServiceVersionException(
                    string.Format(
                    Strings.ClassIncompatibleWithRequestVersion,
                    className,
                    minimumServerVersion));
            }
        }

        /// <summary>
        /// Validates domain name (null value allowed)
        /// </summary>
        /// <param name="domainName">Domain name.</param>
        /// <param name="paramName">Parameter name.</param>
        internal static void ValidateDomainNameAllowNull(string domainName, string paramName)
        {
            if (domainName != null)
            {
                Regex regex = new Regex(DomainRegex);

                if (!regex.IsMatch(domainName))
                {
                    throw new ArgumentException(string.Format(Strings.InvalidDomainName, domainName), paramName);
                }
            }
        }

        /// <summary>
        /// Gets version for enum member.
        /// </summary>
        /// <param name="enumType">Type of the enum.</param>
        /// <param name="enumName">The enum name.</param>
        /// <returns>Exchange version in which the enum value was first defined.</returns>
        private static ExchangeVersion GetEnumVersion(Type enumType, string enumName)
        {
            MemberInfo[] memberInfo = enumType.GetMember(enumName);
            EwsUtilities.Assert(
                                (memberInfo != null) && (memberInfo.Length > 0),
                                "EwsUtilities.GetEnumVersion",
                                "Enum member " + enumName + " not found in " + enumType);

            object[] attrs = memberInfo[0].GetCustomAttributes(typeof(RequiredServerVersionAttribute), false);
            if (attrs != null && attrs.Length > 0)
            {
                return ((RequiredServerVersionAttribute)attrs[0]).Version;
            }
            else
            {
                return ExchangeVersion.Exchange2007_SP1;
            }
        }

        /// <summary>
        /// Builds the enum to version mapping dictionary.
        /// </summary>
        /// <param name="enumType">Type of the enum.</param>
        /// <returns>Dictionary of enum values to versions.</returns>
        private static Dictionary<Enum, ExchangeVersion> BuildEnumDict(Type enumType)
        {
            Dictionary<Enum, ExchangeVersion> dict = new Dictionary<Enum, ExchangeVersion>();
            string[] names = Enum.GetNames(enumType);
            foreach (string name in names)
            {
                Enum value = (Enum)Enum.Parse(enumType, name, false);
                ExchangeVersion version = GetEnumVersion(enumType, name);
                dict.Add(value, version);
            }
            return dict;
        }

        /// <summary>
        /// Gets the schema name for enum member.
        /// </summary>
        /// <param name="enumType">Type of the enum.</param>
        /// <param name="enumName">The enum name.</param>
        /// <returns>The name for the enum used in the protocol, or null if it is the same as the enum's ToString().</returns>
        private static string GetEnumSchemaName(Type enumType, string enumName)
        {
            MemberInfo[] memberInfo = enumType.GetMember(enumName);
            EwsUtilities.Assert(
                                (memberInfo != null) && (memberInfo.Length > 0),
                                "EwsUtilities.GetEnumSchemaName",
                                "Enum member " + enumName + " not found in " + enumType);

            object[] attrs = memberInfo[0].GetCustomAttributes(typeof(EwsEnumAttribute), false);
            if (attrs != null && attrs.Length > 0)
            {
                return ((EwsEnumAttribute)attrs[0]).SchemaName;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Builds the schema to enum mapping dictionary.
        /// </summary>
        /// <param name="enumType">Type of the enum.</param>
        /// <returns>The mapping from enum to schema name</returns>
        private static Dictionary<string, Enum> BuildSchemaToEnumDict(Type enumType)
        {
            Dictionary<string, Enum> dict = new Dictionary<string, Enum>();
            string[] names = Enum.GetNames(enumType);
            foreach (string name in names)
            {
                Enum value = (Enum)Enum.Parse(enumType, name, false);
                string schemaName = EwsUtilities.GetEnumSchemaName(enumType, name);

                if (!String.IsNullOrEmpty(schemaName))
                {
                    dict.Add(schemaName, value);
                }
            }
            return dict;
        }

        /// <summary>
        /// Builds the enum to schema mapping dictionary.
        /// </summary>
        /// <param name="enumType">Type of the enum.</param>
        /// <returns>The mapping from enum to schema name</returns>
        private static Dictionary<Enum, string> BuildEnumToSchemaDict(Type enumType)
        {
            Dictionary<Enum, string> dict = new Dictionary<Enum, string>();
            string[] names = Enum.GetNames(enumType);
            foreach (string name in names)
            {
                Enum value = (Enum)Enum.Parse(enumType, name, false);
                string schemaName = EwsUtilities.GetEnumSchemaName(enumType, name);

                if (!String.IsNullOrEmpty(schemaName))
                {
                    dict.Add(value, schemaName);
                }
            }
            return dict;
        }
        #endregion

        #region IEnumerable utility methods

        /// <summary>
        /// Gets the enumerated object count.
        /// </summary>
        /// <param name="objects">The objects.</param>
        /// <returns>Count of objects in IEnumerable.</returns>
        internal static int GetEnumeratedObjectCount(IEnumerable objects)
        {
            int count = 0;

            foreach (object obj in objects)
            {
                count++;
            }

            return count;
        }

        /// <summary>
        /// Gets enumerated object at index.
        /// </summary>
        /// <param name="objects">The objects.</param>
        /// <param name="index">The index.</param>
        /// <returns>Object at index.</returns>
        internal static object GetEnumeratedObjectAt(IEnumerable objects, int index)
        {
            int count = 0;

            foreach (object obj in objects)
            {
                if (count == index)
                {
                    return obj;
                }

                count++;
            }

            throw new ArgumentOutOfRangeException("index", Strings.IEnumerableDoesNotContainThatManyObject);
        }

        #endregion

        #region Extension methods
        /// <summary>
        /// Count characters in string that match a condition.
        /// </summary>
        /// <param name="str">The string.</param>
        /// <param name="charPredicate">Predicate to evaluate for each character in the string.</param>
        /// <returns>Count of characters that match condition expressed by predicate.</returns>
        internal static int CountMatchingChars(this string str, Predicate<char> charPredicate)
        {
            int count = 0;
            foreach (char ch in str)
            {
                if (charPredicate(ch))
                {
                    count++;
                }
            }

            return count;
        }

        /// <summary>
        /// Determines whether every element in the collection matches the conditions defined by the specified predicate.
        /// </summary>
        /// <typeparam name="T">Entry type.</typeparam>
        /// <param name="collection">The collection.</param>
        /// <param name="predicate">Predicate that defines the conditions to check against the elements.</param>
        /// <returns>True if every element in the collection matches the conditions defined by the specified predicate; otherwise, false.</returns>
        internal static bool TrueForAll<T>(this IEnumerable<T> collection, Predicate<T> predicate)
        {
            foreach (T entry in collection)
            {
                if (!predicate(entry))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Call an action for each member of a collection.
        /// </summary>
        /// <param name="collection">The collection.</param>
        /// <param name="action">The action to apply.</param>
        /// <typeparam name="T">Collection element type.</typeparam>
        internal static void ForEach<T>(this IEnumerable<T> collection, Action<T> action)
        {
            foreach (T entry in collection)
            {
                action(entry);
            }
        }
        #endregion
    }
}