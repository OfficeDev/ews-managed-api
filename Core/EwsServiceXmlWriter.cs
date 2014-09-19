// ---------------------------------------------------------------------------
// <copyright file="EwsServiceXmlWriter.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the EwsServiceXmlWriter class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Globalization;
    using System.IO;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// XML writer
    /// </summary>
    internal class EwsServiceXmlWriter : IDisposable
    {
        /// <summary>
        /// Buffer size for writing Base64 encoded content.
        /// </summary>
        private const int BufferSize = 4096;

        /// <summary>
        /// UTF-8 encoding that does not create leading Byte order marks
        /// </summary>
        private static Encoding utf8Encoding = new UTF8Encoding(false);

        private bool isDisposed;
        private ExchangeServiceBase service;
        private XmlWriter xmlWriter;
        private bool isTimeZoneHeaderEmitted;
        private bool requireWSSecurityUtilityNamespace;

        /// <summary>
        /// Initializes a new instance of the <see cref="EwsServiceXmlWriter"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="stream">The stream.</param>
        internal EwsServiceXmlWriter(ExchangeServiceBase service, Stream stream)
        {
            this.service = service;

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;

            settings.Encoding = EwsServiceXmlWriter.utf8Encoding;

            this.xmlWriter = XmlWriter.Create(stream, settings);
        }

        /// <summary>
        /// Try to convert object to a string.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="strValue">The string representation of value.</param>
        /// <returns>True if object was converted, false otherwise.</returns>
        /// <remarks>A null object will be "successfully" converted to a null string.</remarks>
        internal bool TryConvertObjectToString(object value, out string strValue)
        {
            strValue = null;
            bool converted = true;

            if (value != null)
            {
                // All value types should implement IConvertible. There are a couple of special cases 
                // that need to be handled directly. Otherwise use IConvertible.ToString()
                IConvertible convertible = value as IConvertible;
                if (value.GetType().IsEnum)
                {
                    strValue = EwsUtilities.SerializeEnum((Enum)value);
                }
                else if (convertible != null)
                {
                    switch (convertible.GetTypeCode())
                    {
                        case TypeCode.Boolean:
                            strValue = EwsUtilities.BoolToXSBool((bool)value);
                            break;

                        case TypeCode.DateTime:
                            strValue = this.Service.ConvertDateTimeToUniversalDateTimeString((DateTime)value);
                            break;

                        default:
                            strValue = convertible.ToString(CultureInfo.InvariantCulture);
                            break;
                    }
                }
                else
                {
                    // If the value type doesn't implement IConvertible but implements IFormattable, use its
                    // ToString(format,formatProvider) method to convert to a string.
                    IFormattable formattable = value as IFormattable;
                    if (formattable != null)
                    {
                        // Null arguments mean that we use default format and default locale.
                        strValue = formattable.ToString(null, null);
                    }
                    else if (value is ISearchStringProvider)
                    {
                        // If the value type doesn't implement IConvertible or IFormattable but implements 
                        // ISearchStringProvider convert to a string.
                        // Note: if a value type implements IConvertible or IFormattable we will *not* check
                        // to see if it also implements ISearchStringProvider. We'll always use its IConvertible.ToString 
                        // or IFormattable.ToString method.
                        ISearchStringProvider searchStringProvider = value as ISearchStringProvider;
                        strValue = searchStringProvider.GetSearchString();                        
                    }
                    else if (value is byte[])
                    {
                        // Special case for byte arrays. Convert to Base64-encoded string.
                        strValue = Convert.ToBase64String((byte[])value);
                    }
                    else
                    {
                        converted = false;
                    }
                }
            }

            return converted;
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (!this.isDisposed)
            {
                this.xmlWriter.Close();

                this.isDisposed = true;
            }
        }

        /// <summary>
        /// Flushes this instance.
        /// </summary>
        public void Flush()
        {
            this.xmlWriter.Flush();
        }

        /// <summary>
        /// Writes the start element.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localName">The local name of the element.</param>
        public void WriteStartElement(XmlNamespace xmlNamespace, string localName)
        {
            this.xmlWriter.WriteStartElement(
                EwsUtilities.GetNamespacePrefix(xmlNamespace),
                localName,
                EwsUtilities.GetNamespaceUri(xmlNamespace));
        }

        /// <summary>
        /// Writes the end element.
        /// </summary>
        public void WriteEndElement()
        {
            this.xmlWriter.WriteEndElement();
        }

        /// <summary>
        /// Writes the attribute value.  Does not emit empty string values.
        /// </summary>
        /// <param name="localName">The local name of the attribute.</param>
        /// <param name="value">The value.</param>
        public void WriteAttributeValue(string localName, object value)
        {
            this.WriteAttributeValue(localName, false /* alwaysWriteEmptyString */, value);
        }

        /// <summary>
        /// Writes the attribute value.  Optionally emits empty string values.
        /// </summary>
        /// <param name="localName">The local name of the attribute.</param>
        /// <param name="alwaysWriteEmptyString">Always emit the empty string as the value.</param>
        /// <param name="value">The value.</param>
        public void WriteAttributeValue(string localName, bool alwaysWriteEmptyString, object value)
        {
            string stringValue;
            if (this.TryConvertObjectToString(value, out stringValue))
            {
                if ((stringValue != null) &&
                    (alwaysWriteEmptyString || (stringValue.Length != 0)))
                {
                    this.WriteAttributeString(localName, stringValue);
                }
            }
            else
            {
                throw new ServiceXmlSerializationException(
                            string.Format(Strings.AttributeValueCannotBeSerialized, value.GetType().Name, localName));
            }
        }

        /// <summary>
        /// Writes the attribute value.
        /// </summary>
        /// <param name="namespacePrefix">The namespace prefix.</param>
        /// <param name="localName">The local name of the attribute.</param>
        /// <param name="value">The value.</param>
        public void WriteAttributeValue(
            string namespacePrefix,
            string localName,
            object value)
        {
            string stringValue;
            if (this.TryConvertObjectToString(value, out stringValue))
            {
                if (!string.IsNullOrEmpty(stringValue))
                {
                    this.WriteAttributeString(
                        namespacePrefix,
                        localName,
                        stringValue);
                }
            }
            else
            {
                throw new ServiceXmlSerializationException(
                            string.Format(Strings.AttributeValueCannotBeSerialized, value.GetType().Name, localName));
            }
        }

        /// <summary>
        /// Writes the attribute value.
        /// </summary>
        /// <param name="localName">The local name of the attribute.</param>
        /// <param name="stringValue">The string value.</param>
        /// <exception cref="ServiceXmlSerializationException">Thrown if string value isn't valid for XML.</exception>
        internal void WriteAttributeString(string localName, string stringValue)
        {
            try
            {
                this.xmlWriter.WriteAttributeString(localName, stringValue);
            }
            catch (ArgumentException ex)
            {
                // XmlTextWriter will throw ArgumentException if string includes invalid characters.
                throw new ServiceXmlSerializationException(
                            string.Format(Strings.InvalidAttributeValue, stringValue, localName),
                            ex);
            }
        }

        /// <summary>
        /// Writes the attribute value.
        /// </summary>
        /// <param name="namespacePrefix">The namespace prefix.</param>
        /// <param name="localName">The local name of the attribute.</param>
        /// <param name="stringValue">The string value.</param>
        /// <exception cref="ServiceXmlSerializationException">Thrown if string value isn't valid for XML.</exception>
        internal void WriteAttributeString(
            string namespacePrefix,
            string localName,
            string stringValue)
        {
            try
            {
                this.xmlWriter.WriteAttributeString(
                                    namespacePrefix,
                                    localName,
                                    null,
                                    stringValue);
            }
            catch (ArgumentException ex)
            {
                // XmlTextWriter will throw ArgumentException if string includes invalid characters.
                throw new ServiceXmlSerializationException(
                            string.Format(Strings.InvalidAttributeValue, stringValue, localName),
                            ex);
            }
        }

        /// <summary>
        /// Writes string value. 
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="name">Element name (used for error handling)</param>
        /// <exception cref="ServiceXmlSerializationException">Thrown if string value isn't valid for XML.</exception>
        public void WriteValue(string value, string name)
        {
            try
            {
                this.xmlWriter.WriteValue(value);
            }
            catch (ArgumentException ex)
            {
                // XmlTextWriter will throw ArgumentException if string includes invalid characters.
                throw new ServiceXmlSerializationException(
                            string.Format(Strings.InvalidElementStringValue, value, name),
                            ex);
            }
        }

        /// <summary>
        /// Writes the element value.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localName">The local name of the element.</param>
        /// <param name="displayName">The name that should appear in the exception message when the value can not be serialized.</param>
        /// <param name="value">The value.</param>
        internal void WriteElementValue(XmlNamespace xmlNamespace, string localName, string displayName, object value)
        {
            string stringValue;
            if (this.TryConvertObjectToString(value, out stringValue))
            {
                //  PS # 205106: The code here used to check IsNullOrEmpty on stringValue instead of just null.
                //  Unfortunately, that meant that if someone really needed to update a string property to be the
                //  value "" (String.Empty), they couldn't do it, because we wouldn't emit the element here, causing
                //  an error on the server because an update is required to have a single sub-element that is the
                //  value to update.  So we need to allow an empty string to create an empty element (like <Value />).
                //  Note that changing this check to just check for null is fine, because the other types that get
                //  converted by TryConvertObjectToString() won't return an empty string if the conversion is
                //  successful (for instance, converting an integer to a string won't return an empty string - it'll
                //  always return the stringized integer).
                if (stringValue != null)
                {
                    this.WriteStartElement(xmlNamespace, localName);
                    this.WriteValue(stringValue, displayName);
                    this.WriteEndElement();
                }
            }
            else
            {
                throw new ServiceXmlSerializationException(
                        string.Format(Strings.ElementValueCannotBeSerialized, value.GetType().Name, localName));
            }
        }

        /// <summary>
        /// Writes the Xml Node
        /// </summary>
        /// <param name="xmlNode">The XML node.</param>
        public void WriteNode(XmlNode xmlNode)
        {
            if (xmlNode != null)
            {
                xmlNode.WriteTo(this.xmlWriter);
            }
        }

        /// <summary>
        /// Writes the element value.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localName">The local name of the element.</param>
        /// <param name="value">The value.</param>
        public void WriteElementValue(
            XmlNamespace xmlNamespace,
            string localName,
            object value)
        {
            this.WriteElementValue(xmlNamespace, localName, localName, value);
        }

        /// <summary>
        /// Writes the base64-encoded element value.
        /// </summary>
        /// <param name="buffer">The buffer.</param>
        public void WriteBase64ElementValue(byte[] buffer)
        {
            this.xmlWriter.WriteBase64(buffer, 0, buffer.Length);
        }

        /// <summary>
        /// Writes the base64-encoded element value.
        /// </summary>
        /// <param name="stream">The stream.</param>
        public void WriteBase64ElementValue(Stream stream)
        {
            byte[] buffer = new byte[BufferSize];
            int bytesRead;

            using (BinaryReader reader = new BinaryReader(stream))
            {
                do
                {
                    bytesRead = reader.Read(buffer, 0, BufferSize);

                    if (bytesRead > 0)
                    {
                        this.xmlWriter.WriteBase64(buffer, 0, bytesRead);
                    }
                }
                while (bytesRead > 0);
            }
        }

        /// <summary>
        /// Gets the internal XML writer.
        /// </summary>
        /// <value>The internal writer.</value>
        public XmlWriter InternalWriter
        {
            get { return this.xmlWriter; }
        }

        /// <summary>
        /// Gets the service.
        /// </summary>
        /// <value>The service.</value>
        public ExchangeServiceBase Service
        {
            get { return this.service; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the time zone SOAP header was emitted through this writer.
        /// </summary>
        /// <value>
        ///     <c>true</c> if the time zone SOAP header was emitted; otherwise, <c>false</c>.
        /// </value>
        public bool IsTimeZoneHeaderEmitted
        {
            get { return this.isTimeZoneHeaderEmitted; }
            set { this.isTimeZoneHeaderEmitted = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the SOAP message need WSSecurity Utility namespace.
        /// </summary>
        public bool RequireWSSecurityUtilityNamespace
        {
            get { return this.requireWSSecurityUtilityNamespace; }
            set { this.requireWSSecurityUtilityNamespace = value; }
        }
    }
}