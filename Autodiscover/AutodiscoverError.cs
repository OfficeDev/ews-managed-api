// ---------------------------------------------------------------------------
// <copyright file="AutodiscoverError.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AutodiscoverError class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents an error returned by the Autodiscover service.
    /// </summary>
    [Serializable]
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class AutodiscoverError
    {
        private string time;
        private string id;
        private int errorCode;
        private string message;
        private string debugData;

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverError"/> class.
        /// </summary>
        private AutodiscoverError()
        {
        }

        /// <summary>
        /// Parses the XML through the specified reader and creates an Autodiscover error.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>An Autodiscover error.</returns>
        internal static AutodiscoverError Parse(EwsXmlReader reader)
        {
            AutodiscoverError error = new AutodiscoverError();

            error.time = reader.ReadAttributeValue(XmlAttributeNames.Time);
            error.id = reader.ReadAttributeValue(XmlAttributeNames.Id);

            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.ErrorCode:
                            error.errorCode = reader.ReadElementValue<int>();
                            break;
                        case XmlElementNames.Message:
                            error.message = reader.ReadElementValue();
                            break;
                        case XmlElementNames.DebugData:
                            error.debugData = reader.ReadElementValue();
                            break;
                        default:
                            reader.SkipCurrentElement();
                            break;
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.NotSpecified, XmlElementNames.Error));

            return error;
        }

        /// <summary>
        /// Gets the time when the error was returned.
        /// </summary>
        public string Time
        {
            get { return this.time; }
        }

        /// <summary>
        /// Gets a hash of the name of the computer that is running Microsoft Exchange Server that has the Client Access server role installed.
        /// </summary>
        public string Id
        {
            get { return this.id; }
        }

        /// <summary>
        /// Gets the error code.
        /// </summary>
        public int ErrorCode
        {
            get { return this.errorCode; }
        }

        /// <summary>
        /// Gets the error message.
        /// </summary>
        public string Message
        {
            get { return this.message; }
        }

        /// <summary>
        /// Gets the debug data.
        /// </summary>
        public string DebugData
        {
            get { return this.debugData; }
        }
    }
}
