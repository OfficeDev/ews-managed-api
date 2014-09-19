// ---------------------------------------------------------------------------
// <copyright file="SoapFaultDetails.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SoapFaultDetails class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Net;
    using System.Reflection;
    using System.Text;
    using System.Web.Services.Protocols;
    using System.Xml;

    /// <summary>
    /// Represents SoapFault details.
    /// </summary>
    internal class SoapFaultDetails
    {
        #region Private Members

        private string faultCode;
        private string faultString;
        private string faultActor;

        /// <summary>
        /// Response code returned by EWS requests.
        /// Default to InternalServerError.
        /// </summary>
        private ServiceError responseCode = ServiceError.ErrorInternalServerError;

        /// <summary>
        /// Message text of the error.
        /// </summary>
        private string message;

        /// <summary>
        /// This is returned by Availability requests.
        /// </summary>
        private ServiceError errorCode = ServiceError.NoError;

        /// <summary>
        /// This is returned by UM requests. It's the name of the exception that was raised.
        /// </summary>
        private string exceptionType;

        /// <summary>
        /// When a schema validation error is returned, this is the line number in the request where the error occurred.
        /// </summary>
        private int lineNumber;

        /// <summary>
        /// When a schema validation error is returned, this is the offset into the line of the request where the error occurred.
        /// </summary>
        private int positionWithinLine;

        /// <summary>
        /// Dictionary of key/value pairs from the MessageXml node in the fault. Usually empty but there are 
        /// a few cases where SOAP faults may include MessageXml details (e.g. CASOverBudgetException includes
        /// BackoffTime value).
        /// </summary>
        private Dictionary<string, string> errorDetails = new Dictionary<string, string>();

        #endregion

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="SoapFaultDetails"/> class.
        /// </summary>
        private SoapFaultDetails()
        {
        }

        #endregion

        /// <summary>
        /// Parses the soap:Fault content.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="soapNamespace">The SOAP namespace to use.</param>
        /// <returns>SOAP fault details.</returns>
        internal static SoapFaultDetails Parse(EwsXmlReader reader, XmlNamespace soapNamespace)
        {
            SoapFaultDetails soapFaultDetails = new SoapFaultDetails();

            do
            {
                reader.Read();
                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.SOAPFaultCodeElementName:
                            soapFaultDetails.FaultCode = reader.ReadElementValue();
                            break;

                        case XmlElementNames.SOAPFaultStringElementName:
                            soapFaultDetails.FaultString = reader.ReadElementValue();
                            break;

                        case XmlElementNames.SOAPFaultActorElementName:
                            soapFaultDetails.FaultActor = reader.ReadElementValue();
                            break;

                        case XmlElementNames.SOAPDetailElementName:
                            soapFaultDetails.ParseDetailNode(reader);
                            break;

                        default:
                            break;
                    }
                }
            }
            while (!reader.IsEndElement(soapNamespace, XmlElementNames.SOAPFaultElementName));

            return soapFaultDetails;
        }

        /// <summary>
        /// Parses the specified json object.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <returns></returns>
        internal static SoapFaultDetails Parse(JsonObject jsonObject)
        {
            SoapFaultDetails soapFaultDetails = new SoapFaultDetails();

            foreach (string key in jsonObject.Keys)
            {
                switch (key)
                {
                    // TODO: Fix up?
                    case "FaultMessage":
                        soapFaultDetails.FaultString = jsonObject.ReadAsString(key);
                        break;
                    default:
                        break;
                }
            }

            return soapFaultDetails;
        }

        /// <summary>
        /// Parses the detail node.
        /// </summary>
        /// <param name="reader">The reader.</param>
        private void ParseDetailNode(EwsXmlReader reader)
        {
            do
            {
                reader.Read();
                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.EwsResponseCodeElementName:
                            try
                            {
                                this.ResponseCode = reader.ReadElementValue<ServiceError>();
                            }
                            catch (ArgumentException)
                            {
                                // ServiceError couldn't be mapped to enum value, treat as an ISE
                                this.ResponseCode = ServiceError.ErrorInternalServerError;
                            }

                            break;

                        case XmlElementNames.EwsMessageElementName:
                            this.Message = reader.ReadElementValue();
                            break;

                        case XmlElementNames.EwsLineElementName:
                            this.LineNumber = reader.ReadElementValue<int>();
                            break;

                        case XmlElementNames.EwsPositionElementName:
                            this.PositionWithinLine = reader.ReadElementValue<int>();
                            break;

                        case XmlElementNames.EwsErrorCodeElementName:
                            try
                            {
                                this.ErrorCode = reader.ReadElementValue<ServiceError>();
                            }
                            catch (ArgumentException)
                            {
                                // ServiceError couldn't be mapped to enum value, treat as an ISE
                                this.ErrorCode = ServiceError.ErrorInternalServerError;
                            }

                            break;

                        case XmlElementNames.EwsExceptionTypeElementName:
                            this.ExceptionType = reader.ReadElementValue();
                            break;

                        case XmlElementNames.MessageXml:
                            this.ParseMessageXml(reader);
                            break;

                        default:
                            // Ignore any other details
                            break;
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.NotSpecified, XmlElementNames.SOAPDetailElementName));
        }

        /// <summary>
        /// Parses the message XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        private void ParseMessageXml(EwsXmlReader reader)
        {
            // E12 and E14 return the MessageXml element in different
            // namespaces (types namespace for E12, errors namespace in E14). To
            // avoid this problem, the parser will match the namespace from the
            // start and end elements.
            XmlNamespace elementNS = EwsUtilities.GetNamespaceFromUri(reader.NamespaceUri);

            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.IsStartElement() && !reader.IsEmptyElement)
                    {
                        switch (reader.LocalName)
                        {
                            case XmlElementNames.Value:
                                this.errorDetails.Add(
                                    reader.ReadAttributeValue(XmlAttributeNames.Name),
                                    reader.ReadElementValue());
                                break;

                            default:
                                break;
                        }
                    }
                }
                while (!reader.IsEndElement(elementNS, XmlElementNames.MessageXml));
            }
        }

        /// <summary>
        /// Gets or sets the SOAP fault code.
        /// </summary>
        /// <value>The SOAP fault code.</value>
        internal string FaultCode
        {
            get { return this.faultCode; }
            set { this.faultCode = value; }
        }

        /// <summary>
        /// Gets or sets the SOAP fault string.
        /// </summary>
        /// <value>The fault string.</value>
        internal string FaultString
        {
            get { return this.faultString; }
            set { this.faultString = value; }
        }

        /// <summary>
        /// Gets or sets the SOAP fault actor.
        /// </summary>
        /// <value>The fault actor.</value>
        internal string FaultActor
        {
            get { return this.faultActor; }
            set { this.faultActor = value; }
        }

        /// <summary>
        /// Gets or sets the response code.
        /// </summary>
        /// <value>The response code.</value>
        internal ServiceError ResponseCode
        {
            get { return this.responseCode; }
            set { this.responseCode = value; }
        }

        /// <summary>
        /// Gets or sets the message.
        /// </summary>
        /// <value>The message.</value>
        internal string Message
        {
            get { return this.message; }
            set { this.message = value; }
        }

        /// <summary>
        /// Gets or sets the error code.
        /// </summary>
        /// <value>The error code.</value>
        internal ServiceError ErrorCode
        {
            get { return this.errorCode; }
            set { this.errorCode = value; }
        }

        /// <summary>
        /// Gets or sets the type of the exception.
        /// </summary>
        /// <value>The type of the exception.</value>
        internal string ExceptionType
        {
            get { return this.exceptionType; }
            set { this.exceptionType = value; }
        }

        /// <summary>
        /// Gets or sets the line number.
        /// </summary>
        /// <value>The line number.</value>
        internal int LineNumber
        {
            get { return this.lineNumber; }
            set { this.lineNumber = value; }
        }

        /// <summary>
        /// Gets or sets the position within line.
        /// </summary>
        /// <value>The position within line.</value>
        internal int PositionWithinLine
        {
            get { return this.positionWithinLine; }
            set { this.positionWithinLine = value; }
        }

        /// <summary>
        /// Gets or sets the error details dictionary.
        /// </summary>
        /// <value>The error details dictionary.</value>
        internal Dictionary<string, string> ErrorDetails
        {
            get { return this.errorDetails; }
            set { this.errorDetails = value; }
        }
    }
}
