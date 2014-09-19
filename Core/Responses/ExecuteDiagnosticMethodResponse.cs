// ---------------------------------------------------------------------------
// <copyright file="ExecuteDiagnosticMethodResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ExecuteDiagnosticMethodResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Represents the response to a ExecuteDiagnosticMethod operation
    /// </summary>
    internal sealed class ExecuteDiagnosticMethodResponse : ServiceResponse
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExecuteDiagnosticMethodResponse"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal ExecuteDiagnosticMethodResponse(ExchangeService service)
            : base()
        {
            EwsUtilities.Assert(
                service != null,
                "ExecuteDiagnosticMethodResponse.ctor",
                "service is null");
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.ReturnValue);

            using (XmlReader returnValueReader = reader.GetXmlReaderForNode())
            {
                this.ReturnValue = new SafeXmlDocument();
                this.ReturnValue.Load(returnValueReader);
            }

            reader.SkipCurrentElement();
            reader.ReadEndElementIfNecessary(XmlNamespace.Messages, XmlElementNames.ReturnValue);
        }

        /// <summary>
        /// Gets the return value.
        /// </summary>
        /// <value>The return value.</value>
        internal XmlDocument ReturnValue
        {
            get;
            private set;
        }
    }
}
