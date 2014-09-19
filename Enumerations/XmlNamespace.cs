// ---------------------------------------------------------------------------
// <copyright file="XmlNamespace.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the xMLNamespace enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the namespaces as used by the EwsXmlReader, EwsServiceXmlReader, and EwsServiceXmlWriter classes.
    /// </summary>
    internal enum XmlNamespace
    {
        /// <summary>
        /// The namespace is not specified.
        /// </summary>
        NotSpecified,

        /// <summary>
        /// The EWS Messages namespace.
        /// </summary>
        Messages,

        /// <summary>
        /// The EWS Types namespace.
        /// </summary>
        Types,

        /// <summary>
        /// The EWS Errors namespace.
        /// </summary>
        Errors,

        /// <summary>
        /// The SOAP 1.1 namespace.
        /// </summary>
        Soap,

        /// <summary>
        /// The SOAP 1.2 namespace.
        /// </summary>
        Soap12,

        /// <summary>
        /// XmlSchema-Instance namespace.
        /// </summary>
        XmlSchemaInstance,

        /// <summary>
        /// The Passport SOAP services SOAP fault namespace.
        /// </summary>
        PassportSoapFault,

        /// <summary>
        /// The WS-Trust February 2005 namespace.
        /// </summary>
        WSTrustFebruary2005,

        /// <summary>
        /// The WS Addressing 1.0 namespace.
        /// </summary>
        WSAddressing,

        /// <summary>
        /// The Autodiscover SOAP service namespace.
        /// </summary>
        Autodiscover,
    }
}
