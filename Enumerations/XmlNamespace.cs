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