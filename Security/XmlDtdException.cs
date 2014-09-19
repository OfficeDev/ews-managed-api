//---------------------------------------------------------------------
// <copyright file="XmlDtdException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Data
{
    using System.Xml;

    /// <summary>
    /// Exception class for banned xml parsing
    /// </summary>
    internal class XmlDtdException : XmlException
    {
        /// <summary>
        /// Gets the xml exception message.
        /// </summary>
        public override string Message
        {
            get { return "For security reasons DTD is prohibited in this XML document."; }
        }
    }
}
