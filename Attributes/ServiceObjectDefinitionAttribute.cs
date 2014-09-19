// ---------------------------------------------------------------------------
// <copyright file="ServiceObjectDefinitionAttribute.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServiceObjectDefinitionAttribute attribute</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// ServiceObjectDefinition attribute decorates classes that map to EWS service objects.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    internal sealed class ServiceObjectDefinitionAttribute : Attribute
    {
        private string xmlElementName;
        private bool returnedByServer;

        /// <summary>
        /// Initializes a new instance of the <see cref="ServiceObjectDefinitionAttribute"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal ServiceObjectDefinitionAttribute(string xmlElementName)
            : base()
        {
            this.xmlElementName = xmlElementName;
            this.returnedByServer = true;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <value>The name of the XML element.</value>
        internal string XmlElementName
        {
            get { return this.xmlElementName; }
        }

        /// <summary>
        /// True if this ServiceObject can be returned by the server as an object, false otherwise.
        /// </summary>
        public bool ReturnedByServer
        {
            get { return this.returnedByServer; }
            set { this.returnedByServer = value; }
        }
    }
}
