// ---------------------------------------------------------------------------
// <copyright file="DelegateTypes.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines miscellaneous delegate types.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Net;
    using System.Xml;

    /// <summary>
    /// Defines a delegate that is used to allow applications to emit custom XML when SOAP requests are sent to Exchange.
    /// </summary>
    /// <param name="writer">The XmlWriter to use to emit the custom XML.</param>
    public delegate void CustomXmlSerializationDelegate(XmlWriter writer);

    /// <summary>
    /// Delegate method to handle capturing http response headers.
    /// </summary>
    /// <param name="responseHeaders">Http response headers.</param>
    public delegate void ResponseHeadersCapturedHandler(WebHeaderCollection responseHeaders);

    /// <summary>
    /// Defines a delegate used to notify that a service object has been modified.
    /// </summary>
    /// <param name="serviceObject">The service object that has been modified.</param>
    internal delegate void ServiceObjectChangedDelegate(ServiceObject serviceObject);

    /// <summary>
    /// Indicates that a complex property changed.
    /// </summary>
    /// <param name="complexProperty">Complex property.</param>
    internal delegate void ComplexPropertyChangedDelegate(ComplexProperty complexProperty);

    /// <summary>
    /// Indicates that a property bag changed.
    /// </summary>
    internal delegate void PropertyBagChangedDelegate();

    /// <summary>
    /// Used to produce an instance of a service object based on XML element name.
    /// </summary>
    /// <typeparam name="T">ServiceObject type.</typeparam>
    /// <param name="service">Exchange service instance.</param>
    /// <param name="xmlElementName">XML element name.</param>
    /// <returns>Service object instance.</returns>
    internal delegate T GetObjectInstanceDelegate<T>(ExchangeService service, string xmlElementName) where T : ServiceObject;
}
