// ---------------------------------------------------------------------------
// <copyright file="EwsServiceJsonReader.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the EwsServiceJsonReader class.</summary>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// JSON reader.
    /// </summary>
    internal class EwsServiceJsonReader
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="EwsServiceJsonReader"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal EwsServiceJsonReader(ExchangeService service)
        {
            this.Service = service;
        }

        /// <summary>
        /// Reads the service objects collection from JSON.
        /// </summary>
        /// <typeparam name="TServiceObject">The type of the service object.</typeparam>
        /// <param name="jsonResponse">The json response.</param>
        /// <param name="collectionJsonElementName">Name of the collection XML element.</param>
        /// <param name="getObjectInstanceDelegate">The get object instance delegate.</param>
        /// <param name="clearPropertyBag">if set to <c>true</c> [clear property bag].</param>
        /// <param name="requestedPropertySet">The requested property set.</param>
        /// <param name="summaryPropertiesOnly">if set to <c>true</c> [summary properties only].</param>
        /// <returns>List of service objects.</returns>
        internal List<TServiceObject> ReadServiceObjectsCollectionFromJson<TServiceObject>(
            JsonObject jsonResponse,
            string collectionJsonElementName,
            GetObjectInstanceDelegate<TServiceObject> getObjectInstanceDelegate,
            bool clearPropertyBag,
            PropertySet requestedPropertySet,
            bool summaryPropertiesOnly) where TServiceObject : ServiceObject
        {
            List<TServiceObject> serviceObjects = new List<TServiceObject>();
            TServiceObject serviceObject = null;

            object[] jsonServiceObjects = jsonResponse.ReadAsArray(collectionJsonElementName);
            foreach (object arrayEntry in jsonServiceObjects)
            {
                JsonObject jsonServiceObject = arrayEntry as JsonObject;

                if (jsonServiceObject != null)
                {
                    serviceObject = getObjectInstanceDelegate(this.Service, jsonServiceObject.ReadTypeString());

                    if (serviceObject != null)
                    {
                        if (string.Compare(jsonServiceObject.ReadTypeString(), serviceObject.GetXmlElementName(), StringComparison.Ordinal) != 0)
                        {
                            throw new ServiceLocalException(
                                string.Format(
                                    "The type of the object in the store ({0}) does not match that of the local object ({1}).",
                                    jsonServiceObject.ReadTypeString(),
                                    serviceObject.GetXmlElementName()));
                        }

                        serviceObject.LoadFromJson(
                                        jsonServiceObject,
                                        this.Service,
                                        clearPropertyBag,
                                        requestedPropertySet,
                                        summaryPropertiesOnly);

                        serviceObjects.Add(serviceObject);
                    }
                }
            }

            return serviceObjects;
        }

        /// <summary>
        /// Gets or sets the service.
        /// </summary>
        internal ExchangeService Service { get; set; }
    }
}
