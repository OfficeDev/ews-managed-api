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