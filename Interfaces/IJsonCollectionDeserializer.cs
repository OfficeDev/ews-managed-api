// ---------------------------------------------------------------------------
// <copyright file="IJsonCollectionDeserializer.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Interface for Complex Properties that load from a JSON collection.
    /// </summary>
    internal interface IJsonCollectionDeserializer
    {
        /// <summary>
        /// Loads from json collection to create a new collection item.
        /// </summary>
        /// <param name="jsonCollection">The json collection.</param>
        /// <param name="service">The service.</param>
        void CreateFromJsonCollection(object[] jsonCollection, ExchangeService service);

        /// <summary>
        /// Loads from json collection to update the existing collection item.
        /// </summary>
        /// <param name="jsonCollection">The json collection.</param>
        /// <param name="service">The service.</param>
        void UpdateFromJsonCollection(object[] jsonCollection, ExchangeService service);
    }
}
