// ---------------------------------------------------------------------------
// <copyright file="PropertyDefinitionFlags.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PropertyDefinitionFlags enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines how a complex property behaves.
    /// </summary>
    [Flags]
    internal enum PropertyDefinitionFlags
    {
        /// <summary>
        /// No specific behavior.
        /// </summary>
        None = 0,

        /// <summary>
        /// The property is automatically instantiated when it is read.
        /// </summary>
        AutoInstantiateOnRead = 1,

        /// <summary>
        /// The existing instance of the property is reusable. 
        /// </summary>
        ReuseInstance = 2,

        /// <summary>
        /// The property can be set.
        /// </summary>
        CanSet = 4,

        /// <summary>
        /// The property can be updated.
        /// </summary>
        CanUpdate = 8,

        /// <summary>
        /// The property can be deleted.
        /// </summary>
        CanDelete = 16,

        /// <summary>
        /// The property can be searched.
        /// </summary>
        CanFind = 32,

        /// <summary>
        /// The property must be loaded explicitly
        /// </summary>
        MustBeExplicitlyLoaded = 64,

        /// <summary>
        /// Only meaningful for "collection" property. With this flag, the item in the collection gets updated, 
        /// instead of creating and adding new items to the collection.
        /// Should be used together with the ReuseInstance flag.
        /// </summary>
        UpdateCollectionItems = 128,
    }
}
