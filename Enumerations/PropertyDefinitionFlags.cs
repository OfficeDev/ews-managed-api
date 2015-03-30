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