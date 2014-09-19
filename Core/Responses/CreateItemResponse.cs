// ---------------------------------------------------------------------------
// <copyright file="CreateItemResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CreateItemResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the response to an individual item creation operation.
    /// </summary>
    internal sealed class CreateItemResponse : CreateItemResponseBase
    {
        private Item item;

        /// <summary>
        /// Gets Item instance.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>Item.</returns>
        internal override Item GetObjectInstance(ExchangeService service, string xmlElementName)
        {
            return this.item;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CreateItemResponse"/> class.
        /// </summary>
        /// <param name="item">The item.</param>
        internal CreateItemResponse(Item item)
            : base()
        {
            this.item = item;
        }

        /// <summary>
        /// Clears the change log of the created folder if the creation succeeded.
        /// </summary>
        internal override void Loaded()
        {
            if (this.Result == ServiceResult.Success)
            {
                this.item.ClearChangeLog();
            }
        }
    }
}
