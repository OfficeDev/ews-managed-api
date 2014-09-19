// ---------------------------------------------------------------------------
// <copyright file="SuppressReadReceipt.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SuppressReadReceipt class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a response object created to supress read receipts for an item.
    /// </summary>
    [ServiceObjectDefinition(XmlElementNames.SuppressReadReceipt, ReturnedByServer = false)]
    internal sealed class SuppressReadReceipt : ServiceObject
    {
        private Item referenceItem;

        /// <summary>
        /// Initializes a new instance of the <see cref="SuppressReadReceipt"/> class.
        /// </summary>
        /// <param name="referenceItem">The reference item.</param>
        internal SuppressReadReceipt(Item referenceItem)
            : base(referenceItem.Service)
        {
            EwsUtilities.Assert(
                referenceItem != null,
                "SuppressReadReceipt.ctor",
                "referenceItem is null");

            referenceItem.ThrowIfThisIsNew();

            this.referenceItem = referenceItem;
        }

        /// <summary>
        /// Internal method to return the schema associated with this type of object.
        /// </summary>
        /// <returns>The schema associated with this type of object.</returns>
        internal override ServiceObjectSchema GetSchema()
        {
            return ResponseObjectSchema.Instance;
        }

        /// <summary>
        /// Gets the minimum required server version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2007_SP1;
        }

        /// <summary>
        /// Loads the specified set of properties on the object.
        /// </summary>
        /// <param name="propertySet">The properties to load.</param>
        internal override void InternalLoad(PropertySet propertySet)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Deletes the object.
        /// </summary>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="sendCancellationsMode">Indicates whether meeting cancellation messages should be sent.</param>
        /// <param name="affectedTaskOccurrences">Indicate which occurrence of a recurring task should be deleted.</param>
        internal override void InternalDelete(
            DeleteMode deleteMode,
            SendCancellationsMode? sendCancellationsMode,
            AffectedTaskOccurrence? affectedTaskOccurrences)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Create the response object.
        /// </summary>
        /// <param name="parentFolderId">The parent folder id.</param>
        /// <param name="messageDisposition">The message disposition.</param>
        internal void InternalCreate(FolderId parentFolderId, MessageDisposition? messageDisposition)
        {
            ((ItemId)this.PropertyBag[ResponseObjectSchema.ReferenceItemId]).Assign(this.referenceItem.Id);

            this.Service.InternalCreateResponseObject(
                this,
                parentFolderId,
                messageDisposition);
        }
    }
}
