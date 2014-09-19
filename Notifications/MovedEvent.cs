// ---------------------------------------------------------------------------
// <copyright file="MovedEvent.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MovedEvent class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents an event indicating that an item or folder was moved.
    /// </summary>
    public sealed class MovedEvent : MovedCopiedEvent
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MovedEvent"/> class.
        /// </summary>
        internal MovedEvent()
            : base()
        {
        }
    }
}
