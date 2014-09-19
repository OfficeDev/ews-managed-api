// ---------------------------------------------------------------------------
// <copyright file="Change.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the Change class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents a change as returned by a synchronization operation.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class Change
    {
        /// <summary>
        /// The type of change.
        /// </summary>
        private ChangeType changeType;

        /// <summary>
        /// The service object the change applies to.
        /// </summary>
        private ServiceObject serviceObject;

        /// <summary>
        /// The Id of the service object the change applies to.
        /// </summary>
        private ServiceId id;

        /// <summary>
        /// Initializes a new instance of Change.
        /// </summary>
        internal Change()
        {
        }

        /// <summary>
        /// Creates an Id of the appropriate class.
        /// </summary>
        /// <returns>A ServiceId.</returns>
        internal abstract ServiceId CreateId();

        /// <summary>
        /// Gets the type of the change.
        /// </summary>
        public ChangeType ChangeType
        {
            get { return this.changeType; }
            internal set { this.changeType = value; }
        }

        /// <summary>
        /// Gets or sets the service object the change applies to.
        /// </summary>
        internal ServiceObject ServiceObject
        {
            get { return this.serviceObject; }
            set { this.serviceObject = value; }
        }

        /// <summary>
        /// Gets or sets the Id of the service object the change applies to.
        /// </summary>
        internal ServiceId Id
        {
            get { return this.ServiceObject != null ? this.ServiceObject.GetId() : this.id; }
            set { this.id = value; }
        }
    }
}
