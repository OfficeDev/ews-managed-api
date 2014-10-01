#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

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
