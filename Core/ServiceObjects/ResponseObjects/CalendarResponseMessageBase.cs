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
// <summary>Defines the CalendarResponseMessageBase class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents the base class for all calendar-related response messages.
    /// </summary>
    /// <typeparam name="TMessage">The type of message that is created when this response message is saved.</typeparam>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class CalendarResponseMessageBase<TMessage> : ResponseObject<TMessage>
        where TMessage : EmailMessage
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CalendarResponseMessageBase&lt;TMessage&gt;"/> class.
        /// </summary>
        /// <param name="referenceItem">The reference item.</param>
        internal CalendarResponseMessageBase(Item referenceItem)
            : base(referenceItem)
        {
        }

        /// <summary>
        /// Saves the response in the specified folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="destinationFolderId">The Id of the folder in which to save the response.</param>
        /// <returns>
        /// A CalendarActionResults object containing the various items that were created or modified as a
        /// results of this operation.
        /// </returns>
        public new CalendarActionResults Save(FolderId destinationFolderId)
        {
            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");

            return new CalendarActionResults(this.InternalCreate(destinationFolderId, MessageDisposition.SaveOnly));
        }

        /// <summary>
        /// Saves the response in the specified folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="destinationFolderName">The name of the folder in which to save the response.</param>
        /// <returns>
        /// A CalendarActionResults object containing the various items that were created or modified as a
        /// results of this operation.
        /// </returns>
        public new CalendarActionResults Save(WellKnownFolderName destinationFolderName)
        {
            return new CalendarActionResults(this.InternalCreate(new FolderId(destinationFolderName), MessageDisposition.SaveOnly));
        }

        /// <summary>
        /// Saves the response in the Drafts folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <returns>
        /// A CalendarActionResults object containing the various items that were created or modified as a
        /// results of this operation.
        /// </returns>
        public new CalendarActionResults Save()
        {
            return new CalendarActionResults(this.InternalCreate(null, MessageDisposition.SaveOnly));
        }

        /// <summary>
        /// Sends this response without saving a copy. Calling this method results in a call to EWS.
        /// </summary>
        /// <returns>
        /// A CalendarActionResults object containing the various items that were created or modified as a
        /// results of this operation.
        /// </returns>
        public new CalendarActionResults Send()
        {
            return new CalendarActionResults(this.InternalCreate(null, MessageDisposition.SendOnly));
        }

        /// <summary>
        /// Sends this response ans saves a copy in the specified folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="destinationFolderId">The Id of the folder in which to save the copy of the message.</param>
        /// <returns>
        /// A CalendarActionResults object containing the various items that were created or modified as a
        /// results of this operation.
        /// </returns>
        public new CalendarActionResults SendAndSaveCopy(FolderId destinationFolderId)
        {
            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");

            return new CalendarActionResults(this.InternalCreate(destinationFolderId, MessageDisposition.SendAndSaveCopy));
        }

        /// <summary>
        /// Sends this response and saves a copy in the specified folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="destinationFolderName">The name of the folder in which to save the copy of the message.</param>
        /// <returns>
        /// A CalendarActionResults object containing the various items that were created or modified as a
        /// results of this operation.
        /// </returns>
        public new CalendarActionResults SendAndSaveCopy(WellKnownFolderName destinationFolderName)
        {
            return new CalendarActionResults(this.InternalCreate(new FolderId(destinationFolderName), MessageDisposition.SendAndSaveCopy));
        }

        /// <summary>
        /// Sends this response ans saves a copy in the Sent Items folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <returns>
        /// A CalendarActionResults object containing the various items that were created or modified as a
        /// results of this operation.
        /// </returns>
        public new CalendarActionResults SendAndSaveCopy()
        {
            return new CalendarActionResults(this.InternalCreate(null, MessageDisposition.SendAndSaveCopy));
        }
    }
}
