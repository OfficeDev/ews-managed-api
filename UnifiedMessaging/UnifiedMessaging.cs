// ---------------------------------------------------------------------------
// <copyright file="UnifiedMessaging.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the UnifiedMessaging class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the Unified Messaging functionalities.
    /// </summary>
    public sealed class UnifiedMessaging
    {
        private ExchangeService service;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="service">EWS service to which this object belongs.</param>
        internal UnifiedMessaging(ExchangeService service)
        {
            this.service = service;
        }

        /// <summary>
        /// Calls a phone and reads a message to the person who picks up.
        /// </summary>
        /// <param name="itemId">The Id of the message to read.</param>
        /// <param name="dialString">The full dial string used to call the phone.</param>
        /// <returns>An object providing status for the phone call.</returns>
        public PhoneCall PlayOnPhone(ItemId itemId, string dialString)
        {
            EwsUtilities.ValidateParam(itemId, "itemId");
            EwsUtilities.ValidateParam(dialString, "dialString");

            PlayOnPhoneRequest request = new PlayOnPhoneRequest(service);
            request.DialString = dialString;
            request.ItemId = itemId;
            PlayOnPhoneResponse serviceResponse = request.Execute();

            PhoneCall callInformation = new PhoneCall(service, serviceResponse.PhoneCallId);

            return callInformation;
        }

        /// <summary>
        /// Retrieves information about a current phone call.
        /// </summary>
        /// <param name="id">The Id of the phone call.</param>
        /// <returns>An object providing status for the phone call.</returns>
        internal PhoneCall GetPhoneCallInformation(PhoneCallId id)
        {
            GetPhoneCallRequest request = new GetPhoneCallRequest(service);
            request.Id = id;
            GetPhoneCallResponse response = request.Execute();

            return response.PhoneCall;
        }

        /// <summary>
        /// Disconnects a phone call.
        /// </summary>
        /// <param name="id">The Id of the phone call.</param>
        internal void DisconnectPhoneCall(PhoneCallId id)
        {
            DisconnectPhoneCallRequest request = new DisconnectPhoneCallRequest(service);
            request.Id = id;
            request.Execute();
        }
    }
}
