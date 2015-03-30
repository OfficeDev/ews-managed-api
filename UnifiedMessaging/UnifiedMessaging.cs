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